using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using ExcelDna.Integration;
using ExcelModelingToolkit.Core.Formatting;
using ExcelModelingToolkit.Core.Undo;

namespace ExcelModelingToolkit.AddIn;

/// <summary>
/// Runs a formatting cycle on the selection: a thin COM adapter over <see cref="CycleEngine"/>. Reads the active
/// cell, captures the selection's current values for undo (<see cref="SnapshotPlanner"/>), writes the whole
/// selection with one COM property write, then reads the active cell back and records the undo snapshot. Call in
/// macro context on the main thread. Any COM write clears Excel's undo history (spike K2c); our own undo stack
/// (<see cref="Session.Undo"/>, <see cref="UndoCommand"/>) replaces it for our changes.
/// </summary>
internal static class CycleCommand
{
    /// <summary>
    /// Applies the next item of the cycle whose id is <paramref name="actionId"/>, shows the result in the status bar
    /// and writes one diagnostics log line. <paramref name="source"/> is how it was invoked (the key, or <c>ribbon</c>).
    /// Never throws.
    /// </summary>
    public static void Run(string actionId, string source)
    {
        var stopwatch = Stopwatch.StartNew();
        var trace = new Trace();
        string result;
        try
        {
            result = Execute(actionId, stopwatch, trace);
        }
        catch (Exception ex)
        {
            result = "error: " + ex.Message;
            StatusBar.Show($"{trace.Label ?? actionId}: failed: {ex.Message}");
        }

        // ms: entry to the end of the COM write (including the undo capture). totalMs: entry to here (adds the
        // read-back, recording the undo snapshot and the status bar).
        var total = stopwatch.Elapsed.TotalMilliseconds;
        var elapsed = trace.ElapsedMs ?? total;
        DiagnosticsLog.Write(
            actionId,
            "key=" + source,
            "ms=" + elapsed.ToString("0.0", CultureInfo.InvariantCulture),
            "totalMs=" + total.ToString("0.0", CultureInfo.InvariantCulture),
            "cells=" + trace.Cells.ToString(CultureInfo.InvariantCulture),
            "item=" + trace.Item,
            result);
    }

    private static string Execute(string actionId, Stopwatch stopwatch, Trace trace)
    {
        var cycle = Session.Settings.FindCycle(actionId);
        if (cycle is null || cycle.Items.Count == 0)
        {
            StatusBar.Show($"{ProductInfo.Name}: the settings have no cycle '{actionId}'.");
            return "no cycle";
        }

        var label = cycle.DisplayName;
        trace.Label = label;

        dynamic app = ExcelDnaUtil.Application;
        object? workbook = app.ActiveWorkbook;
        if (workbook is null)
        {
            StatusBar.Show($"{label}: No editable workbook is active (Protected View?).");
            return "no workbook";
        }

        if (!SelectionIsCells())
        {
            StatusBar.Show($"{label}: select cells first.");
            return "selection is not cells";
        }

        object activeSheet = app.ActiveSheet;
        if (CellFormats.FormattingIsProtected(activeSheet))
        {
            StatusBar.Show($"{label}: The sheet is protected; formatting is not allowed.");
            return "sheet protected";
        }

        // COM objects are held as object so only the adapter helpers bind late.
        object selection = app.Selection;
        object activeCell = app.ActiveCell;
        trace.Cells = CellCount(selection);

        var current = CellFormats.Read(activeCell, cycle.Kind);
        Session.States.TryGetValue(cycle.Id, out var previous);
        var step = Session.Engine.Next(cycle, current, SelectionKey(selection), previous);
        trace.Item = $"{step.Index + 1}/{cycle.Items.Count} {step.Item.Name}";

        var capture = CaptureForUndo(actionId, activeSheet, selection, cycle.Kind);
        try
        {
            CellFormats.Write(selection, cycle.Kind, step.Item.Value);
        }
        catch (Exception ex) when (cycle.Kind == CycleKind.NumberFormat)
        {
            StatusBar.Show(
                $"{label}: Excel refused \"{step.Item.Name}\". The workbook may have too many number formats, " +
                $"or the sheet may be protected. ({ex.Message})");
            return "error: " + ex.Message;
        }

        trace.ElapsedMs = stopwatch.Elapsed.TotalMilliseconds;

        // The write is done: record it for undo first, so nothing below can skip that. Undo checks that the cells
        // still hold what Excel reports, so the read-back form is recorded.
        var readBack = ReadBack(activeCell, cycle.Kind);
        var undoNote = RecordUndo(actionId, cycle, activeSheet, capture, readBack);

        // Excel may normalize a number format code or map a color to another; remember what it reports so the
        // cell is recognized next time (and, for number formats, learn the alias).
        var state = Session.Engine.RecordReadBack(cycle, step.State, readBack);
        if (!readBack.IsUnknown && readBack != step.Item.Value)
        {
            DiagnosticsLog.Write(
                cycle.Kind == CycleKind.NumberFormat ? "AliasLearned" : "ColorReadBackDiffers",
                cycle.Id,
                $"{step.Item.Value} -> {readBack}");
        }

        Session.States[cycle.Id] = state;
        StatusBar.Show($"{label}: {step.Item.Name} ({step.Index + 1}/{cycle.Items.Count}){undoNote}");
        return "ok " + step.Reason;
    }

    /// <summary>
    /// Captures the selection's current values of <paramref name="kind"/>, before the write, and logs the blocks,
    /// reads and time taken. Never throws: a failure gives an unavailable plan with the reason.
    /// </summary>
    private static SnapshotPlan CaptureForUndo(string actionId, object activeSheet, object selection, CycleKind kind)
    {
        var stopwatch = Stopwatch.StartNew();
        SnapshotPlan plan;
        try
        {
            plan = SnapshotPlanner.Plan(
                Areas(selection),
                UsedRange(activeSheet),
                Session.Settings.UndoCellCap,
                new SheetFormatReader(activeSheet, kind));
        }
        catch (Exception ex)
        {
            plan = SnapshotPlan.Unavailable("could not read the current formats: " + ex.Message, 0);
        }

        DiagnosticsLog.Write(
            "UndoCapture",
            actionId,
            "blocks=" + plan.Blocks.Count.ToString(CultureInfo.InvariantCulture),
            "reads=" + plan.Reads.ToString(CultureInfo.InvariantCulture),
            "cells=" + plan.CellCount.ToString(CultureInfo.InvariantCulture),
            "ms=" + stopwatch.Elapsed.TotalMilliseconds.ToString("0.0", CultureInfo.InvariantCulture),
            plan.IsAvailable ? "ok" : "unavailable: " + plan.UnavailableReason);
        return plan;
    }

    /// <summary>
    /// Pushes the undo snapshot for a completed write (every block now holds <paramref name="applied"/>, the
    /// active cell's read-back). If the change cannot be recorded (the capture failed, the read-back is unknown,
    /// or the snapshot could not be made), pushes a barrier instead, so Ctrl+Z stops there rather than reach past
    /// this change to older ones. Either way our redo stack is cleared. Returns the status-bar suffix: empty, or
    /// why this change cannot be undone. Never throws.
    /// </summary>
    private static string RecordUndo(string actionId, CycleDefinition cycle, object activeSheet, SnapshotPlan capture, CycleValue applied)
    {
        var workbook = "(unknown workbook)";
        var sheetName = "(unknown sheet)";
        string? reason = null;
        try
        {
            dynamic sheet = activeSheet;
            string? name = sheet.Name;
            string? fullName = sheet.Parent.FullName;
            if (name is null || fullName is null)
            {
                reason = "could not identify the sheet";
            }
            else
            {
                sheetName = name;
                workbook = fullName;
            }
        }
        catch (Exception ex)
        {
            reason = "could not identify the sheet: " + ex.Message;
        }

        reason ??= capture.UnavailableReason ?? (applied.IsUnknown ? "the applied format could not be read back" : null);
        if (reason is null)
        {
            try
            {
                Session.Undo.Push(FormatSnapshot.Create(
                    cycle.DisplayName,
                    cycle.Kind,
                    workbook,
                    sheetName,
                    capture.Blocks.Select(b => b.WithApplied(applied))));
                return string.Empty;
            }
            catch (Exception ex)
            {
                reason = ex.Message;
            }
        }

        Session.Undo.Push(FormatSnapshot.Unavailable(cycle.DisplayName, cycle.Kind, workbook, sheetName, reason));
        DiagnosticsLog.Write("UndoCapture", actionId, "not recorded, barrier pushed: " + reason);
        return $" (undo unavailable: {reason})";
    }

    /// <summary>The selection's areas as rectangles.</summary>
    private static IReadOnlyList<CellRect> Areas(object selection)
    {
        dynamic range = selection;
        dynamic areas = range.Areas;
        int count = areas.Count;
        var rects = new List<CellRect>(count);
        for (var i = 1; i <= count; i++)
        {
            object area = areas.Item(i);
            rects.Add(ToRect(area));
        }

        return rects;
    }

    /// <summary>The worksheet's used range as a rectangle.</summary>
    private static CellRect UsedRange(object worksheet)
    {
        dynamic sheet = worksheet;
        object used = sheet.UsedRange;
        return ToRect(used);
    }

    private static CellRect ToRect(object range)
    {
        dynamic r = range;
        object row = r.Row;
        object column = r.Column;
        object rows = r.Rows.Count;
        object columns = r.Columns.Count;
        return new CellRect(
            Convert.ToInt32(row, CultureInfo.InvariantCulture),
            Convert.ToInt32(column, CultureInfo.InvariantCulture),
            Convert.ToInt32(rows, CultureInfo.InvariantCulture),
            Convert.ToInt32(columns, CultureInfo.InvariantCulture));
    }

    /// <summary>True if cells (not a chart, shape or other object) are selected. Uses the C API <c>SELECTION()</c>.</summary>
    private static bool SelectionIsCells()
    {
        try
        {
            return XlCall.Excel(XlCall.xlfSelection) is ExcelReference;
        }
        catch (XlCallException)
        {
            return false;
        }
    }

    /// <summary>The active cell's value right after a write; unknown if it cannot be read (the write still stands).</summary>
    private static CycleValue ReadBack(object activeCell, CycleKind kind)
    {
        try
        {
            return CellFormats.Read(activeCell, kind);
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("ReadBackFailed", kind.ToString(), ex.Message);
            return CycleValue.Unknown;
        }
    }

    private static long CellCount(object range)
    {
        dynamic r = range;
        object count = r.CountLarge;
        return Convert.ToInt64(count, CultureInfo.InvariantCulture);
    }

    /// <summary><c>workbook|sheet|address</c>; the address falls back to a cell count if Excel cannot produce it.</summary>
    private static string SelectionKey(object range)
    {
        dynamic selection = range;
        dynamic sheet = selection.Worksheet;
        string workbookName = sheet.Parent.Name;
        string sheetName = sheet.Name;
        string address;
        try
        {
            address = selection.Address;
        }
        catch (Exception)
        {
            object areas = selection.Areas.Count;
            address = string.Format(CultureInfo.InvariantCulture, "({0} areas, {1} cells)", areas, CellCount(range));
        }

        return workbookName + "|" + sheetName + "|" + address;
    }

    /// <summary>What one run did, for the diagnostics log.</summary>
    private sealed class Trace
    {
        public string? Label { get; set; }

        public long Cells { get; set; }

        public string Item { get; set; } = "-";

        public double? ElapsedMs { get; set; }
    }
}
