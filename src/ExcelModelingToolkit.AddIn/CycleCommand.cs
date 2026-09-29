using System;
using System.Diagnostics;
using System.Globalization;
using ExcelDna.Integration;
using ExcelModelingToolkit.Core.Formatting;

namespace ExcelModelingToolkit.AddIn;

/// <summary>
/// Runs a formatting cycle on the selection: a thin COM adapter over <see cref="CycleEngine"/>. Reads the active
/// cell only, then writes the whole selection with one COM property write. Call in macro context on the main
/// thread. Any COM write clears Excel's undo history (spike K2c); our own undo arrives in Phase 3b.
/// </summary>
internal static class CycleCommand
{
    /// <summary><c>xlNone</c>: <c>Interior.Pattern</c> of a cell with no fill.</summary>
    private const int XlNone = -4142;

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

        var elapsed = trace.ElapsedMs ?? stopwatch.Elapsed.TotalMilliseconds;
        DiagnosticsLog.Write(
            actionId,
            "key=" + source,
            "ms=" + elapsed.ToString("0.0", CultureInfo.InvariantCulture),
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
            StatusBar.Show($"{label}: open a workbook first.");
            return "no workbook";
        }

        if (!SelectionIsCells())
        {
            StatusBar.Show($"{label}: select cells first.");
            return "selection is not cells";
        }

        // COM objects are held as object so only the adapter helpers below bind late.
        object selection = app.Selection;
        object activeCell = app.ActiveCell;
        trace.Cells = CellCount(selection);

        var current = ReadValue(activeCell, cycle.Kind);
        Session.States.TryGetValue(cycle.Id, out var previous);
        var step = Session.Engine.Next(cycle, current, SelectionKey(selection), previous);
        trace.Item = $"{step.Index + 1}/{cycle.Items.Count} {step.Item.Name}";

        try
        {
            Write(selection, cycle.Kind, step.Item);
        }
        catch (Exception ex) when (cycle.Kind == CycleKind.NumberFormat)
        {
            StatusBar.Show(
                $"{label}: Excel refused \"{step.Item.Name}\". The workbook may have too many number formats, " +
                $"or the sheet may be protected. ({ex.Message})");
            return "error: " + ex.Message;
        }

        trace.ElapsedMs = stopwatch.Elapsed.TotalMilliseconds;

        var state = step.State;
        if (cycle.Kind == CycleKind.NumberFormat)
        {
            // Excel may normalize the code; remember what it reports so the cell is recognized next time.
            if (ReadValue(activeCell, CycleKind.NumberFormat).NumberFormat is string text)
            {
                state = Session.Engine.RecordReadBack(cycle, state, text);
            }
        }

        Session.States[cycle.Id] = state;
        StatusBar.Show($"{label}: {step.Item.Name} ({step.Index + 1}/{cycle.Items.Count})");
        return "ok " + step.Reason;
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

    /// <summary>The active cell's current value for <paramref name="kind"/>; unknown if mixed or unreadable.</summary>
    private static CycleValue ReadValue(object activeCell, CycleKind kind)
    {
        dynamic cell = activeCell;
        switch (kind)
        {
            case CycleKind.NumberFormat:
                object format = cell.NumberFormat;
                return format is string code ? CycleValue.FromNumberFormat(code) : CycleValue.Unknown;
            case CycleKind.FontColor:
                object fontColor = cell.Font.Color;
                return ToColor(fontColor);
            case CycleKind.FillColor:
                dynamic interior = cell.Interior;
                object pattern = interior.Pattern;
                if (ToInt(pattern) == XlNone)
                {
                    return CycleValue.FromColor(OleColor.NoFill);
                }

                object fillColor = interior.Color;
                return ToColor(fillColor);
            default:
                return CycleValue.Unknown;
        }
    }

    /// <summary>Applies <paramref name="item"/> to the whole selection with one COM property write.</summary>
    private static void Write(object range, CycleKind kind, CycleItem item)
    {
        dynamic selection = range;
        switch (item)
        {
            case NumberFormatItem format when kind == CycleKind.NumberFormat:
                selection.NumberFormat = format.Code;
                break;
            case ColorItem color when kind == CycleKind.FontColor:
                selection.Font.Color = color.Color.OleValue;
                break;
            case ColorItem color when kind == CycleKind.FillColor && color.Color.IsNoFill:
                selection.Interior.Pattern = XlNone;
                break;
            case ColorItem color when kind == CycleKind.FillColor:
                selection.Interior.Color = color.Color.OleValue;
                break;
            default:
                throw new InvalidOperationException($"\"{item.Name}\" cannot be applied by a {kind} cycle.");
        }
    }

    /// <summary>An OLE color from a COM value (Excel returns a double); unknown for DBNull (mixed) or out of range.</summary>
    private static CycleValue ToColor(object? value)
    {
        var ole = ToInt(value);
        return ole is int v && v >= 0 && v <= OleColor.MaxOleValue
            ? CycleValue.FromColor(OleColor.FromOle(v))
            : CycleValue.Unknown;
    }

    private static int? ToInt(object? value) => value switch
    {
        int i => i,
        double d when d >= int.MinValue && d <= int.MaxValue && d == Math.Floor(d) => (int)d,
        _ => null,
    };

    /// <summary>What one run did, for the diagnostics log.</summary>
    private sealed class Trace
    {
        public string? Label { get; set; }

        public long Cells { get; set; }

        public string Item { get; set; } = "-";

        public double? ElapsedMs { get; set; }
    }
}
