using System;
using System.Diagnostics;
using System.Globalization;
using ExcelDna.Integration;
using Modelwright.Core.Trace;

namespace Modelwright.AddIn;

/// <summary>
/// Trace In (Ctrl+Shift+[) and Last Audited Cell (Ctrl+Shift+\): docs/PLAN.md section 4.5. Call in macro context on
/// the main thread (they are OnKey and ribbon macros, and use the status bar). Neither writes to a cell (Trace In may
/// open a closed workbook a precedent is in; see <see cref="ExcelPrecedentProvider"/>), and each writes one
/// diagnostics log line with its timing. Whether running a macro that writes nothing keeps Excel's undo history is
/// not yet measured (docs/research/05: unverified); the navigation after it avoids macro context
/// (<see cref="TraceSession"/>), and tests/excel-smoke/trace-smoke.ps1 checks undo across a whole trace.
/// </summary>
internal static class TraceCommand
{
    /// <summary>
    /// Opens Trace In on the active cell, or, with the window open, traces the active cell instead. A cell without
    /// a formula only gets a status-bar message (as Macabacus does). Each trace is recorded in
    /// <see cref="Session.Audits"/> for Last Audited Cell. <paramref name="source"/> is the key or <c>ribbon</c>.
    /// Never throws.
    /// </summary>
    public static void TraceIn(string source)
    {
        var stopwatch = Stopwatch.StartNew();
        var log = new OpenLog();
        string result;
        try
        {
            result = Open(log, stopwatch);
        }
        catch (Exception ex)
        {
            result = "error: " + ex.Message;
            DiagnosticsLog.Write("TraceOpenFailed", ex.ToString());
            StatusBar.Show("Trace In failed: " + ex.Message);
        }

        // ms: the whole command (tree, window and status bar). treeMs: reading the audited cell's precedents.
        DiagnosticsLog.Write(
            "TraceOpen",
            "key=" + source,
            "ms=" + Ms(stopwatch.Elapsed.TotalMilliseconds),
            "treeMs=" + Ms(log.TreeMs),
            "windowMs=" + Ms(log.WindowMs),
            "refs=" + log.References.ToString(CultureInfo.InvariantCulture),
            "retrace=" + (log.Retrace ? "true" : "false"),
            "target=" + log.Target,
            result);
    }

    /// <summary>
    /// Goes back to the newest audited cell and forgets it, so each press goes one audit further back (up to
    /// <see cref="AuditHistory.Capacity"/>). Closes an open Trace In window first (staying put). A workbook that
    /// has closed, or a sheet renamed or deleted since, is reported. Never throws.
    /// </summary>
    public static void LastAuditedCell(string source)
    {
        var stopwatch = Stopwatch.StartNew();
        var target = "-";
        string result;
        try
        {
            result = GoBack(ref target);
        }
        catch (Exception ex)
        {
            result = "error: " + ex.Message;
            StatusBar.Show("Last Audited Cell failed: " + ex.Message);
        }

        DiagnosticsLog.Write(
            "LastAuditedCell",
            "key=" + source,
            "ms=" + Ms(stopwatch.Elapsed.TotalMilliseconds),
            "target=" + target,
            "left=" + Session.Audits.Count.ToString(CultureInfo.InvariantCulture),
            result);
    }

    private static string Open(OpenLog log, Stopwatch stopwatch)
    {
        dynamic app = ExcelDnaUtil.Application;
        object? cell = null;
        try
        {
            cell = app.ActiveCell;
        }
        catch (Exception)
        {
            // A chart sheet or Protected View: no active cell.
        }

        if (cell is null)
        {
            StatusBar.Show("Trace In: select a cell first.");
            return "no active cell";
        }

        dynamic c = cell;
        object sheetObject = c.Worksheet;
        dynamic sheet = sheetObject;
        object bookObject = sheet.Parent;
        dynamic book = bookObject;
        string sheetName = sheet.Name;
        string address = ExcelPrecedentProvider.LocalAddress(cell);
        log.Target = (string)book.Name + "|" + sheetName + "|" + address;
        object hasFormula = c.HasFormula;
        if (hasFormula is not true)
        {
            StatusBar.Show($"Trace In: {sheetName}!{address} has no formula to trace.");
            return "no formula";
        }

        var provider = new ExcelPrecedentProvider();
        var root = provider.CreateRoot(cell);
        var tree = new PrecedentTree(provider, root);
        log.TreeMs = stopwatch.Elapsed.TotalMilliseconds;
        log.References = tree.Root.Children.Count;
        log.Retrace = TraceSession.Current is not null;

        Session.Audits.Push(root);
        string fullName = book.FullName;
        var note = TraceSession.Show(provider, tree, cell, fullName);
        log.WindowMs = stopwatch.Elapsed.TotalMilliseconds - log.TreeMs;

        var count = log.References;
        var message = $"Trace In: {sheetName}!{address}: {count} precedent{(count == 1 ? string.Empty : "s")}{note}";
        if (provider.OpenedWorkbooks.Count > 0)
        {
            message += "; opened " + string.Join(", ", provider.OpenedWorkbooks) + " (read-only)";
        }

        if (provider.ParseFallbacks.Count > 0)
        {
            message += "; formula not parsed, showing Excel's same-sheet precedents";
        }

        StatusBar.Show(message);
        return provider.ParseFallbacks.Count > 0 ? "ok (direct precedents fallback)" : "ok";
    }

    private static string GoBack(ref string target)
    {
        dynamic app = ExcelDnaUtil.Application;

        // An audit of the cell Excel is already on (Esc returned there) is skipped when there is an older one, so
        // each press moves.
        var here = ActiveCellKey((object)app);
        var audit = Session.Audits.Pop();
        while (audit is not null && Session.Audits.Count > 0 && string.Equals(Key(audit), here, StringComparison.OrdinalIgnoreCase))
        {
            audit = Session.Audits.Pop();
        }

        if (audit is null)
        {
            StatusBar.Show("Last Audited Cell: there is no audited cell to go back to.");
            return "history empty";
        }

        target = Key(audit);
        TraceSession.Current?.Close(TraceCloseMode.StayOnCurrentCell, "last audited cell");

        object? book = null;
        try
        {
            book = app.Workbooks.Item(audit.Workbook);
        }
        catch (Exception)
        {
            // Closed (or renamed by Save As).
        }

        if (book is null)
        {
            StatusBar.Show($"Last Audited Cell: {audit.Workbook} is no longer open.");
            return "workbook closed";
        }

        object? sheet = null;
        try
        {
            dynamic wb = book;
            sheet = wb.Worksheets.Item(audit.Sheet);
        }
        catch (Exception)
        {
            // Renamed or deleted.
        }

        if (sheet is null)
        {
            StatusBar.Show($"Last Audited Cell: sheet '{audit.Sheet}' was renamed or deleted.");
            return "sheet missing";
        }

        dynamic ws = sheet;
        object visible = ws.Visible;
        if (Convert.ToInt32(visible, CultureInfo.InvariantCulture) != -1)
        {
            StatusBar.Show($"Last Audited Cell: '{audit.Sheet}' is a hidden sheet: unhide it to go there.");
            return "hidden sheet";
        }

        dynamic workbook = book;
        if (!ExcelPrecedentProvider.WorkbookIsVisible(book))
        {
            // As Trace In's own navigation says it (TraceSession.GoTo).
            StatusBar.Show($"Last Audited Cell: {(string)workbook.Name} is a hidden workbook: unhide it to go there.");
            return "hidden workbook";
        }

        object range = ws.Range(audit.Address);
        string? activeName = null;
        try
        {
            activeName = app.ActiveWorkbook?.Name;
        }
        catch (Exception)
        {
            // No active workbook; activate below.
        }

        if (!string.Equals((string)workbook.Name, activeName, StringComparison.OrdinalIgnoreCase))
        {
            workbook.Activate();
        }

        app.Goto(range);
        var left = Session.Audits.Count;
        StatusBar.Show($"Last Audited Cell: {audit.Label} ({left} earlier audit{(left == 1 ? string.Empty : "s")})");
        return "ok";
    }

    private static string Key(PrecedentItem audit) => audit.Workbook + "|" + audit.Sheet + "|" + audit.Address;

    // workbook|sheet|address of the active cell, or empty if there is none.
    private static string ActiveCellKey(object application)
    {
        try
        {
            dynamic app = application;
            object cell = app.ActiveCell;
            dynamic c = cell;
            string sheet = c.Worksheet.Name;
            string book = c.Worksheet.Parent.Name;
            return book + "|" + sheet + "|" + ExcelPrecedentProvider.LocalAddress(cell);
        }
        catch (Exception)
        {
            return string.Empty;
        }
    }

    private static string Ms(double milliseconds) => milliseconds.ToString("0.0", CultureInfo.InvariantCulture);

    /// <summary>What one Trace In did, for the diagnostics log.</summary>
    private sealed class OpenLog
    {
        public double TreeMs { get; set; }

        public double WindowMs { get; set; }

        public int References { get; set; }

        public bool Retrace { get; set; }

        public string Target { get; set; } = "-";
    }
}
