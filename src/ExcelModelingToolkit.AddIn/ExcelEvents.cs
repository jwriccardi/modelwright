using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using ExcelDna.Integration;

namespace ExcelModelingToolkit.AddIn;

/// <summary>
/// Drops undo history that can no longer be trusted, and keeps an open Trace In window with the workbook window in
/// use, from three Excel <c>Application</c> events received late-bound (a COM connection point on
/// <see cref="AppEvents"/>, so no Office PIA is needed):
/// <list type="bullet">
/// <item><c>WorkbookBeforeClose</c>: every snapshot of that workbook (by <c>FullName</c>). If the user then cancels
/// the close, the history is gone anyway: the safe side. An open Trace In window on a cell of that workbook is
/// closed too.</item>
/// <item><c>SheetChange</c> on whole rows or whole columns (a row or column inserted or deleted, or cleared): every
/// snapshot of that sheet, since its cells may have moved. Other changes are ignored; the restore still checks
/// every block before writing.</item>
/// <item><c>WindowActivate</c>: an open Trace In window is re-owned to the workbook window just activated (the user
/// clicked another workbook's window), so it stays in front of the window in use.</item>
/// </list>
/// Events are raised on Excel's main thread. If <c>Application.EnableEvents</c> is off, none arrive.
/// </summary>
internal static class ExcelEvents
{
    // Strong references: the connection point, and the sink Excel calls, must live until Disconnect.
    private static IConnectionPoint? _point;
    private static ApplicationEventSink? _sink;
    private static int _cookie;

    /// <summary>Starts receiving the events (call from AutoOpen). Returns true on success. Never throws.</summary>
    public static bool Connect()
    {
        Disconnect();
        try
        {
            var container = (IConnectionPointContainer)ExcelDnaUtil.Application;
            var events = typeof(AppEvents).GUID;
            container.FindConnectionPoint(ref events, out var point);
            var sink = new ApplicationEventSink();
            point.Advise(sink, out var cookie);
            _point = point;
            _sink = sink;
            _cookie = cookie;
            DiagnosticsLog.Write("ExcelEvents", "connected", "cookie=" + cookie.ToString(CultureInfo.InvariantCulture));
            return true;
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("ExcelEvents", "failed", ex.Message);
            return false;
        }
    }

    /// <summary>Stops receiving the events (AutoClose). Never throws.</summary>
    public static void Disconnect()
    {
        try
        {
            if (_point is not null && _sink is not null)
            {
                _point.Unadvise(_cookie);
                DiagnosticsLog.Write("ExcelEvents", "disconnected");
            }
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("ExcelEvents", "disconnect failed", ex.Message);
        }

        _point = null;
        _sink = null;
        _cookie = 0;
    }

    /// <summary>
    /// Handles <c>WorkbookBeforeClose</c>: closes an open Trace In window whose audited cell is in the workbook, and
    /// drops the workbook's undo history. Never throws.
    /// </summary>
    internal static void OnWorkbookBeforeClose(object workbook)
    {
        if (TraceSession.Current is TraceSession trace)
        {
            try
            {
                dynamic closing = workbook;
                string closingName = closing.FullName;
                trace.OnWorkbookClosing(workbook, closingName);
            }
            catch (Exception ex)
            {
                DiagnosticsLog.Write("TraceClose", "WorkbookBeforeClose failed", ex.Message);
            }
        }

        if (Session.Undo.UndoCount + Session.Undo.RedoCount == 0)
        {
            return;
        }

        try
        {
            dynamic book = workbook;
            string fullName = book.FullName;
            var dropped = Session.Undo.InvalidateWorkbook(fullName);
            DiagnosticsLog.Write("UndoInvalidated", "workbook closing", fullName, "dropped=" + dropped.ToString(CultureInfo.InvariantCulture));
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("UndoInvalidated", "WorkbookBeforeClose failed", ex.Message);
        }
    }

    /// <summary>
    /// Handles <c>WindowActivate</c>: re-owns an open Trace In window to the activated workbook window (its
    /// <c>Hwnd</c>; <see cref="TraceSession.Follow"/> accepts only an Excel workbook window). Never throws.
    /// </summary>
    internal static void OnWindowActivate(object window)
    {
        if (TraceSession.Current is not TraceSession trace)
        {
            return;
        }

        try
        {
            dynamic activated = window;
            object hwnd = activated.Hwnd;
            trace.Follow(new IntPtr(Convert.ToInt64(hwnd, CultureInfo.InvariantCulture)));
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("TraceWindowError", "WindowActivate failed", ex.Message);
        }
    }

    /// <summary>Handles <c>SheetChange</c>: acts only on whole rows or whole columns. Never throws.</summary>
    internal static void OnSheetChange(object worksheet, object target)
    {
        if (Session.Undo.UndoCount + Session.Undo.RedoCount == 0)
        {
            return;
        }

        try
        {
            dynamic sheet = worksheet;
            dynamic range = target;
            int columns = range.Columns.Count;
            int rows = range.Rows.Count;
            int sheetColumns = sheet.Columns.Count;
            int sheetRows = sheet.Rows.Count;
            if (columns != sheetColumns && rows != sheetRows)
            {
                return;
            }

            string sheetName = sheet.Name;
            string fullName = sheet.Parent.FullName;
            string address = range.Address;
            var dropped = Session.Undo.InvalidateSheet(fullName, sheetName);
            DiagnosticsLog.Write(
                "UndoInvalidated",
                "whole rows or columns changed",
                fullName + "|" + sheetName,
                "address=" + address,
                "dropped=" + dropped.ToString(CultureInfo.InvariantCulture));
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("UndoInvalidated", "SheetChange failed", ex.Message);
        }
    }
}

/// <summary>
/// The Excel <c>Application</c> events interface (the <c>AppEvents</c> dispinterface), declaring only the events
/// the add-in uses. The GUID and DispIds are Excel's, verified against the type library in <c>EXCEL.EXE</c> (Office
/// 16, read with <c>LoadTypeLibEx</c>, without starting Excel): <c>SheetChange</c> is 0x61C, <c>WorkbookBeforeClose</c>
/// 0x622, <c>WindowActivate</c> 0x614 (two parameters: the workbook and the window). Excel calls the others too; with
/// no member for their DispIds, the call is answered "member not found" and ignored.
/// </summary>
[ComVisible(true)]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
[Guid("00024413-0000-0000-C000-000000000046")]
public interface AppEvents
{
    /// <summary>Cells on a worksheet were changed by the user or an external link.</summary>
    [DispId(0x61c)]
    void SheetChange(object sheet, object target);

    /// <summary>A workbook is about to close.</summary>
    [DispId(0x622)]
    void WorkbookBeforeClose(object workbook, ref bool cancel);

    /// <summary>A workbook window was activated.</summary>
    [DispId(0x614)]
    void WindowActivate(object workbook, object window);
}

/// <summary>Receives <see cref="AppEvents"/> from Excel and hands them to <see cref="ExcelEvents"/>. Not for other callers.</summary>
[ComVisible(true)]
[ClassInterface(ClassInterfaceType.None)]
public sealed class ApplicationEventSink : AppEvents
{
    /// <inheritdoc />
    public void SheetChange(object sheet, object target) => ExcelEvents.OnSheetChange(sheet, target);

    /// <inheritdoc />
    public void WorkbookBeforeClose(object workbook, ref bool cancel) => ExcelEvents.OnWorkbookBeforeClose(workbook);

    /// <inheritdoc />
    public void WindowActivate(object workbook, object window) => ExcelEvents.OnWindowActivate(window);
}
