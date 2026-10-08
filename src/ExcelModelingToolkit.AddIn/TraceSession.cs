using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Windows.Threading;
using ExcelDna.Integration;
using ExcelModelingToolkit.Core.Trace;

namespace ExcelModelingToolkit.AddIn;

/// <summary>
/// The open Trace In window and what it shows (docs/PLAN.md section 4.5): the audited cell, its
/// <see cref="PrecedentTree"/> read through an <see cref="ExcelPrecedentProvider"/>, and the
/// <see cref="TraceWindow"/>. Turns keys (<see cref="TraceKeyHook"/>) and clicks into tree moves and
/// <c>Application.Goto</c> calls, and closes the window. At most one is open (<see cref="Current"/>).
/// </summary>
/// <remarks>
/// <para>
/// Main thread. Navigation only changes the selection (<c>Application.Goto</c>, and activating another workbook's
/// window), so no cell is written. It also runs <b>outside macro context</b>: keys and clicks are queued and run from
/// Excel's message loop (posted to the hook's hidden window, or the window's dispatcher) and call Excel there, not through
/// <c>ExcelAsyncUtil.QueueAsMacro</c>. Spike K2/K2b found that work done in macro context loses Excel's earlier undo
/// history while the same work from the message loop keeps it, so this is what keeps native undo across a trace
/// (the C API is not available there, so navigation messages go to the window's footer, not the status bar). Opening
/// Trace In itself is an OnKey (or ribbon) macro; see <see cref="TraceCommand"/>.
/// </para>
/// <para>
/// Keys, clicks, OK and Cancel all go through one first-in, first-out queue (<see cref="SerialCommandQueue"/>), run one
/// at a time in the order they arrived: a command that arrives while another is still running (opening a closed
/// workbook can pump messages) waits for it. While a cell is being edited, tree moves and clicks are ignored (Excel
/// would reject the calls). Every open, navigation and close writes one diagnostics log line with its timing.
/// </para>
/// <para>
/// <b>F2</b> edits the selected reference where it is written (<see cref="ReferenceEdit"/>): it goes to that cell
/// and sends Excel the keys that select the reference and switch to Point mode (<see cref="TraceKeyHook.Send"/>).
/// The edit is Excel's own, made with keys and <c>Application.Goto</c> only, so Excel's undo of it should survive.
/// When the key that ends it (Enter, Tab, Esc) has gone to Excel, the session goes back to the cell; then, or when
/// the cell changes otherwise (a click elsewhere commits), if the formula changed it rebuilds the tree in the same
/// window.
/// </para>
/// </remarks>
internal sealed class TraceSession
{
    private readonly TraceWindow _window;
    private readonly Stopwatch _openFor = Stopwatch.StartNew();
    private readonly SerialCommandQueue _queue;
    private ExcelPrecedentProvider _provider;
    private PrecedentTree _tree;
    private object _audited;
    private string _auditedWorkbook;
    private TraceUiState _ui;
    private bool _closing;
    private bool _closed;
    private int _navigations;

    // Up/Down/Left/Right presses posted by the hook and not yet run: while more are queued (an arrow held down), a
    // move updates the tree without going there, so Excel only goes to where the presses end.
    private int _pendingMoves;

    // True if a move left the tree's selection without going there (a newer move was queued).
    private bool _skippedGoTo;

    // The F2 edit Excel is doing, from the keys sent until its end is seen; null when none.
    private PendingEdit? _edit;

    private TraceSession(ExcelPrecedentProvider provider, PrecedentTree tree, object audited, string auditedWorkbook)
    {
        _provider = provider;
        _tree = tree;
        _audited = audited;
        _auditedWorkbook = auditedWorkbook;
        _ui = UiStateStore.Load();
        _queue = new SerialCommandQueue(ScheduleDrain, ex => DiagnosticsLog.Write("TraceWindowError", ex.ToString()));
        _window = new TraceWindow();
        _window.RowClicked += OnRowClicked;
        _window.OkClicked += () => Enqueue(() => Close(TraceCloseMode.StayOnCurrentCell, "ok"));
        _window.CancelClicked += () => Enqueue(() => Close(TraceKeys.CancelCloseMode, "cancel"));
        _window.WrapChanged += wrap => _ui = _ui.WithWrapFormula(wrap);
        _window.Closed += (sender, e) => OnWindowClosed();
    }

    /// <summary>The open trace, or null.</summary>
    public static TraceSession? Current { get; private set; }

    /// <summary>The window's handle, for the hook's focus check.</summary>
    public IntPtr WindowHandle => _window.Handle;

    /// <summary>The Excel window that owns the window (the hook follows the keyboard away from it).</summary>
    public IntPtr OwnerHandle => _window.OwnerHandle;

    /// <summary>True while an F2 edit has started and its end has not been seen (see <see cref="EndEdit"/>).</summary>
    public bool AwaitsEditEnd => _edit is not null && !_closing;

    /// <summary>
    /// Shows a trace: opens the window, or, if one is open, shows the new trace in it, and makes sure the key hook is
    /// installed. Returns a note for the status bar (empty if none). If the window cannot be filled it is closed and
    /// the exception propagates. Call in macro context (the Trace In command).
    /// </summary>
    public static string Show(ExcelPrecedentProvider provider, PrecedentTree tree, object audited, string auditedWorkbook)
    {
        var session = Current;
        if (session is null)
        {
            session = new TraceSession(provider, tree, audited, auditedWorkbook);
            session._window.ShowFor(ActiveWindowHandle(), session._ui.Bounds, session._ui.WrapFormula);
            Current = session;
        }
        else
        {
            // Open already: the new trace replaces the old one in the same window.
            session._provider = provider;
            session._tree = tree;
            session._audited = audited;
            session._auditedWorkbook = auditedWorkbook;
            session._queue.Clear(); // keys meant for the old trace
            session._pendingMoves = 0;
            session._edit = null;
            session.ReOwn();
        }

        try
        {
            session.Load();
        }
        catch (Exception)
        {
            session.Abort("could not show the trace");
            throw;
        }

        if (!TraceKeyHook.Install())
        {
            session.Message("The keyboard could not be connected; use the mouse, OK and Cancel.");
            return " (keyboard unavailable: use the mouse)";
        }

        return string.Empty;
    }

    /// <summary>
    /// Queues a command (a key from <see cref="TraceKeyHook"/>, a click, OK, Cancel) behind those already queued; it
    /// runs from Excel's message loop, outside macro context, after the caller returns. False if it could not be
    /// queued.
    /// </summary>
    public bool Enqueue(Action command) => _queue.Enqueue(command);

    /// <summary>Counts a move key the hook has queued (see <see cref="Handle"/>).</summary>
    public void MoveQueued() => _pendingMoves++;

    /// <summary>
    /// Re-owns the window to <paramref name="excelWindow"/> (the user switched workbooks: the top-level window of the
    /// worksheet grid that has the keyboard, or the workbook window Excel activated), so the window the keys drive
    /// stays in front. Only an Excel workbook window (class XLMAIN) of Excel's main thread is followed. No Excel call.
    /// </summary>
    public void Follow(IntPtr excelWindow)
    {
        if (!_closing && excelWindow != _window.Handle && excelWindow != _window.OwnerHandle &&
            TraceKeyHook.IsExcelWorkbookWindow(excelWindow))
        {
            _window.ReOwn(excelWindow);
        }
    }

    /// <summary>
    /// Runs a key's command (queued by <see cref="TraceKeyHook"/> at <paramref name="pressed"/>, a
    /// <see cref="Stopwatch"/> timestamp) and logs it with the time from the key press. Never throws.
    /// </summary>
    public void Handle(TraceKeyCommand command, long pressed, string source)
    {
        var isMove = command == TraceKeyCommand.Up || command == TraceKeyCommand.Down ||
            command == TraceKeyCommand.Left || command == TraceKeyCommand.Right;
        if (isMove && _pendingMoves > 0)
        {
            _pendingMoves--;
        }

        if (_closing || !ReferenceEquals(Current, this))
        {
            return;
        }

        if (TraceKeys.CloseMode(command) is null && IsEditing())
        {
            LogNavigation(command.ToString(), source, Elapsed(pressed), Stopwatch.StartNew(), TreeMove.None, "-", "ignored: editing a cell");
            return;
        }

        Run(command, pressed, source, goThere: !isMove || _pendingMoves == 0);
    }

    private void Run(TraceKeyCommand command, long pressed, string source, bool goThere)
    {
        var closeMode = TraceKeys.CloseMode(command);
        if (closeMode is TraceCloseMode mode)
        {
            Close(mode, source);
            return;
        }

        if (command == TraceKeyCommand.EditReference)
        {
            StartEdit(pressed, source);
            return;
        }

        var stopwatch = Stopwatch.StartNew();
        var move = TreeMove.None;
        var target = "-";
        string result;
        try
        {
            switch (command)
            {
                case TraceKeyCommand.Up:
                    move = _tree.MoveUp();
                    break;
                case TraceKeyCommand.Down:
                    move = _tree.MoveDown();
                    break;
                case TraceKeyCommand.Left:
                    move = _tree.MoveLeft();
                    break;
                case TraceKeyCommand.Right:
                    move = _tree.MoveRight();
                    break;
                case TraceKeyCommand.ToggleEvaluate:
                    Message("Evaluate functions & groups (Ctrl+E) is coming in a later version.");
                    break;
            }

            result = Apply(move, command == TraceKeyCommand.Up || command == TraceKeyCommand.Down, goThere, ref target);
            if (goThere && move != TreeMove.Moved && _skippedGoTo)
            {
                // The last of a run of queued moves did not move (the end of the list): go where the run ended.
                target = _tree.Selected.Item.Label;
                result += "; " + Navigate(_tree.Selected);
            }
        }
        catch (Exception ex)
        {
            result = "error: " + ex.Message;
            Message("Trace In: " + ex.Message);
        }

        LogNavigation(command.ToString(), source, Elapsed(pressed), stopwatch, move, target, result);
    }

    /// <summary>Moves, snaps or resizes the window (no Excel call). Never throws.</summary>
    public void ApplyWindowCommand(TraceKeyCommand command)
    {
        try
        {
            if (!_closing)
            {
                var bounds = _window.ApplyWindowCommand(command);
                DiagnosticsLog.Write("TraceWindowKey", command.ToString(), bounds.ToString());
            }
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("TraceWindowKey", command.ToString(), "error: " + ex.Message);
        }
    }

    /// <summary>
    /// The end of an F2 edit (queued by the hook after the Enter, Tab or Esc that went to Excel, or after a change to
    /// the edited cell): unless Excel is still editing (an Enter on a formula Excel rejected, say), goes back to the
    /// edited cell if <paramref name="goBack"/> (Enter may have moved the selection), and if its formula changed,
    /// rebuilds the tree in the same window with the selection where it was. Logs the end. Never throws.
    /// </summary>
    public void EndEdit(bool goBack, string source)
    {
        var edit = _edit;
        if (edit is null || _closing || !ReferenceEquals(Current, this) || IsEditing())
        {
            return;
        }

        _edit = null;
        var stopwatch = Stopwatch.StartNew();
        var changed = false;
        string result;
        try
        {
            var formula = ExcelPrecedentProvider.FormulaText(edit.Cell);
            changed = !string.Equals(formula, edit.Formula, StringComparison.Ordinal);
            result = goBack ? GoTo(edit.Cell) : "stayed";
            if (changed)
            {
                Reload(edit.Path);
                result += "; tree rebuilt";
            }
        }
        catch (Exception ex)
        {
            result = "error: " + ex.Message;
            Message("Trace In: " + ex.Message);
        }

        DiagnosticsLog.Write(
            "TraceEditEnd",
            "source=" + source,
            "changed=" + (changed ? "true" : "false"),
            "ms=" + Ms(edit.Started.Elapsed.TotalMilliseconds),
            "handleMs=" + Ms(stopwatch.Elapsed.TotalMilliseconds),
            "owner=" + edit.Span,
            result);
    }

    /// <summary>
    /// Handles <c>SheetChange</c>: a change to the cell of an F2 edit (committed some other way than Enter or Tab, a
    /// click elsewhere say) ends the edit once the event has returned (<see cref="EndEdit"/>, staying put). Never throws.
    /// </summary>
    public void OnSheetChange(object worksheet, object target)
    {
        var edit = _edit;
        if (edit is null || _closing)
        {
            return;
        }

        try
        {
            dynamic sheet = worksheet;
            string sheetName = sheet.Name;
            string bookName = sheet.Parent.Name;
            if (!string.Equals(sheetName, edit.Span.Sheet, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(bookName, edit.Span.Workbook, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            dynamic app = ExcelDnaUtil.Application;
            object? overlap = app.Intersect(target, edit.Cell);
            if (overlap is not null)
            {
                Enqueue(() => EndEdit(goBack: false, "sheet change"));
            }
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("TraceWindowError", "SheetChange failed", ex.Message);
        }
    }

    /// <summary>
    /// Closes the window: <see cref="TraceCloseMode.ReturnToAuditedCell"/> first goes back to the audited cell.
    /// Logs the close. Never throws.
    /// </summary>
    public void Close(TraceCloseMode mode, string source)
    {
        if (_closing)
        {
            return;
        }

        _closing = true;
        var stopwatch = Stopwatch.StartNew();
        var result = "ok";
        if (mode == TraceCloseMode.ReturnToAuditedCell)
        {
            try
            {
                result = GoTo(_audited);
            }
            catch (Exception ex)
            {
                result = "could not return to the audited cell: " + ex.Message;
            }
        }

        CloseWindow();
        DiagnosticsLog.Write(
            "TraceClose",
            "mode=" + mode,
            "source=" + source,
            "ms=" + Ms(stopwatch.Elapsed.TotalMilliseconds),
            "navigations=" + _navigations.ToString(CultureInfo.InvariantCulture),
            "openSeconds=" + _openFor.Elapsed.TotalSeconds.ToString("0.0", CultureInfo.InvariantCulture),
            result);
    }

    /// <summary>
    /// Closes the window without going anywhere (the audited workbook is closing, or the add-in is unloading).
    /// Never throws.
    /// </summary>
    public void Abort(string reason)
    {
        if (_closing)
        {
            return;
        }

        _closing = true;
        CloseWindow();
        DiagnosticsLog.Write("TraceClose", "mode=abort", "source=" + reason, "navigations=" + _navigations.ToString(CultureInfo.InvariantCulture));
    }

    /// <summary>
    /// A workbook is closing: if it is the audited cell's, the trace closes; if the window is owned by one of its
    /// windows (the trace went into it), the window is re-owned to the audited workbook's window first, since
    /// Windows destroys a window with its owner. Never throws.
    /// </summary>
    public void OnWorkbookClosing(object workbook, string fullName)
    {
        if (string.Equals(fullName, _auditedWorkbook, StringComparison.OrdinalIgnoreCase))
        {
            Abort("audited workbook closing");
            return;
        }

        try
        {
            if (!WindowsOf(workbook).Contains(_window.OwnerHandle))
            {
                return;
            }

            dynamic audited = _audited;
            object book = audited.Worksheet.Parent;
            var windows = WindowsOf(book);
            if (windows.Count > 0)
            {
                _window.ReOwn(windows[0]);
            }
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("TraceWindowError", "re-own before a workbook closed failed", ex.Message);
        }
    }

    private void Load()
    {
        var parsed = _provider.ParsedFormulaOf(_tree.Root.Item);
        var formula = parsed?.Formula ?? string.Empty;
        var segments = parsed is null ? new FormulaSegment[0] : FormulaColoring.Segment(parsed);
        string? note = null;
        if (_provider.ParseFallbacks.Count > 0)
        {
            note = "Not parsed (" + _provider.ParseFallbacks[0] + "): showing Excel's precedents on the same sheet only.";
        }

        _window.SetFormula(formula, segments, note);
        RefreshRows();
        _window.SetStatus(Where(_tree.Root));
    }

    // The audited cell's precedents changed (an F2 edit): a new tree in the same window, with the selection brought
    // back along the same path, or as near as the new tree allows.
    private void Reload(IReadOnlyList<int> path)
    {
        var provider = new ExcelPrecedentProvider();
        var tree = new PrecedentTree(provider, provider.CreateRoot(_audited));
        _provider = provider;
        _tree = tree;
        _pendingMoves = 0;
        _skippedGoTo = false;
        try
        {
            tree.SelectPath(path);
        }
        catch (PrecedentsUnavailableException)
        {
            // Excel is busy: the selection stays as deep as the path was followed.
        }

        Load();
        _window.SetStatus(Where(_tree.Selected));
    }

    // F2 (outside macro context): see EditReference. Logs the start.
    private void StartEdit(long pressed, string source)
    {
        var stopwatch = Stopwatch.StartNew();
        var owner = "-";
        var keyCount = 0;
        string result;
        try
        {
            result = EditReference(ref owner, ref keyCount);
        }
        catch (Exception ex)
        {
            result = "error: " + ex.Message;
            Message("Trace In: " + ex.Message);
        }

        DiagnosticsLog.Write(
            "TraceEditReference",
            "source=" + source,
            "ms=" + Ms(Elapsed(pressed)),
            "handleMs=" + Ms(stopwatch.Elapsed.TotalMilliseconds),
            "owner=" + owner,
            "keys=" + keyCount.ToString(CultureInfo.InvariantCulture),
            result);
    }

    // Goes to the cell whose formula holds the selected reference and sends Excel the keys that select the reference
    // and switch to Point mode (ReferenceEdit.Keys). A row that is not such a reference, or a reference that cannot be
    // selected that way, gets a plain F2 on the active cell, as without the window. Returns the log result.
    private string EditReference(ref string owner, ref int keyCount)
    {
        _edit = null;
        var span = ReferenceEdit.SpanOf(_tree.Selected);
        if (span is null)
        {
            return PlainF2("not a reference in a cell's formula", null);
        }

        owner = span.ToString();
        var place = span.Sheet + "!" + span.Address;
        var cell = ExcelPrecedentProvider.CellAt(span.Workbook, span.Sheet, span.Address);
        if (cell is null)
        {
            return PlainF2("its cell is gone", $"Edit the reference: {place} is no longer open.");
        }

        var formula = ExcelPrecedentProvider.FormulaText(cell);
        if (!string.Equals(formula, span.Formula, StringComparison.Ordinal))
        {
            return PlainF2("the formula changed", $"Edit the reference: {place} changed since it was traced (Ctrl+Shift+[ traces it again).");
        }

        // The keys count characters in the formula as Excel's editor shows it; only en-US Excel shows the parsed text.
        if (!string.Equals(ExcelPrecedentProvider.LocalFormulaText(cell), formula, StringComparison.Ordinal))
        {
            return PlainF2("the editor shows the formula localized", "Edit the reference: not available for this Excel language yet.");
        }

        if (IsLockedOnProtectedSheet(cell))
        {
            Message($"Edit the reference: {place} is locked on a protected sheet.");
            return "protected";
        }

        var keys = ReferenceEdit.Keys(span);
        if (keys is null)
        {
            return PlainF2("too many keys", "Edit the reference: the formula is too long to select it with keystrokes.");
        }

        keyCount = keys.Count;
        var path = _tree.PathOf(_tree.Selected);
        var went = GoTo(cell);
        if (went != "ok")
        {
            return went;
        }

        if (!TraceKeyHook.Send(keys, out var failure))
        {
            Message("Edit the reference: " + failure + ".");
            return "not sent: " + failure;
        }

        _edit = new PendingEdit(span, cell, formula!, path);
        Message($"Editing {span.Text} in {place}: point at its replacement; Enter commits, Esc cancels.");
        return "ok";
    }

    // F2 for the active cell, as Excel would have had it without the window. Returns the log result.
    private string PlainF2(string reason, string? message)
    {
        if (message is not null)
        {
            Message(message);
        }

        if (TraceKeyHook.Send(ReferenceEdit.PlainF2, out var failure))
        {
            return "plain F2: " + reason;
        }

        Message("F2 could not be passed to Excel: " + failure + ".");
        return "plain F2 not sent (" + failure + "): " + reason;
    }

    // True if Excel would refuse to edit the cell (locked, on a protected sheet). False if that cannot be read.
    private static bool IsLockedOnProtectedSheet(object cell)
    {
        try
        {
            dynamic c = cell;
            object protectedContents = c.Worksheet.ProtectContents;
            object locked = c.Locked;
            return protectedContents is true && locked is true;
        }
        catch (Exception)
        {
            return false;
        }
    }

    // After a tree move: redraws the rows if they changed, and goes to the newly selected node (unless goThere is
    // false: another move is already queued).
    private string Apply(TreeMove move, bool selectionOnly, bool goThere, ref string target)
    {
        switch (move)
        {
            case TreeMove.None:
                return "no move";
            case TreeMove.Expanded:
            case TreeMove.Collapsed:
                RefreshRows();
                return move == TreeMove.Expanded ? "expanded" : "collapsed";
            default:
                if (selectionOnly)
                {
                    _window.Select(_tree.SelectedIndex);
                }
                else
                {
                    RefreshRows();
                }

                target = _tree.Selected.Item.Label;
                if (!goThere)
                {
                    _skippedGoTo = true;
                    return "selected (a newer key is queued)";
                }

                return Navigate(_tree.Selected);
        }
    }

    // The row is resolved now, while the list still shows what was clicked; the click itself runs after the mouse
    // handler returns.
    private void OnRowClicked(int index, bool doubleClick, bool onExpander)
    {
        var visible = _tree.VisibleNodes;
        if (index >= 0 && index < visible.Count)
        {
            var node = visible[index];
            var tree = _tree;
            Enqueue(() => Click(tree, node, doubleClick, onExpander));
        }
    }

    // A click on a row (outside macro context): selects it and goes there; a double-click or a click on the expander
    // expands or collapses it instead. Ignored if the trace changed since, or while a cell is being edited.
    private void Click(PrecedentTree tree, PrecedentNode node, bool doubleClick, bool onExpander)
    {
        if (_closing || !ReferenceEquals(Current, this) || !ReferenceEquals(tree, _tree))
        {
            return;
        }

        if (IsEditing())
        {
            Message("Finish editing the cell (Enter or Esc) to use Trace In.");
            return;
        }

        var stopwatch = Stopwatch.StartNew();
        var target = node.Item.Label;
        var move = TreeMove.None;
        string result;
        try
        {
            var index = IndexOf(node);
            if (index < 0)
            {
                result = "row no longer shown";
            }
            else if (doubleClick || onExpander)
            {
                _tree.Select(node);
                if (node.IsExpanded)
                {
                    move = _tree.Collapse(node) ? TreeMove.Collapsed : TreeMove.None;
                }
                else
                {
                    move = _tree.Expand(node) ? TreeMove.Expanded : TreeMove.None;
                }

                RefreshRows();
                result = move.ToString().ToLowerInvariant();
            }
            else if (_tree.Select(node))
            {
                move = TreeMove.Moved;
                _window.Select(index);
                result = Navigate(node);
            }
            else
            {
                result = "already selected";
            }
        }
        catch (Exception ex)
        {
            // Excel was busy, say: the click may have selected the row before the expand failed.
            result = "error: " + ex.Message;
            Message("Trace In: " + ex.Message);
            _window.Select(_tree.SelectedIndex);
        }

        LogNavigation(doubleClick ? "DoubleClick" : onExpander ? "ExpanderClick" : "Click", "mouse",
            stopwatch.Elapsed.TotalMilliseconds, stopwatch, move, target, result);
    }

    private int IndexOf(PrecedentNode node)
    {
        var visible = _tree.VisibleNodes;
        for (var i = 0; i < visible.Count; i++)
        {
            if (ReferenceEquals(visible[i], node))
            {
                return i;
            }
        }

        return -1;
    }

    // Goes to a node's cells, unless it is not a place or its sheet is hidden. Returns the log result.
    private string Navigate(PrecedentNode node)
    {
        _skippedGoTo = false;
        object? range;
        string? reason = null;
        range = ReferenceEquals(node, _tree.Root) ? _audited : _provider.NavigationRange(node.Item, out reason);
        if (range is null)
        {
            if (reason is not null)
            {
                Message(reason);
            }

            return reason is null ? "not a place" : "not navigable";
        }

        var result = GoTo(range);
        if (result == "ok")
        {
            _navigations++;
            _window.SetStatus(Where(node));
        }

        return result;
    }

    // Application.Goto, activating the range's workbook window first if needed, then re-owning the window to the
    // active Excel window (spike K4 fix). A hidden sheet or workbook is reported, not gone to.
    private string GoTo(object range)
    {
        dynamic r = range;
        object sheetObject = r.Worksheet;
        dynamic sheet = sheetObject;
        string sheetName = sheet.Name;
        object visible = sheet.Visible;
        if (Convert.ToInt32(visible, CultureInfo.InvariantCulture) != -1)
        {
            Message($"'{sheetName}' is a hidden sheet: unhide it to go there.");
            return "hidden sheet";
        }

        object bookObject = sheet.Parent;
        if (!ExcelPrecedentProvider.WorkbookIsVisible(bookObject))
        {
            dynamic hiddenBook = bookObject;
            Message($"{(string)hiddenBook.Name} is a hidden workbook: unhide it to go there.");
            return "hidden workbook";
        }

        dynamic app = ExcelDnaUtil.Application;
        dynamic book = bookObject;
        string bookName = book.Name;
        string? activeName = null;
        try
        {
            activeName = app.ActiveWorkbook?.Name;
        }
        catch (Exception)
        {
            // No active workbook; activate below.
        }

        try
        {
            if (!string.Equals(bookName, activeName, StringComparison.OrdinalIgnoreCase))
            {
                book.Activate();
            }

            app.Goto(range);
        }
        catch (Exception ex)
        {
            Message($"Trace In: could not go to {ExcelPrecedentProvider.LocalAddress(range)}: {ex.Message}");
            return "goto failed: " + ex.Message;
        }

        ReOwn();
        return "ok";
    }

    private void ReOwn()
    {
        var active = ActiveWindowHandle();
        if (active != IntPtr.Zero && active != _window.OwnerHandle)
        {
            _window.ReOwn(active);
        }
    }

    private void RefreshRows()
    {
        var visible = _tree.VisibleNodes;
        var rows = new List<TraceRow>(visible.Count);
        foreach (var node in visible)
        {
            rows.Add(Row(node, ReferenceEquals(node, _tree.Root)));
        }

        _window.SetRows(rows, _tree.SelectedIndex);
    }

    private static TraceRow Row(PrecedentNode node, bool isRoot)
    {
        var item = node.Item;
        var glyph = item.Kind == PrecedentKind.MoreCells ? string.Empty
            : node.IsExpanded && node.Children.Count > 0 ? "▾"
            : node.CanExpand && !node.IsExpanded ? "▸"
            : string.Empty;
        var badges = new List<string>(2);
        if (node.IsCycle)
        {
            badges.Add("↻ circular");
        }

        if (item.HiddenNote is not null)
        {
            badges.Add("[" + item.HiddenNote + "]");
        }

        var dim = item.Kind == PrecedentKind.MoreCells || item.Kind == PrecedentKind.Truncated || item.Kind == PrecedentKind.Error;
        return new TraceRow(node.Depth, glyph, isRoot ? "⌂" : Icon(item.Kind), isRoot ? "Audited cell" : item.Kind.ToString(),
            item.Label, dim, string.Join(" ", badges), item.Argument ?? string.Empty, item.ValueText ?? string.Empty);
    }

    private static string Icon(PrecedentKind kind)
    {
        switch (kind)
        {
            case PrecedentKind.Cell:
                return "▫";
            case PrecedentKind.Range:
                return "▦";
            case PrecedentKind.Name:
                return "N";
            case PrecedentKind.Table:
                return "▤";
            case PrecedentKind.DynamicReference:
                return "ƒ";
            case PrecedentKind.Function:
                return "ƒx";
            case PrecedentKind.Group:
                return "(x)";
            case PrecedentKind.Error:
                return "⚠";
            case PrecedentKind.MoreCells:
                return "⋯";
            default:
                return "✂";
        }
    }

    // The footer: where the selection is ([Book]Sheet!A1 or the label), and its value.
    private static string Where(PrecedentNode node)
    {
        var item = node.Item;
        var place = item.Sheet is not null && item.Address is not null
            ? (item.Workbook is null ? string.Empty : "[" + item.Workbook + "]") + item.Sheet + "!" + item.Address
            : item.Label;
        return string.IsNullOrEmpty(item.ValueText) ? place : place + " = " + item.ValueText;
    }

    // A message in the window's footer. (Not the status bar: that needs the C API, so macro context, which navigation
    // avoids; see the remarks.)
    private void Message(string text) => _window.SetStatus(text);

    // True while Excel is editing a cell (Enter, Edit or Point mode); false if that cannot be read.
    private static bool IsEditing()
    {
        try
        {
            return ExcelDnaUtil.IsInFormulaEditMode();
        }
        catch (Exception)
        {
            return false;
        }
    }

    // Has the queue drained from Excel's message loop after the current handler returns (outside macro context): the
    // key hook's hidden window, or the window's dispatcher if the keyboard could not be connected. Never throws.
    private bool ScheduleDrain(Action drain)
    {
        if (TraceKeyHook.RunLater(drain))
        {
            return true;
        }

        try
        {
            _window.Dispatcher.BeginInvoke(DispatcherPriority.Background, drain);
            return true;
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("TraceWindowError", "could not post", ex.Message);
            return false;
        }
    }

    // The window handles of a workbook (one per window of it). Never throws.
    private static List<IntPtr> WindowsOf(object workbook)
    {
        var handles = new List<IntPtr>();
        try
        {
            dynamic book = workbook;
            object windowsObject = book.Windows;
            dynamic windows = windowsObject;
            int count = windows.Count;
            for (var i = 1; i <= count; i++)
            {
                object hwnd = windows.Item(i).Hwnd;
                handles.Add(new IntPtr(Convert.ToInt64(hwnd, CultureInfo.InvariantCulture)));
            }
        }
        catch (Exception)
        {
            // No windows to report.
        }

        return handles;
    }

    private void CloseWindow()
    {
        try
        {
            _window.CloseWindow();
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("TraceWindowError", "close failed", ex.Message);
            OnWindowClosed();
        }
    }

    // The window is gone, whatever closed it (us, or its owner window being destroyed): unhook and remember where it was.
    private void OnWindowClosed()
    {
        if (_closed)
        {
            return;
        }

        if (!_closing)
        {
            // Not closed by us: its owner window was destroyed (that workbook was closed).
            _closing = true;
            DiagnosticsLog.Write("TraceClose", "mode=window destroyed", "navigations=" + _navigations.ToString(CultureInfo.InvariantCulture));
        }

        _closed = true;
        _edit = null;
        _queue.Clear();
        if (ReferenceEquals(Current, this))
        {
            Current = null;
            TraceKeyHook.Release();
        }

        if (_window.LastBounds is WindowRect bounds)
        {
            _ui = _ui.WithBounds(bounds);
        }

        UiStateStore.Save(_ui);
    }

    private void LogNavigation(string key, string source, double sincePressMs, Stopwatch handled, TreeMove move, string target, string result) =>
        DiagnosticsLog.Write(
            "TraceNavigate",
            "key=" + key,
            "source=" + source,
            "ms=" + Ms(sincePressMs),
            "handleMs=" + Ms(handled.Elapsed.TotalMilliseconds),
            "move=" + move,
            "target=" + target,
            result);

    /// <summary>The handle of Excel's active workbook window (Excel is SDI: one top-level window per workbook), or zero.</summary>
    internal static IntPtr ActiveWindowHandle()
    {
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            object hwnd = app.ActiveWindow.Hwnd;
            return new IntPtr(Convert.ToInt64(hwnd, CultureInfo.InvariantCulture));
        }
        catch (Exception)
        {
            return IntPtr.Zero;
        }
    }

    private static double Elapsed(long pressed) =>
        (Stopwatch.GetTimestamp() - pressed) * 1000.0 / Stopwatch.Frequency;

    private static string Ms(double milliseconds) => milliseconds.ToString("0.0", CultureInfo.InvariantCulture);

    /// <summary>An F2 edit Excel is doing: the reference, its cell, the formula before, and where the selection was.</summary>
    private sealed class PendingEdit
    {
        public PendingEdit(ReferenceSpan span, object cell, string formula, IReadOnlyList<int> path)
        {
            Span = span;
            Cell = cell;
            Formula = formula;
            Path = path;
        }

        public ReferenceSpan Span { get; }

        public object Cell { get; }

        public string Formula { get; }

        public IReadOnlyList<int> Path { get; }

        public Stopwatch Started { get; } = Stopwatch.StartNew();
    }
}
