using ExcelModelingToolkit.Core.Keys;

namespace ExcelModelingToolkit.Core.Trace;

/// <summary>What a key does while the Trace In window is open (docs/PLAN.md section 4.5, research/07).</summary>
public enum TraceKeyCommand
{
    /// <summary>Not a Trace In key: it goes to Excel.</summary>
    None,

    /// <summary>Up: select the previous row and go to it.</summary>
    Up,

    /// <summary>Down: select the next row and go to it.</summary>
    Down,

    /// <summary>Left: collapse, or go up a level.</summary>
    Left,

    /// <summary>Right: expand (loading the children), or go to the first child.</summary>
    Right,

    /// <summary>Enter (OK): close, staying on the current cell.</summary>
    Close,

    /// <summary>Esc (Cancel, the title bar's close button): close as <see cref="TraceKeys.CancelCloseMode"/> says.</summary>
    Cancel,

    /// <summary>Ctrl+E: "Evaluate functions &amp; groups" (not in this version: says so).</summary>
    ToggleEvaluate,

    /// <summary>Ctrl+Up: move the window up.</summary>
    MoveWindowUp,

    /// <summary>Ctrl+Down: move the window down.</summary>
    MoveWindowDown,

    /// <summary>Ctrl+Left: move the window left.</summary>
    MoveWindowLeft,

    /// <summary>Ctrl+Right: move the window right.</summary>
    MoveWindowRight,

    /// <summary>Ctrl+Home: snap the window to the top-left of its monitor's work area.</summary>
    SnapTopLeft,

    /// <summary>Ctrl+End: snap the window to the bottom-right of its monitor's work area.</summary>
    SnapBottomRight,

    /// <summary>Shift+Up: make the window shorter.</summary>
    ShrinkHeight,

    /// <summary>Shift+Down: make the window taller.</summary>
    GrowHeight,

    /// <summary>Shift+Left: make the window narrower.</summary>
    ShrinkWidth,

    /// <summary>Shift+Right: make the window wider.</summary>
    GrowWidth,
}

/// <summary>Where Excel's selection is left when the Trace In window closes.</summary>
public enum TraceCloseMode
{
    /// <summary>On the cell the last navigation went to.</summary>
    StayOnCurrentCell,

    /// <summary>Back on the audited cell.</summary>
    ReturnToAuditedCell,
}

/// <summary>Whether Excel can hand a key to Trace In right now.</summary>
public enum TraceKeyContext
{
    /// <summary>A worksheet grid (or the Trace In window) has the focus and no cell is being edited.</summary>
    Ready,

    /// <summary>A cell is being edited (Enter, Edit or Point mode, e.g. after F2): keys belong to Excel.</summary>
    Editing,

    /// <summary>The focus is elsewhere (a dialog, the formula bar, a task pane), a menu is open, or the mouse is captured.</summary>
    Elsewhere,
}

/// <summary>
/// The Trace In keyboard (variant C of spike K4): which presses the thread keyboard hook takes while the window is
/// open, and what they do. Every other key, and every key while a cell is being edited or the focus is not on a
/// worksheet grid, goes to Excel untouched; so F2 edits the active cell with the window open.
/// </summary>
public static class TraceKeys
{
    /// <summary>
    /// What Esc, Cancel and the title bar's close button do. The spec's default is to return to the audited cell
    /// (docs/PLAN.md section 4.5); what Macabacus does is still to be confirmed (docs/open-questions.md, item 7).
    /// Change it here only.
    /// </summary>
    public const TraceCloseMode CancelCloseMode = TraceCloseMode.ReturnToAuditedCell;

    /// <summary>Virtual-key codes the table uses (Win32 <c>VK_*</c>).</summary>
    private const int VkReturn = 0x0D;
    private const int VkEscape = 0x1B;
    private const int VkEnd = 0x23;
    private const int VkHome = 0x24;
    private const int VkLeft = 0x25;
    private const int VkUp = 0x26;
    private const int VkRight = 0x27;
    private const int VkDown = 0x28;
    private const int VkE = 0x45;

    /// <summary>
    /// The command for a press of <paramref name="virtualKey"/> with exactly <paramref name="modifiers"/> held, or
    /// <see cref="TraceKeyCommand.None"/>: arrows, Enter and Esc alone; Ctrl with an arrow, Home, End or E; Shift with
    /// an arrow. Anything with Alt, or with Ctrl and Shift together, is not ours.
    /// </summary>
    public static TraceKeyCommand Command(int virtualKey, KeyModifiers modifiers)
    {
        switch (modifiers)
        {
            case KeyModifiers.None:
                switch (virtualKey)
                {
                    case VkUp:
                        return TraceKeyCommand.Up;
                    case VkDown:
                        return TraceKeyCommand.Down;
                    case VkLeft:
                        return TraceKeyCommand.Left;
                    case VkRight:
                        return TraceKeyCommand.Right;
                    case VkReturn:
                        return TraceKeyCommand.Close;
                    case VkEscape:
                        return TraceKeyCommand.Cancel;
                }

                break;
            case KeyModifiers.Ctrl:
                switch (virtualKey)
                {
                    case VkUp:
                        return TraceKeyCommand.MoveWindowUp;
                    case VkDown:
                        return TraceKeyCommand.MoveWindowDown;
                    case VkLeft:
                        return TraceKeyCommand.MoveWindowLeft;
                    case VkRight:
                        return TraceKeyCommand.MoveWindowRight;
                    case VkHome:
                        return TraceKeyCommand.SnapTopLeft;
                    case VkEnd:
                        return TraceKeyCommand.SnapBottomRight;
                    case VkE:
                        return TraceKeyCommand.ToggleEvaluate;
                }

                break;
            case KeyModifiers.Shift:
                switch (virtualKey)
                {
                    case VkUp:
                        return TraceKeyCommand.ShrinkHeight;
                    case VkDown:
                        return TraceKeyCommand.GrowHeight;
                    case VkLeft:
                        return TraceKeyCommand.ShrinkWidth;
                    case VkRight:
                        return TraceKeyCommand.GrowWidth;
                }

                break;
        }

        return TraceKeyCommand.None;
    }

    /// <summary>
    /// Whether Excel can hand keys to Trace In, from what the hook observes at the press, in this order: the mouse
    /// is captured, or a menu is open or a window is being moved or sized (Excel is in a modal loop):
    /// <see cref="TraceKeyContext.Elsewhere"/>; a cell is being edited (Excel's formula-edit state, which F2,
    /// typing, a double-click or a click in the formula bar all enter): <see cref="TraceKeyContext.Editing"/>; the
    /// focus is not on a worksheet grid or the Trace In window (a dialog, the Name Box, a task pane):
    /// <see cref="TraceKeyContext.Elsewhere"/>; else <see cref="TraceKeyContext.Ready"/>. The edit state is read at
    /// every press rather than tracked from F2, so an edit started or ended any other way is seen too.
    /// </summary>
    public static TraceKeyContext Context(bool modalLoop, bool editing, bool focusOnGridOrTraceWindow) =>
        modalLoop ? TraceKeyContext.Elsewhere
        : editing ? TraceKeyContext.Editing
        : focusOnGridOrTraceWindow ? TraceKeyContext.Ready
        : TraceKeyContext.Elsewhere;

    /// <summary>
    /// True if the hook takes the key (swallowing it and its key-up): a Trace In command, while Excel is
    /// <see cref="TraceKeyContext.Ready"/>.
    /// </summary>
    public static bool Takes(TraceKeyCommand command, TraceKeyContext context) =>
        command != TraceKeyCommand.None && context == TraceKeyContext.Ready;

    /// <summary>
    /// True if holding the key down repeats the command (the arrows, alone or with Ctrl or Shift). For the others
    /// (Enter, Esc, Ctrl+Home, Ctrl+End, Ctrl+E) the auto-repeats are swallowed with the press.
    /// </summary>
    public static bool Repeats(TraceKeyCommand command)
    {
        switch (command)
        {
            case TraceKeyCommand.None:
            case TraceKeyCommand.Close:
            case TraceKeyCommand.Cancel:
            case TraceKeyCommand.ToggleEvaluate:
            case TraceKeyCommand.SnapTopLeft:
            case TraceKeyCommand.SnapBottomRight:
                return false;
            default:
                return true;
        }
    }

    /// <summary>True for the commands that only move, snap or resize the window (no Excel call is needed).</summary>
    public static bool IsWindowCommand(TraceKeyCommand command) =>
        command >= TraceKeyCommand.MoveWindowUp && command <= TraceKeyCommand.GrowWidth;

    /// <summary>
    /// Where the selection is left by a closing command: <see cref="TraceKeyCommand.Close"/> stays on the current
    /// cell, <see cref="TraceKeyCommand.Cancel"/> does what <see cref="CancelCloseMode"/> says. Null for a command
    /// that does not close the window.
    /// </summary>
    public static TraceCloseMode? CloseMode(TraceKeyCommand command) =>
        command == TraceKeyCommand.Close ? TraceCloseMode.StayOnCurrentCell
        : command == TraceKeyCommand.Cancel ? CancelCloseMode
        : (TraceCloseMode?)null;
}
