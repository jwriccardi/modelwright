using System;
using System.Linq;
using ExcelModelingToolkit.Core.Keys;
using ExcelModelingToolkit.Core.Trace;
using Xunit;

namespace ExcelModelingToolkit.Core.Tests.Trace;

public class TraceKeysTests
{
    private const int Enter = 0x0D;
    private const int Escape = 0x1B;
    private const int End = 0x23;
    private const int Home = 0x24;
    private const int Left = 0x25;
    private const int Up = 0x26;
    private const int Right = 0x27;
    private const int Down = 0x28;
    private const int E = 0x45;
    private const int F2 = 0x71;
    private const int Z = 0x5A;
    private const int Tab = 0x09;
    private const int OemOpenBracket = 0xDB;

    [Theory]
    [InlineData(Up, KeyModifiers.None, TraceKeyCommand.Up)]
    [InlineData(Down, KeyModifiers.None, TraceKeyCommand.Down)]
    [InlineData(Left, KeyModifiers.None, TraceKeyCommand.Left)]
    [InlineData(Right, KeyModifiers.None, TraceKeyCommand.Right)]
    [InlineData(Enter, KeyModifiers.None, TraceKeyCommand.Close)]
    [InlineData(Escape, KeyModifiers.None, TraceKeyCommand.Cancel)]
    [InlineData(E, KeyModifiers.Ctrl, TraceKeyCommand.ToggleEvaluate)]
    [InlineData(Up, KeyModifiers.Ctrl, TraceKeyCommand.MoveWindowUp)]
    [InlineData(Down, KeyModifiers.Ctrl, TraceKeyCommand.MoveWindowDown)]
    [InlineData(Left, KeyModifiers.Ctrl, TraceKeyCommand.MoveWindowLeft)]
    [InlineData(Right, KeyModifiers.Ctrl, TraceKeyCommand.MoveWindowRight)]
    [InlineData(Home, KeyModifiers.Ctrl, TraceKeyCommand.SnapTopLeft)]
    [InlineData(End, KeyModifiers.Ctrl, TraceKeyCommand.SnapBottomRight)]
    [InlineData(Up, KeyModifiers.Shift, TraceKeyCommand.ShrinkHeight)]
    [InlineData(Down, KeyModifiers.Shift, TraceKeyCommand.GrowHeight)]
    [InlineData(Left, KeyModifiers.Shift, TraceKeyCommand.ShrinkWidth)]
    [InlineData(Right, KeyModifiers.Shift, TraceKeyCommand.GrowWidth)]
    public void Trace_keys_map_to_their_commands(int key, KeyModifiers modifiers, TraceKeyCommand expected)
    {
        Assert.Equal(expected, TraceKeys.Command(key, modifiers));
    }

    [Theory]
    [InlineData(F2, KeyModifiers.None)] // edits the active cell in Excel; the window stays open
    [InlineData(Z, KeyModifiers.Ctrl)] // undo stays Excel's (and the undo hook's)
    [InlineData(OemOpenBracket, KeyModifiers.Ctrl | KeyModifiers.Shift)] // Ctrl+Shift+[ re-traces through OnKey
    [InlineData(Tab, KeyModifiers.None)]
    [InlineData(Home, KeyModifiers.None)]
    [InlineData(End, KeyModifiers.None)]
    [InlineData(E, KeyModifiers.None)] // typing starts an edit in Excel
    [InlineData(Down, KeyModifiers.Ctrl | KeyModifiers.Shift)]
    [InlineData(Down, KeyModifiers.Alt)]
    [InlineData(Enter, KeyModifiers.Shift)]
    [InlineData(Escape, KeyModifiers.Ctrl)]
    [InlineData(E, KeyModifiers.Ctrl | KeyModifiers.Alt)]
    public void Every_other_key_goes_to_excel(int key, KeyModifiers modifiers)
    {
        Assert.Equal(TraceKeyCommand.None, TraceKeys.Command(key, modifiers));
    }

    [Fact]
    public void Keys_are_taken_only_while_excel_is_ready()
    {
        foreach (var command in Enum.GetValues(typeof(TraceKeyCommand)).Cast<TraceKeyCommand>())
        {
            Assert.Equal(command != TraceKeyCommand.None, TraceKeys.Takes(command, TraceKeyContext.Ready));
            Assert.False(TraceKeys.Takes(command, TraceKeyContext.Editing));
            Assert.False(TraceKeys.Takes(command, TraceKeyContext.Elsewhere));
        }
    }

    [Theory]
    [InlineData(false, false, true, TraceKeyContext.Ready)]
    [InlineData(false, true, true, TraceKeyContext.Editing)] // after F2, typing, or a click in the formula bar
    [InlineData(false, true, false, TraceKeyContext.Editing)]
    [InlineData(false, false, false, TraceKeyContext.Elsewhere)] // a dialog, the Name Box, a task pane
    [InlineData(true, false, true, TraceKeyContext.Elsewhere)] // a menu, a drag, a move or size loop
    [InlineData(true, true, true, TraceKeyContext.Elsewhere)]
    public void The_context_comes_from_what_the_hook_sees(bool modalLoop, bool editing, bool focusOnGrid, TraceKeyContext expected)
    {
        Assert.Equal(expected, TraceKeys.Context(modalLoop, editing, focusOnGrid));
    }

    [Fact]
    public void Arrows_repeat_while_held_and_the_other_commands_do_not()
    {
        Assert.True(TraceKeys.Repeats(TraceKeyCommand.Down));
        Assert.True(TraceKeys.Repeats(TraceKeyCommand.Up));
        Assert.True(TraceKeys.Repeats(TraceKeyCommand.Left));
        Assert.True(TraceKeys.Repeats(TraceKeyCommand.Right));
        Assert.True(TraceKeys.Repeats(TraceKeyCommand.MoveWindowLeft));
        Assert.True(TraceKeys.Repeats(TraceKeyCommand.GrowHeight));
        Assert.False(TraceKeys.Repeats(TraceKeyCommand.Close));
        Assert.False(TraceKeys.Repeats(TraceKeyCommand.Cancel));
        Assert.False(TraceKeys.Repeats(TraceKeyCommand.ToggleEvaluate));
        Assert.False(TraceKeys.Repeats(TraceKeyCommand.SnapTopLeft));
        Assert.False(TraceKeys.Repeats(TraceKeyCommand.SnapBottomRight));
        Assert.False(TraceKeys.Repeats(TraceKeyCommand.None));
    }

    [Fact]
    public void Window_commands_need_no_excel_call()
    {
        var window = Enum.GetValues(typeof(TraceKeyCommand)).Cast<TraceKeyCommand>().Where(TraceKeys.IsWindowCommand).ToArray();

        Assert.Equal(
            new[]
            {
                TraceKeyCommand.MoveWindowUp, TraceKeyCommand.MoveWindowDown, TraceKeyCommand.MoveWindowLeft,
                TraceKeyCommand.MoveWindowRight, TraceKeyCommand.SnapTopLeft, TraceKeyCommand.SnapBottomRight,
                TraceKeyCommand.ShrinkHeight, TraceKeyCommand.GrowHeight, TraceKeyCommand.ShrinkWidth, TraceKeyCommand.GrowWidth,
            },
            window);
    }

    [Fact]
    public void Enter_stays_and_escape_returns_to_the_audited_cell()
    {
        Assert.Equal(TraceCloseMode.StayOnCurrentCell, TraceKeys.CloseMode(TraceKeyCommand.Close));
        Assert.Equal(TraceCloseMode.ReturnToAuditedCell, TraceKeys.CancelCloseMode);
        Assert.Equal(TraceKeys.CancelCloseMode, TraceKeys.CloseMode(TraceKeyCommand.Cancel));
        Assert.Null(TraceKeys.CloseMode(TraceKeyCommand.Down));
    }
}
