using System;
using System.Collections.Generic;

namespace Modelwright.Core.Trace;

/// <summary>
/// Where the Trace In window goes: its default place, the keyboard moves, snaps and resizes (Ctrl+arrows,
/// Ctrl+Home, Ctrl+End, Shift+arrows; research/07), and recovery of a window left off-screen. Sizes are in screen
/// pixels; <c>scale</c> is the window's DPI over 96, so steps and minimum sizes look the same on every monitor.
/// </summary>
public static class TraceWindowGeometry
{
    /// <summary>Pixels (at 96 DPI) a Ctrl+arrow moves the window or a Shift+arrow resizes it.</summary>
    public const int Step = 24;

    /// <summary>The smallest width (at 96 DPI) Shift+Left or a remembered size can give.</summary>
    public const int MinWidth = 320;

    /// <summary>The smallest height (at 96 DPI) Shift+Up or a remembered size can give.</summary>
    public const int MinHeight = 180;

    /// <summary>The default width at 96 DPI.</summary>
    public const int DefaultWidth = 620;

    /// <summary>The default height at 96 DPI.</summary>
    public const int DefaultHeight = 380;

    // The strip along the top of the window that holds the title bar (at 96 DPI), and how much of it, across, must
    // be on some monitor for the window to count as visible: enough to drag it back.
    private const int TitleHeight = 24;
    private const int MinVisibleWidth = 64;

    /// <summary>
    /// The window's bounds after <paramref name="command"/>; unchanged for a command that is not a window command.
    /// </summary>
    /// <param name="command">A move, snap or resize command (<see cref="TraceKeys.IsWindowCommand"/>).</param>
    /// <param name="window">The window's bounds now.</param>
    /// <param name="workArea">The work area of the monitor the window is on (snaps go there; sizes are capped by it).</param>
    /// <param name="workAreas">Every monitor's work area: a move that would leave the title bar on none is refused.</param>
    /// <param name="scale">The window's DPI over 96.</param>
    /// <exception cref="ArgumentNullException"><paramref name="workAreas"/> is null.</exception>
    public static WindowRect Apply(TraceKeyCommand command, WindowRect window, WindowRect workArea,
        IReadOnlyList<WindowRect> workAreas, double scale)
    {
        if (workAreas is null)
        {
            throw new ArgumentNullException(nameof(workAreas));
        }

        var step = Scaled(Step, scale);
        switch (command)
        {
            case TraceKeyCommand.MoveWindowUp:
                return MoveBy(window, 0, -step, workAreas, scale);
            case TraceKeyCommand.MoveWindowDown:
                return MoveBy(window, 0, step, workAreas, scale);
            case TraceKeyCommand.MoveWindowLeft:
                return MoveBy(window, -step, 0, workAreas, scale);
            case TraceKeyCommand.MoveWindowRight:
                return MoveBy(window, step, 0, workAreas, scale);
            case TraceKeyCommand.SnapTopLeft:
            {
                var fitted = FitSize(window, workArea, scale);
                return fitted.MoveTo(workArea.Left, workArea.Top);
            }

            case TraceKeyCommand.SnapBottomRight:
            {
                var fitted = FitSize(window, workArea, scale);
                return fitted.MoveTo(workArea.Right - fitted.Width, workArea.Bottom - fitted.Height);
            }

            case TraceKeyCommand.ShrinkHeight:
                return Resize(window, 0, -step, workArea, scale);
            case TraceKeyCommand.GrowHeight:
                return Resize(window, 0, step, workArea, scale);
            case TraceKeyCommand.ShrinkWidth:
                return Resize(window, -step, 0, workArea, scale);
            case TraceKeyCommand.GrowWidth:
                return Resize(window, step, 0, workArea, scale);
            default:
                return window;
        }
    }

    /// <summary>
    /// True if enough of the window's title bar is on one of <paramref name="workAreas"/> to drag it: at least
    /// 64 pixels across (at 96 DPI; the whole width if narrower) and half the title bar's height.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="workAreas"/> is null.</exception>
    public static bool IsVisibleEnough(WindowRect window, IReadOnlyList<WindowRect> workAreas, double scale)
    {
        if (workAreas is null)
        {
            throw new ArgumentNullException(nameof(workAreas));
        }

        var title = new WindowRect(window.Left, window.Top, window.Width, Scaled(TitleHeight, scale));
        var needWidth = Math.Max(1, Math.Min(Scaled(MinVisibleWidth, scale), window.Width));
        var needHeight = Math.Max(1, title.Height / 2);
        foreach (var area in workAreas)
        {
            var width = Math.Min(title.Right, area.Right) - Math.Max(title.Left, area.Left);
            var height = Math.Min(title.Bottom, area.Bottom) - Math.Max(title.Top, area.Top);
            if (width >= needWidth && height >= needHeight)
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// A remembered window brought back on screen if needed (a monitor was unplugged, or the resolution changed): if
    /// its title bar is visible enough (<see cref="IsVisibleEnough"/>) it is kept, at least the minimum size;
    /// otherwise it is moved into the work area it overlaps most (or <paramref name="fallbackArea"/>, the Excel
    /// window's monitor, if it overlaps none) and shrunk to fit there.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="workAreas"/> is null.</exception>
    public static WindowRect EnsureVisible(WindowRect window, IReadOnlyList<WindowRect> workAreas, WindowRect fallbackArea, double scale)
    {
        if (workAreas is null)
        {
            throw new ArgumentNullException(nameof(workAreas));
        }

        var sized = new WindowRect(window.Left, window.Top,
            Math.Max(window.Width, Scaled(MinWidth, scale)), Math.Max(window.Height, Scaled(MinHeight, scale)));
        if (IsVisibleEnough(sized, workAreas, scale))
        {
            return sized;
        }

        var target = fallbackArea;
        long best = 0;
        foreach (var area in workAreas)
        {
            var overlap = sized.OverlapArea(area);
            if (overlap > best)
            {
                best = overlap;
                target = area;
            }
        }

        var fitted = FitSize(sized, target, scale);
        return fitted.MoveTo(
            Clamp(fitted.Left, target.Left, target.Right - fitted.Width),
            Clamp(fitted.Top, target.Top, target.Bottom - fitted.Height));
    }

    /// <summary>
    /// Where the window opens the first time: the default size (fitted to the work area), against the right-hand
    /// side of the Excel window, a little below its top, kept on <paramref name="workArea"/>.
    /// </summary>
    public static WindowRect DefaultBounds(WindowRect excelWindow, WindowRect workArea, double scale)
    {
        var fitted = FitSize(new WindowRect(0, 0, Scaled(DefaultWidth, scale), Scaled(DefaultHeight, scale)), workArea, scale);
        var margin = Scaled(Step, scale);
        var left = excelWindow.Right - fitted.Width - margin;
        var top = excelWindow.Top + Scaled(160, scale);
        return fitted.MoveTo(
            Clamp(left, workArea.Left, workArea.Right - fitted.Width),
            Clamp(top, workArea.Top, workArea.Bottom - fitted.Height));
    }

    private static WindowRect MoveBy(WindowRect window, int dx, int dy, IReadOnlyList<WindowRect> workAreas, double scale)
    {
        var moved = window.MoveTo(window.Left + dx, window.Top + dy);
        return IsVisibleEnough(moved, workAreas, scale) ? moved : window;
    }

    private static WindowRect Resize(WindowRect window, int dw, int dh, WindowRect workArea, double scale)
    {
        var width = Clamp(window.Width + dw, Scaled(MinWidth, scale), Math.Max(window.Width, workArea.Width));
        var height = Clamp(window.Height + dh, Scaled(MinHeight, scale), Math.Max(window.Height, workArea.Height));
        return new WindowRect(window.Left, window.Top, width, height);
    }

    // The window's size capped by the work area (never below 1 pixel), position unchanged.
    private static WindowRect FitSize(WindowRect window, WindowRect area, double scale) =>
        new WindowRect(window.Left, window.Top,
            Math.Max(1, Math.Min(Math.Max(window.Width, Scaled(MinWidth, scale)), area.Width)),
            Math.Max(1, Math.Min(Math.Max(window.Height, Scaled(MinHeight, scale)), area.Height)));

    private static int Clamp(int value, int min, int max) => max < min ? min : Math.Min(Math.Max(value, min), max);

    private static int Scaled(int pixels, double scale) =>
        (int)Math.Round(pixels * (scale > 0 && !double.IsNaN(scale) && !double.IsInfinity(scale) ? scale : 1.0));
}
