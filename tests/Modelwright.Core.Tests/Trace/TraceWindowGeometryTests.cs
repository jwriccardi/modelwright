using System;
using Modelwright.Core.Trace;
using Xunit;

namespace Modelwright.Core.Tests.Trace;

public class TraceWindowGeometryTests
{
    // Two 1920x1080 monitors side by side, each with a 40-pixel taskbar at the bottom.
    private static readonly WindowRect Primary = new WindowRect(0, 0, 1920, 1040);
    private static readonly WindowRect Secondary = new WindowRect(1920, 0, 1920, 1040);
    private static readonly WindowRect[] Monitors = { Primary, Secondary };
    private static readonly WindowRect Window = new WindowRect(400, 300, 620, 380);

    private static WindowRect Apply(TraceKeyCommand command, WindowRect window, double scale = 1.0) =>
        TraceWindowGeometry.Apply(command, window, Primary, Monitors, scale);

    [Fact]
    public void Ctrl_arrows_move_by_one_step()
    {
        Assert.Equal(new WindowRect(400, 276, 620, 380), Apply(TraceKeyCommand.MoveWindowUp, Window));
        Assert.Equal(new WindowRect(400, 324, 620, 380), Apply(TraceKeyCommand.MoveWindowDown, Window));
        Assert.Equal(new WindowRect(376, 300, 620, 380), Apply(TraceKeyCommand.MoveWindowLeft, Window));
        Assert.Equal(new WindowRect(424, 300, 620, 380), Apply(TraceKeyCommand.MoveWindowRight, Window));
    }

    [Fact]
    public void Steps_scale_with_the_dpi()
    {
        Assert.Equal(new WindowRect(448, 300, 620, 380), Apply(TraceKeyCommand.MoveWindowRight, Window, 2.0));
        // At 200% the minimum width is 640, so the 620-pixel window widens to it.
        Assert.Equal(new WindowRect(400, 300, 640, 428), Apply(TraceKeyCommand.GrowHeight, Window, 2.0));
    }

    [Fact]
    public void A_move_onto_another_monitor_is_allowed()
    {
        var nearEdge = new WindowRect(1900, 300, 620, 380);

        Assert.Equal(new WindowRect(1924, 300, 620, 380), Apply(TraceKeyCommand.MoveWindowRight, nearEdge));
    }

    [Fact]
    public void A_move_that_would_lose_the_title_bar_is_refused()
    {
        var atTop = new WindowRect(400, 0, 620, 380);
        var atRight = new WindowRect(3840 - 64, 300, 620, 380);

        Assert.Equal(atTop, Apply(TraceKeyCommand.MoveWindowUp, atTop));
        Assert.Equal(atRight, Apply(TraceKeyCommand.MoveWindowRight, atRight));
        Assert.Equal(new WindowRect(3840 - 88, 300, 620, 380), Apply(TraceKeyCommand.MoveWindowLeft, atRight));
    }

    [Fact]
    public void Ctrl_home_and_end_snap_to_the_work_area_corners()
    {
        Assert.Equal(new WindowRect(0, 0, 620, 380), Apply(TraceKeyCommand.SnapTopLeft, Window));
        Assert.Equal(new WindowRect(1300, 660, 620, 380), Apply(TraceKeyCommand.SnapBottomRight, Window));
    }

    [Fact]
    public void Snapping_recovers_a_window_left_off_screen_and_fits_it()
    {
        var lost = new WindowRect(-5000, 9000, 2500, 1200);

        Assert.Equal(new WindowRect(0, 0, 1920, 1040), Apply(TraceKeyCommand.SnapTopLeft, lost));
        Assert.Equal(new WindowRect(0, 0, 1920, 1040), Apply(TraceKeyCommand.SnapBottomRight, lost));
    }

    [Fact]
    public void Shift_arrows_resize_from_the_top_left_corner()
    {
        Assert.Equal(new WindowRect(400, 300, 620, 356), Apply(TraceKeyCommand.ShrinkHeight, Window));
        Assert.Equal(new WindowRect(400, 300, 620, 404), Apply(TraceKeyCommand.GrowHeight, Window));
        Assert.Equal(new WindowRect(400, 300, 596, 380), Apply(TraceKeyCommand.ShrinkWidth, Window));
        Assert.Equal(new WindowRect(400, 300, 644, 380), Apply(TraceKeyCommand.GrowWidth, Window));
    }

    [Fact]
    public void Resizing_stops_at_the_minimum_and_at_the_work_area()
    {
        var small = new WindowRect(0, 0, TraceWindowGeometry.MinWidth, TraceWindowGeometry.MinHeight);
        var full = new WindowRect(0, 0, 1920, 1040);

        Assert.Equal(small, Apply(TraceKeyCommand.ShrinkWidth, small));
        Assert.Equal(small, Apply(TraceKeyCommand.ShrinkHeight, small));
        Assert.Equal(full, Apply(TraceKeyCommand.GrowWidth, full));
        Assert.Equal(full, Apply(TraceKeyCommand.GrowHeight, full));
    }

    [Fact]
    public void Other_commands_leave_the_window_alone()
    {
        Assert.Equal(Window, Apply(TraceKeyCommand.Down, Window));
        Assert.Equal(Window, Apply(TraceKeyCommand.None, Window));
        Assert.Throws<ArgumentNullException>(() => TraceWindowGeometry.Apply(TraceKeyCommand.Up, Window, Primary, null!, 1.0));
    }

    [Fact]
    public void A_visible_window_is_kept_where_it_was()
    {
        Assert.True(TraceWindowGeometry.IsVisibleEnough(Window, Monitors, 1.0));
        Assert.Equal(Window, TraceWindowGeometry.EnsureVisible(Window, Monitors, Primary, 1.0));

        var spanning = new WindowRect(1700, 100, 620, 380);
        Assert.Equal(spanning, TraceWindowGeometry.EnsureVisible(spanning, Monitors, Primary, 1.0));
    }

    [Fact]
    public void A_window_on_an_unplugged_monitor_comes_back_to_the_fallback()
    {
        var onMissingMonitor = new WindowRect(4200, 300, 620, 380);

        Assert.False(TraceWindowGeometry.IsVisibleEnough(onMissingMonitor, Monitors, 1.0));
        Assert.Equal(new WindowRect(1300, 300, 620, 380), TraceWindowGeometry.EnsureVisible(onMissingMonitor, new[] { Primary }, Primary, 1.0));
    }

    [Fact]
    public void A_window_whose_title_bar_is_above_the_screen_is_pulled_into_the_monitor_it_overlaps()
    {
        var titleAbove = new WindowRect(2400, -200, 620, 380);

        Assert.Equal(new WindowRect(2400, 0, 620, 380), TraceWindowGeometry.EnsureVisible(titleAbove, Monitors, Primary, 1.0));
    }

    [Fact]
    public void A_remembered_size_below_the_minimum_is_raised()
    {
        var tiny = new WindowRect(100, 100, 10, 10);

        Assert.Equal(
            new WindowRect(100, 100, TraceWindowGeometry.MinWidth, TraceWindowGeometry.MinHeight),
            TraceWindowGeometry.EnsureVisible(tiny, Monitors, Primary, 1.0));
    }

    [Fact]
    public void The_default_place_is_against_the_right_of_the_excel_window()
    {
        var excel = new WindowRect(0, 0, 1920, 1040);

        Assert.Equal(new WindowRect(1276, 160, 620, 380), TraceWindowGeometry.DefaultBounds(excel, Primary, 1.0));
    }

    [Fact]
    public void The_default_place_fits_a_small_screen()
    {
        var small = new WindowRect(0, 0, 500, 300);

        Assert.Equal(new WindowRect(0, 0, 500, 300), TraceWindowGeometry.DefaultBounds(small, small, 1.0));
    }

    [Fact]
    public void Rects_compare_by_value()
    {
        var rect = new WindowRect(1, 2, 3, 4);

        Assert.Equal(4, rect.Right);
        Assert.Equal(6, rect.Bottom);
        Assert.True(rect == new WindowRect(1, 2, 3, 4));
        Assert.True(rect != new WindowRect(1, 2, 3, 5));
        Assert.Equal(rect.GetHashCode(), new WindowRect(1, 2, 3, 4).GetHashCode());
        Assert.False(rect.Equals((object)"x"));
        Assert.Equal("(1,2) 3x4", rect.ToString());
        Assert.Equal(0, new WindowRect(0, 0, -5, -5).Width);
        Assert.Equal(0, rect.OverlapArea(new WindowRect(100, 100, 5, 5)));
        Assert.Equal(2, rect.OverlapArea(new WindowRect(3, 4, 5, 5)));
    }
}
