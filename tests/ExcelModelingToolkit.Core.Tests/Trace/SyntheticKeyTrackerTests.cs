using System;
using ExcelModelingToolkit.Core.Trace;
using Xunit;

namespace ExcelModelingToolkit.Core.Tests.Trace;

public class SyntheticKeyTrackerTests
{
    private const int F2 = 0x71;
    private const int Right = 0x27;
    private const int Shift = 0x10;
    private const int Down = 0x28;

    private static SyntheticKeyTracker Tracker() => new SyntheticKeyTracker(new[]
    {
        SyntheticKey.Down(F2), SyntheticKey.Up(F2), SyntheticKey.Down(Shift), SyntheticKey.Down(Right),
        SyntheticKey.Up(Right), SyntheticKey.Up(Shift), SyntheticKey.Down(F2), SyntheticKey.Up(F2),
    });

    [Fact]
    public void The_sequence_is_counted_off_in_order_and_ends_at_its_last_key()
    {
        var tracker = Tracker();

        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(F2, false));
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(F2, true));
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(Shift, false));
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(Right, false));
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(Right, true));
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(Shift, true));
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(F2, false));
        Assert.False(tracker.IsOver);
        Assert.Equal(SyntheticKeyMatch.Completed, tracker.Observe(F2, true));
        Assert.True(tracker.IsOver);
        Assert.Equal(8, tracker.Seen);
        Assert.Equal(SyntheticKeyMatch.Unexpected, tracker.Observe(F2, false));
    }

    [Fact]
    public void A_stray_release_before_or_inside_the_sequence_does_not_end_it()
    {
        // The user lets go of the F2 that started it just before the sent keys arrive.
        var tracker = Tracker();

        Assert.Equal(SyntheticKeyMatch.Stray, tracker.Observe(F2, true));
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(F2, false));
        Assert.Equal(SyntheticKeyMatch.Stray, tracker.Observe(Down, true));
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(F2, true));
        Assert.Equal(2, tracker.Seen);
        Assert.False(tracker.IsOver);
    }

    [Fact]
    public void A_press_that_is_not_the_next_key_ends_the_sequence()
    {
        var tracker = Tracker();
        tracker.Observe(F2, false);

        Assert.Equal(SyntheticKeyMatch.Unexpected, tracker.Observe(Down, false));
        Assert.True(tracker.IsOver);
        Assert.Equal(SyntheticKeyMatch.Unexpected, tracker.Observe(F2, true));
    }

    [Fact]
    public void A_sequence_has_keys()
    {
        Assert.Throws<ArgumentException>(() => new SyntheticKeyTracker(new SyntheticKey[0]));
        Assert.Throws<ArgumentNullException>(() => new SyntheticKeyTracker(null!));
    }

    [Fact]
    public void Keys_compare_by_value()
    {
        Assert.Equal(SyntheticKey.Down(F2), new SyntheticKey(F2, false));
        Assert.NotEqual(SyntheticKey.Down(F2), SyntheticKey.Up(F2));
        Assert.Equal("0x71 up", SyntheticKey.Up(F2).ToString());
    }
}
