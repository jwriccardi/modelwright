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
        Assert.False(SyntheticKey.Down(F2).IsCharacter);
    }

    [Fact]
    public void A_typed_character_is_a_VK_PACKET_event_that_knows_its_character()
    {
        var a = SyntheticKey.CharacterDown('A');

        Assert.True(a.IsCharacter);
        Assert.Equal(SyntheticKey.VkPacket, a.VirtualKey);
        Assert.False(a.KeyUp);
        Assert.True(SyntheticKey.CharacterUp('A').KeyUp);
        Assert.Equal(SyntheticKey.CharacterDown('A'), a);
        Assert.NotEqual(SyntheticKey.CharacterDown('B'), a);
        Assert.NotEqual(SyntheticKey.Down(SyntheticKey.VkPacket), a);
        Assert.Equal("'A' down", a.ToString());
        Assert.Throws<ArgumentOutOfRangeException>(() => SyntheticKey.CharacterDown('\0'));
    }

    [Fact]
    public void Typed_characters_are_counted_off_as_VK_PACKET_presses_and_releases()
    {
        // The Go To step: F5, 'A2' typed, Enter. The hook sees each character as VK_PACKET, not which one it is.
        const int F5 = 0x74;
        const int Enter = 0x0D;
        var tracker = new SyntheticKeyTracker(new[]
        {
            SyntheticKey.Down(F5), SyntheticKey.Up(F5), SyntheticKey.CharacterDown('A'), SyntheticKey.CharacterUp('A'),
            SyntheticKey.CharacterDown('2'), SyntheticKey.CharacterUp('2'), SyntheticKey.Down(Enter), SyntheticKey.Up(Enter),
        });

        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(F5, false));
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(F5, true));
        for (var i = 0; i < 2; i++)
        {
            Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(SyntheticKey.VkPacket, false));
            Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(SyntheticKey.VkPacket, true));
        }

        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(Enter, false));
        Assert.Equal(SyntheticKeyMatch.Completed, tracker.Observe(Enter, true));
    }

    [Fact]
    public void A_real_key_where_a_character_is_expected_ends_the_sequence()
    {
        // The user typed A (0x41) as the Go To text was due: not one of ours.
        var tracker = new SyntheticKeyTracker(new[] { SyntheticKey.CharacterDown('A'), SyntheticKey.CharacterUp('A') });

        Assert.Equal(SyntheticKeyMatch.Unexpected, tracker.Observe(0x41, false));
        Assert.True(tracker.IsOver);
    }

    [Fact]
    public void Keys_inserted_next_are_expected_before_the_rest()
    {
        // An extra Ctrl+Tab sent while the F5 that follows is not sent yet.
        const int Ctrl = 0x11;
        const int Tab = 0x09;
        const int F5 = 0x74;
        var tracker = new SyntheticKeyTracker(new[] { SyntheticKey.Down(Tab), SyntheticKey.Up(Tab), SyntheticKey.Down(F5), SyntheticKey.Up(F5) });
        tracker.Observe(Tab, false);
        tracker.Observe(Tab, true);

        tracker.InsertNext(new[] { SyntheticKey.Down(Ctrl), SyntheticKey.Down(Tab), SyntheticKey.Up(Tab), SyntheticKey.Up(Ctrl) });

        Assert.Equal(8, tracker.Count);
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(Ctrl, false));
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(Tab, false));
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(Tab, true));
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(Ctrl, true));
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(F5, false));
        Assert.Equal(SyntheticKeyMatch.Completed, tracker.Observe(F5, true));
        Assert.Throws<InvalidOperationException>(() => tracker.InsertNext(new[] { SyntheticKey.Down(F5) }));
        Assert.Throws<ArgumentNullException>(() => Tracker().InsertNext(null!));
    }
}
