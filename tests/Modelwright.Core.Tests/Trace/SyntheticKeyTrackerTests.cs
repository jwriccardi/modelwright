using System;
using Modelwright.Core.Trace;
using Xunit;

namespace Modelwright.Core.Tests.Trace;

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
        Assert.Throws<ArgumentException>(() => new SyntheticKeyTracker(new[] { SyntheticKey.CharacterDown('A'), SyntheticKey.CharacterUp('A') }));
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
    public void Typed_characters_are_never_expected_and_a_VK_PACKET_is_ignored()
    {
        // The Go To step: F5, 'A2' typed, Enter. The hook is not shown each typed character (VK_PACKET) reliably.
        const int F5 = 0x74;
        const int Enter = 0x0D;
        var tracker = new SyntheticKeyTracker(new[]
        {
            SyntheticKey.Down(F5), SyntheticKey.Up(F5), SyntheticKey.CharacterDown('A'), SyntheticKey.CharacterUp('A'),
            SyntheticKey.CharacterDown('2'), SyntheticKey.CharacterUp('2'), SyntheticKey.Down(Enter), SyntheticKey.Up(Enter),
        });

        Assert.Equal(4, tracker.Count);
        Assert.Equal(2, tracker.Typed);
        Assert.Equal(SyntheticKeyMatch.Ignored, tracker.Observe(SyntheticKey.VkPacket, false));
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(F5, false));
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(F5, true));

        // One press seen, its release not, the other press untagged: none of it counts.
        Assert.Equal(SyntheticKeyMatch.Ignored, tracker.Observe(SyntheticKey.VkPacket, false));
        Assert.Equal(SyntheticKeyMatch.Ignored, tracker.Observe(SyntheticKey.VkPacket, false, foreign: true));
        Assert.Equal(SyntheticKeyMatch.Ignored, tracker.Observe(SyntheticKey.VkPacket, true));
        Assert.Equal(2, tracker.Seen);

        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(Enter, false));
        Assert.Equal(SyntheticKeyMatch.Completed, tracker.Observe(Enter, true));
        Assert.Equal(SyntheticKeyMatch.Ignored, tracker.Observe(SyntheticKey.VkPacket, true));
    }

    [Fact]
    public void A_real_key_in_typed_text_is_foreign_until_the_batch_has_begun_then_ends_the_sequence()
    {
        // The user typed A (0x41) while the Go To text arrived: not one of ours.
        const int Enter = 0x0D;
        var tracker = new SyntheticKeyTracker(new[]
        {
            SyntheticKey.CharacterDown('A'), SyntheticKey.CharacterUp('A'), SyntheticKey.Down(Enter), SyntheticKey.Up(Enter),
        });

        Assert.Equal(SyntheticKeyMatch.Ignored, tracker.Observe(SyntheticKey.VkPacket, false));
        Assert.Equal(SyntheticKeyMatch.Foreign, tracker.Observe(0x41, false));
        Assert.False(tracker.IsOver);
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(Enter, false));
        Assert.Equal(SyntheticKeyMatch.Unexpected, tracker.Observe(0x41, false));
        Assert.True(tracker.IsOver);
    }

    [Fact]
    public void A_batch_of_typed_characters_and_Enter_completes_at_Enters_release()
    {
        // The Go To text batch: 'A2' typed, Enter; however many VK_PACKETs the hook is shown.
        const int F5 = 0x74;
        const int Enter = 0x0D;
        var first = new[] { SyntheticKey.Down(F5), SyntheticKey.Up(F5) };
        var text = new[]
        {
            SyntheticKey.CharacterDown('A'), SyntheticKey.CharacterUp('A'), SyntheticKey.CharacterDown('2'), SyntheticKey.CharacterUp('2'),
            SyntheticKey.Down(Enter), SyntheticKey.Up(Enter),
        };
        var keys = new SyntheticKey[first.Length + text.Length];
        first.CopyTo(keys, 0);
        text.CopyTo(keys, first.Length);
        var tracker = new SyntheticKeyTracker(keys, sentInBatches: true);

        tracker.Sending(first);
        tracker.Observe(F5, false);
        tracker.Observe(F5, true);
        tracker.Sending(text);

        Assert.Equal(4, tracker.Sent);
        Assert.Equal(2, tracker.Typed);
        Assert.Equal(SyntheticKeyMatch.Ignored, tracker.Observe(SyntheticKey.VkPacket, false));
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(Enter, false));
        Assert.False(tracker.IsOver);
        Assert.Equal(SyntheticKeyMatch.Completed, tracker.Observe(Enter, true));
        Assert.True(tracker.IsOver);
        Assert.Equal(tracker.Count, tracker.Seen);
    }

    [Fact]
    public void A_press_queued_before_the_batch_is_foreign_and_the_sequence_goes_on()
    {
        // Rollover: the user pressed Down (and Esc) after F2, before the add-in's keys went.
        var tracker = Tracker();

        Assert.Equal(SyntheticKeyMatch.Foreign, tracker.Observe(Down, false));
        Assert.Equal(SyntheticKeyMatch.Foreign, tracker.Observe(0x1B, false));
        Assert.False(tracker.IsOver);
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(F2, false));
        Assert.Equal(1, tracker.Seen);
    }

    [Fact]
    public void An_auto_repeat_is_never_taken_for_a_sent_key()
    {
        // The user still holds the F2 that started the sequence: its auto-repeats come before the add-in's F2.
        var tracker = Tracker();

        Assert.Equal(SyntheticKeyMatch.Foreign, tracker.Observe(F2, false, repeat: true));
        Assert.Equal(SyntheticKeyMatch.Foreign, tracker.Observe(F2, false, repeat: true));
        Assert.Equal(0, tracker.Seen);
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(F2, false));

        // A release always has the previous-state bit: it still matches.
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(F2, true, repeat: true));
        Assert.Equal(2, tracker.Seen);
    }

    [Fact]
    public void A_key_known_not_to_be_ours_never_matches()
    {
        // The user's own F2 (no tag) ahead of the add-in's: foreign before the batch; after it began, it ends it.
        var tracker = Tracker();

        Assert.Equal(SyntheticKeyMatch.Foreign, tracker.Observe(F2, false, foreign: true));
        Assert.Equal(SyntheticKeyMatch.Stray, tracker.Observe(F2, true, foreign: true));
        Assert.Equal(0, tracker.Seen);
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(F2, false));
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(F2, true));
        Assert.Equal(SyntheticKeyMatch.Unexpected, tracker.Observe(Shift, false, foreign: true));
        Assert.True(tracker.IsOver);
    }

    [Fact]
    public void Keys_sent_in_batches_match_only_once_sent_and_each_batch_starts_foreign()
    {
        const int F5 = 0x74;
        const int Enter = 0x0D;
        var first = new[] { SyntheticKey.Down(F2), SyntheticKey.Up(F2), SyntheticKey.Down(F5), SyntheticKey.Up(F5) };
        var text = new[] { SyntheticKey.CharacterDown('A'), SyntheticKey.CharacterUp('A'), SyntheticKey.Down(Enter), SyntheticKey.Up(Enter) };
        var tracker = new SyntheticKeyTracker(
            new[] { first[0], first[1], first[2], first[3], text[0], text[1], text[2], text[3] },
            sentInBatches: true);

        // Nothing sent yet: the user's keys are foreign, even one that looks like ours.
        Assert.Equal(0, tracker.Sent);
        Assert.Equal(0, tracker.Typed);
        Assert.Equal(SyntheticKeyMatch.Foreign, tracker.Observe(F2, false));
        Assert.Equal(0, tracker.Seen);

        tracker.Sending(first);
        Assert.Equal(4, tracker.Sent);
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(F2, false));
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(F2, true));
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(F5, false));
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(F5, true));

        // Waiting for the Go To dialog, the next batch not sent: an Enter is not ours yet.
        Assert.Equal(SyntheticKeyMatch.Stray, tracker.Observe(Enter, true));
        tracker.Sending(text);
        Assert.Equal(6, tracker.Sent);
        Assert.Equal(1, tracker.Typed);
        Assert.Equal(SyntheticKeyMatch.Foreign, tracker.Observe(0x41, false));
        Assert.Equal(SyntheticKeyMatch.Ignored, tracker.Observe(SyntheticKey.VkPacket, false));
        Assert.Equal(SyntheticKeyMatch.Expected, tracker.Observe(Enter, false));
        Assert.Equal(SyntheticKeyMatch.Completed, tracker.Observe(Enter, true));
    }

    [Fact]
    public void A_press_between_batches_ends_the_sequence()
    {
        // The first batch has all arrived and the next is not sent (waiting for the Go To dialog): the user typed.
        const int F5 = 0x74;
        var tracker = new SyntheticKeyTracker(
            new[] { SyntheticKey.Down(F5), SyntheticKey.Up(F5), SyntheticKey.CharacterDown('A'), SyntheticKey.CharacterUp('A'), SyntheticKey.Down(0x0D), SyntheticKey.Up(0x0D) },
            sentInBatches: true);
        tracker.Sending(new[] { SyntheticKey.Down(F5), SyntheticKey.Up(F5) });
        tracker.Observe(F5, false);
        tracker.Observe(F5, true);

        Assert.Equal(SyntheticKeyMatch.Unexpected, tracker.Observe(0x1B, false));
        Assert.True(tracker.IsOver);
    }

    [Fact]
    public void A_batch_has_keys_that_are_left()
    {
        var keys = new[] { SyntheticKey.Down(F2), SyntheticKey.Up(F2) };
        var tracker = new SyntheticKeyTracker(keys, sentInBatches: true);

        Assert.Throws<ArgumentNullException>(() => tracker.Sending(null!));
        Assert.Throws<ArgumentException>(() => tracker.Sending(new SyntheticKey[0]));
        Assert.Throws<ArgumentException>(() => tracker.Sending(new[] { SyntheticKey.CharacterDown('A'), SyntheticKey.CharacterUp('A') }));
        Assert.Throws<ArgumentException>(() => tracker.Sending(new[] { keys[0], keys[1], keys[0] }));
        tracker.Sending(keys);
        Assert.Throws<ArgumentException>(() => tracker.Sending(new[] { keys[0] }));
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
        Assert.Equal(SyntheticKeyMatch.Foreign, tracker.Observe(0x41, false));
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
