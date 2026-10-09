using System;
using System.Collections.Generic;
using System.Globalization;

namespace Modelwright.Core.Trace;

/// <summary>
/// One key event the add-in sends to Excel itself: a press or release of a virtual key (Win32 <c>VK_*</c>), or of a
/// character typed as such (<see cref="Character"/>: <c>SendInput</c> with <c>KEYEVENTF_UNICODE</c>, which the window
/// and the key hook see as <see cref="VkPacket"/>).
/// </summary>
public readonly struct SyntheticKey : IEquatable<SyntheticKey>
{
    /// <summary><c>VK_PACKET</c>: the virtual key a typed character arrives as.</summary>
    public const int VkPacket = 0xE7;

    /// <summary>Creates a key event.</summary>
    public SyntheticKey(int virtualKey, bool keyUp)
        : this(virtualKey, keyUp, '\0')
    {
    }

    private SyntheticKey(int virtualKey, bool keyUp, char character)
    {
        VirtualKey = virtualKey;
        KeyUp = keyUp;
        Character = character;
    }

    /// <summary>The virtual-key code (<see cref="VkPacket"/> for a typed character).</summary>
    public int VirtualKey { get; }

    /// <summary>True for the release, false for the press.</summary>
    public bool KeyUp { get; }

    /// <summary>The character typed (a UTF-16 code unit), or <c>\0</c> for a virtual key.</summary>
    public char Character { get; }

    /// <summary>True for a typed character (sent with <c>KEYEVENTF_UNICODE</c>).</summary>
    public bool IsCharacter => Character != '\0';

    /// <summary>A press of <paramref name="virtualKey"/>.</summary>
    public static SyntheticKey Down(int virtualKey) => new SyntheticKey(virtualKey, false);

    /// <summary>A release of <paramref name="virtualKey"/>.</summary>
    public static SyntheticKey Up(int virtualKey) => new SyntheticKey(virtualKey, true);

    /// <summary>A press of the key that types <paramref name="character"/> (not <c>\0</c>).</summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="character"/> is <c>\0</c>.</exception>
    public static SyntheticKey CharacterDown(char character) => Typed(character, false);

    /// <summary>A release of the key that types <paramref name="character"/> (not <c>\0</c>).</summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="character"/> is <c>\0</c>.</exception>
    public static SyntheticKey CharacterUp(char character) => Typed(character, true);

    /// <inheritdoc />
    public bool Equals(SyntheticKey other) => VirtualKey == other.VirtualKey && KeyUp == other.KeyUp && Character == other.Character;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is SyntheticKey other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => (((VirtualKey * 2) + (KeyUp ? 1 : 0)) * 31) + Character;

    /// <summary><c>0x71 down</c> / <c>0x71 up</c>; a typed character as <c>'A' down</c>.</summary>
    public override string ToString() =>
        (IsCharacter ? "'" + Character + "'" : "0x" + VirtualKey.ToString("X2", CultureInfo.InvariantCulture)) + (KeyUp ? " up" : " down");

    private static SyntheticKey Typed(char character, bool keyUp)
    {
        if (character == '\0')
        {
            throw new ArgumentOutOfRangeException(nameof(character), "A typed character cannot be NUL.");
        }

        return new SyntheticKey(VkPacket, keyUp, character);
    }
}

/// <summary>How a key the hook sees compares with the synthetic keys it is waiting for.</summary>
public enum SyntheticKeyMatch
{
    /// <summary>The next key of the sequence, and more are to come: pass it to Excel.</summary>
    Expected,

    /// <summary>The last key of the sequence: pass it to Excel; the sequence is over.</summary>
    Completed,

    /// <summary>
    /// Not the next key, but a release (the user letting go of the F2 that started the sequence, say): no command acts
    /// on a release, so it is handled as usual and the sequence goes on.
    /// </summary>
    Stray,

    /// <summary>A press that is not the next key (the user typed): the sequence is over; handle the key as usual.</summary>
    Unexpected,

    /// <summary>
    /// A press that is not one of ours and came before any key of the batch sent last (the user's key, queued before
    /// it: an auto-repeat of the F2 that started the sequence, a key rolled over after it, Esc): <c>SendInput</c>
    /// inserts a batch whole, so this key is not inside it. Handle the key as usual; the sequence goes on.
    /// </summary>
    Foreign,

    /// <summary>
    /// A typed character (<see cref="SyntheticKey.VkPacket"/>): never counted off, since the hook is not shown each one
    /// reliably (a release may come without the tag, a press may never come), nor told which character it is. A
    /// keyboard never sends one (only <c>SendInput</c>, an IME or an on-screen keyboard), so pass it to Excel
    /// untouched; the sequence goes on.
    /// </summary>
    Ignored,
}

/// <summary>
/// The keys the add-in has sent to Excel (with <c>SendInput</c>) and the Trace In key hook has not yet seen. While a
/// sequence is in flight the hook passes its keys to Excel untouched, although some are Trace In keys (F2, Right,
/// Shift+Right) and arrive before Excel is editing; it stops at the sequence's last key, or at the first press that
/// is not the next one once a key of the batch sent last has arrived (see <see cref="SyntheticKeyMatch"/>). Typed
/// characters are sent but never tracked: the keys counted off are the sequence's other keys, in order, and a
/// <see cref="SyntheticKey.VkPacket"/> the hook sees is <see cref="SyntheticKeyMatch.Ignored"/> (so a batch of
/// characters and Enter is over at Enter's release). <c>SendInput</c> inserts a batch into the input stream without
/// the user's keys in between, so the batch arrives whole, but the user's keys queued before it (a release, an
/// auto-repeat of a held key, a key rolled over after F2) are still ahead of it: a press before the batch's first key
/// is <see cref="SyntheticKeyMatch.Foreign"/>, and an auto-repeat (or a key the hook knows is not ours: see
/// <see cref="Observe"/>) is never taken for one of the sequence's keys.
/// </summary>
/// <remarks>Not thread-safe: the hook's thread (Excel's main thread).</remarks>
public sealed class SyntheticKeyTracker
{
    // The keys counted off: the sequence's keys that are not typed characters.
    private readonly List<SyntheticKey> _expected;
    private int _next;

    // How many of the expected keys have been sent, and where the batch sent last starts.
    private int _sent;
    private int _batchStart;

    /// <summary>
    /// Starts waiting for the keys of <paramref name="expected"/> that are not typed characters, in order: sent at
    /// once, as one batch, unless <paramref name="sentInBatches"/> (each batch then counts once <see cref="Sending"/>
    /// says it goes).
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="expected"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="expected"/> has no key that is not a typed character.</exception>
    public SyntheticKeyTracker(IReadOnlyList<SyntheticKey> expected, bool sentInBatches = false)
    {
        if (expected is null)
        {
            throw new ArgumentNullException(nameof(expected));
        }

        _expected = Tracked(expected, out var typed);
        if (_expected.Count == 0)
        {
            throw new ArgumentException("A sequence has at least one key that is not a typed character.", nameof(expected));
        }

        if (!sentInBatches)
        {
            _sent = _expected.Count;
            Typed = typed;
        }
    }

    /// <summary>The number of keys in the sequence that are counted off (not typed characters).</summary>
    public int Count => _expected.Count;

    /// <summary>The number of the keys counted off seen so far.</summary>
    public int Seen => _next;

    /// <summary>The number of the keys counted off sent so far.</summary>
    public int Sent => _sent;

    /// <summary>The number of characters typed (presses) in the keys sent so far: sent, never counted off.</summary>
    public int Typed { get; private set; }

    /// <summary>True once the last key has been seen, or a key that ends the sequence early.</summary>
    public bool IsOver { get; private set; }

    /// <summary>
    /// The next keys of the sequence, <paramref name="batch"/>, go now, as one batch (call it before
    /// <c>SendInput</c>): until the first of them that is counted off arrives, a press that is not one of them is
    /// <see cref="SyntheticKeyMatch.Foreign"/>. Its typed characters only add to <see cref="Typed"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="batch"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="batch"/> has no key that is not a typed character, or more such keys than are left.
    /// </exception>
    /// <exception cref="InvalidOperationException">The sequence is over.</exception>
    public void Sending(IReadOnlyList<SyntheticKey> batch)
    {
        if (batch is null)
        {
            throw new ArgumentNullException(nameof(batch));
        }

        var count = Tracked(batch, out var typed).Count;
        if (count < 1 || count > _expected.Count - _sent)
        {
            throw new ArgumentException("A batch has at least one key that is not a typed character, and no more than are left.", nameof(batch));
        }

        if (IsOver)
        {
            throw new InvalidOperationException("The sequence is over.");
        }

        _batchStart = _sent;
        _sent += count;
        Typed += typed;
    }

    /// <summary>
    /// Expects <paramref name="keys"/> next, before the rest of the sequence, sent now as a batch of their own: keys
    /// sent in between, once every key sent so far has been seen (an extra Ctrl+Tab while waiting for a window, say).
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="keys"/> is null.</exception>
    /// <exception cref="InvalidOperationException">The sequence is over.</exception>
    public void InsertNext(IReadOnlyList<SyntheticKey> keys)
    {
        if (keys is null)
        {
            throw new ArgumentNullException(nameof(keys));
        }

        if (IsOver)
        {
            throw new InvalidOperationException("The sequence is over.");
        }

        var tracked = Tracked(keys, out var typed);
        _expected.InsertRange(_next, tracked);
        _batchStart = _next;
        _sent += tracked.Count;
        Typed += typed;
    }

    /// <summary>
    /// Compares a key event the hook sees with the next expected one, and moves on if it matches; a typed character
    /// (<see cref="SyntheticKey.VkPacket"/>) is always <see cref="SyntheticKeyMatch.Ignored"/>. Never a match: a
    /// press that is an auto-repeat (<paramref name="repeat"/>: the key was already down, bit 30 of the hook's
    /// <c>lParam</c>), a key that is <paramref name="foreign"/> (the hook knows it is not one the add-in sent: it lacks
    /// the tag the add-in's keys carry), or a key not sent yet.
    /// </summary>
    public SyntheticKeyMatch Observe(int virtualKey, bool keyUp, bool repeat = false, bool foreign = false)
    {
        if (virtualKey == SyntheticKey.VkPacket)
        {
            return SyntheticKeyMatch.Ignored;
        }

        if (IsOver)
        {
            return SyntheticKeyMatch.Unexpected;
        }

        if (!foreign && (keyUp || !repeat) && _next < _sent)
        {
            var next = _expected[_next];
            if (next.VirtualKey == virtualKey && next.KeyUp == keyUp)
            {
                _next++;
                IsOver = _next == _expected.Count;
                return IsOver ? SyntheticKeyMatch.Completed : SyntheticKeyMatch.Expected;
            }
        }

        if (keyUp)
        {
            return SyntheticKeyMatch.Stray;
        }

        // Nothing of the batch sent last has arrived (or nothing is sent yet): the user's key, from before it.
        if (_next == _batchStart)
        {
            return SyntheticKeyMatch.Foreign;
        }

        IsOver = true;
        return SyntheticKeyMatch.Unexpected;
    }

    // The keys of keys that are counted off (not typed characters); in typed, how many character presses it has.
    private static List<SyntheticKey> Tracked(IReadOnlyList<SyntheticKey> keys, out int typed)
    {
        var tracked = new List<SyntheticKey>(keys.Count);
        typed = 0;
        foreach (var key in keys)
        {
            if (!key.IsCharacter)
            {
                tracked.Add(key);
            }
            else if (!key.KeyUp)
            {
                typed++;
            }
        }

        return tracked;
    }
}
