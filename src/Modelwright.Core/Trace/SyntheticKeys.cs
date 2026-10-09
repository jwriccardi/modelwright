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
}

/// <summary>
/// The keys the add-in has sent to Excel (with <c>SendInput</c>) and the Trace In key hook has not yet seen. While a
/// sequence is in flight the hook passes its keys to Excel untouched, although some are Trace In keys (F2, Right,
/// Shift+Right) and arrive before Excel is editing; it stops at the sequence's last key, or at the first press that
/// is not the next one (see <see cref="SyntheticKeyMatch"/>). A typed character arrives as
/// <see cref="SyntheticKey.VkPacket"/> (the hook is not told which character), so it is matched by that key and
/// whether it is a press or release. <c>SendInput</c> inserts a batch into the input stream
/// without the user's keys in between, so the sequence arrives whole, but a release the user made just before can
/// still be ahead of it.
/// </summary>
/// <remarks>Not thread-safe: the hook's thread (Excel's main thread).</remarks>
public sealed class SyntheticKeyTracker
{
    private readonly List<SyntheticKey> _expected;
    private int _next;

    /// <summary>Starts waiting for <paramref name="expected"/>, in order.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="expected"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="expected"/> is empty.</exception>
    public SyntheticKeyTracker(IReadOnlyList<SyntheticKey> expected)
    {
        if (expected is null)
        {
            throw new ArgumentNullException(nameof(expected));
        }

        if (expected.Count == 0)
        {
            throw new ArgumentException("A sequence has at least one key.", nameof(expected));
        }

        _expected = new List<SyntheticKey>(expected);
    }

    /// <summary>The number of keys in the sequence.</summary>
    public int Count => _expected.Count;

    /// <summary>The number of the sequence's keys seen so far.</summary>
    public int Seen => _next;

    /// <summary>True once the last key has been seen, or a key that ends the sequence early.</summary>
    public bool IsOver { get; private set; }

    /// <summary>
    /// Expects <paramref name="keys"/> next, before the rest of the sequence: keys sent in between, once every key
    /// sent so far has been seen (an extra Ctrl+Tab while waiting for a window, say).
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

        _expected.InsertRange(_next, keys);
    }

    /// <summary>Compares a key event the hook sees with the next expected one, and moves on if it matches.</summary>
    public SyntheticKeyMatch Observe(int virtualKey, bool keyUp)
    {
        if (IsOver)
        {
            return SyntheticKeyMatch.Unexpected;
        }

        var next = _expected[_next];
        if (next.VirtualKey == virtualKey && next.KeyUp == keyUp)
        {
            _next++;
            IsOver = _next == _expected.Count;
            return IsOver ? SyntheticKeyMatch.Completed : SyntheticKeyMatch.Expected;
        }

        if (keyUp)
        {
            return SyntheticKeyMatch.Stray;
        }

        IsOver = true;
        return SyntheticKeyMatch.Unexpected;
    }
}
