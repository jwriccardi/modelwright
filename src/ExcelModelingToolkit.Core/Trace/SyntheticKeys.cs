using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelModelingToolkit.Core.Trace;

/// <summary>One key event the add-in sends to Excel itself: a press or release of a virtual key (Win32 <c>VK_*</c>).</summary>
public readonly struct SyntheticKey : IEquatable<SyntheticKey>
{
    /// <summary>Creates a key event.</summary>
    public SyntheticKey(int virtualKey, bool keyUp)
    {
        VirtualKey = virtualKey;
        KeyUp = keyUp;
    }

    /// <summary>The virtual-key code.</summary>
    public int VirtualKey { get; }

    /// <summary>True for the release, false for the press.</summary>
    public bool KeyUp { get; }

    /// <summary>A press of <paramref name="virtualKey"/>.</summary>
    public static SyntheticKey Down(int virtualKey) => new SyntheticKey(virtualKey, false);

    /// <summary>A release of <paramref name="virtualKey"/>.</summary>
    public static SyntheticKey Up(int virtualKey) => new SyntheticKey(virtualKey, true);

    /// <inheritdoc />
    public bool Equals(SyntheticKey other) => VirtualKey == other.VirtualKey && KeyUp == other.KeyUp;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is SyntheticKey other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => (VirtualKey * 2) + (KeyUp ? 1 : 0);

    /// <summary><c>0x71 down</c> / <c>0x71 up</c>.</summary>
    public override string ToString() =>
        "0x" + VirtualKey.ToString("X2", CultureInfo.InvariantCulture) + (KeyUp ? " up" : " down");
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
/// is not the next one (see <see cref="SyntheticKeyMatch"/>). <c>SendInput</c> inserts a batch into the input stream
/// without the user's keys in between, so the sequence arrives whole, but a release the user made just before can
/// still be ahead of it.
/// </summary>
/// <remarks>Not thread-safe: the hook's thread (Excel's main thread).</remarks>
public sealed class SyntheticKeyTracker
{
    private readonly IReadOnlyList<SyntheticKey> _expected;
    private int _next;

    /// <summary>Starts waiting for <paramref name="expected"/>, in order.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="expected"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="expected"/> is empty.</exception>
    public SyntheticKeyTracker(IReadOnlyList<SyntheticKey> expected)
    {
        _expected = expected ?? throw new ArgumentNullException(nameof(expected));
        if (expected.Count == 0)
        {
            throw new ArgumentException("A sequence has at least one key.", nameof(expected));
        }
    }

    /// <summary>The number of keys in the sequence.</summary>
    public int Count => _expected.Count;

    /// <summary>The number of the sequence's keys seen so far.</summary>
    public int Seen => _next;

    /// <summary>True once the last key has been seen, or a key that ends the sequence early.</summary>
    public bool IsOver { get; private set; }

    /// <summary>Compares a key event the hook sees with the next expected one, and moves on if it matches.</summary>
    public SyntheticKeyMatch Observe(int virtualKey, bool keyUp)
    {
        if (IsOver)
        {
            return SyntheticKeyMatch.Unexpected;
        }

        if (_expected[_next].Equals(new SyntheticKey(virtualKey, keyUp)))
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
