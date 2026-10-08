using System;
using System.Collections.Generic;

namespace ExcelModelingToolkit.Core.Trace;

/// <summary>
/// F2 in the Trace In window (docs/PLAN.md section 4.5): which reference the selected row edits, and the keys that
/// put Excel in Point mode on it. Excel has no API for the edit caret, so the add-in goes to the cell whose formula
/// holds the reference and sends Excel these keys, as if typed: F2 (Edit mode, the caret at the end), the caret to the
/// reference, Shift+Right over it, and F2 again (Excel's Point mode: the arrow keys and sheet tabs now replace the
/// selected reference with the cell they point at).
/// </summary>
public static class ReferenceEdit
{
    /// <summary>
    /// The most key events (presses and releases) one F2 sends; a reference further into a longer formula is not
    /// selected (F2 then edits the cell as usual).
    /// </summary>
    public const int MaxKeyEvents = 4000;

    private const int VkShift = 0x10;
    private const int VkControl = 0x11;
    private const int VkEnd = 0x23;
    private const int VkHome = 0x24;
    private const int VkLeft = 0x25;
    private const int VkRight = 0x27;
    private const int VkF2 = 0x71;

    /// <summary>A plain F2 (press and release): Excel edits the active cell, as without the window.</summary>
    public static IReadOnlyList<SyntheticKey> PlainF2 { get; } = new[] { SyntheticKey.Down(VkF2), SyntheticKey.Up(VkF2) };

    /// <summary>
    /// The reference F2 edits for the selected row <paramref name="node"/>: the nearest of the node and its ancestors
    /// that is written in a cell's formula (<see cref="PrecedentItem.Span"/>), going up only from a row that stands for
    /// its parent's reference (a cell of a range, a name's or table reference's target or the references in a name's
    /// formula). Null for the audited cell, an error, a "more cells" or "truncated" row, a computed reference
    /// (INDEX, OFFSET, ...) and what it points to, a function or group, and Excel's same-sheet precedents (a formula
    /// that could not be parsed): F2 then edits the active cell as usual.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="node"/> is null.</exception>
    public static ReferenceSpan? SpanOf(PrecedentNode node)
    {
        if (node is null)
        {
            throw new ArgumentNullException(nameof(node));
        }

        switch (node.Item.Kind)
        {
            case PrecedentKind.Error:
            case PrecedentKind.MoreCells:
            case PrecedentKind.Truncated:
                return null;
        }

        for (var current = node; current is not null; current = current.Parent)
        {
            if (current.Item.Kind == PrecedentKind.Function || current.Item.Kind == PrecedentKind.Group)
            {
                return null;
            }

            if (current.Item.Span is ReferenceSpan span)
            {
                return span;
            }

            if (current.Parent is null || !StandsForItsChildren(current.Parent.Item.Kind))
            {
                return null;
            }
        }

        return null;
    }

    /// <summary>
    /// The keys that select <paramref name="span"/>'s reference in Excel's editor and switch to Point mode (see
    /// <see cref="Keys(int, int, int)"/>), or null if that takes more than <see cref="MaxKeyEvents"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="span"/> is null.</exception>
    public static IReadOnlyList<SyntheticKey>? Keys(ReferenceSpan span)
    {
        if (span is null)
        {
            throw new ArgumentNullException(nameof(span));
        }

        return Keys(span.Formula.Length, span.Start, span.Length);
    }

    /// <summary>
    /// The keys that select the text from <paramref name="start"/> (0-based, the <c>=</c> being 0) for
    /// <paramref name="length"/> characters of a cell's formula of <paramref name="textLength"/> characters, as Excel's
    /// editor shows it, and switch to Point mode; null if that takes more than <see cref="MaxKeyEvents"/> key events.
    /// In order: F2 (Edit mode, the caret at the end); the caret to <paramref name="start"/> from the nearer end, so
    /// Ctrl+Home then Right <paramref name="start"/> times, or Ctrl+End then Left
    /// (<paramref name="textLength"/> - <paramref name="start"/>) times; Shift+Right <paramref name="length"/> times
    /// (the selection always runs forward, so the caret ends after the reference); F2 (Point mode).
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The text is not inside the formula, or is empty.</exception>
    public static IReadOnlyList<SyntheticKey>? Keys(int textLength, int start, int length)
    {
        if (start < 0 || length < 1 || start > textLength - length)
        {
            throw new ArgumentOutOfRangeException(nameof(start), start, "The text must lie inside the formula.");
        }

        var fromStart = start <= textLength - start;
        var moves = fromStart ? start : textLength - start;

        // F2, Ctrl+Home/End, Shift and F2: 10 events; each Right or Left: 2.
        if (10 + (2L * moves) + (2L * length) > MaxKeyEvents)
        {
            return null;
        }

        var keys = new List<SyntheticKey>(10 + (2 * moves) + (2 * length));
        Press(keys, VkF2);
        keys.Add(SyntheticKey.Down(VkControl));
        Press(keys, fromStart ? VkHome : VkEnd);
        keys.Add(SyntheticKey.Up(VkControl));
        for (var i = 0; i < moves; i++)
        {
            Press(keys, fromStart ? VkRight : VkLeft);
        }

        keys.Add(SyntheticKey.Down(VkShift));
        for (var i = 0; i < length; i++)
        {
            Press(keys, VkRight);
        }

        keys.Add(SyntheticKey.Up(VkShift));
        Press(keys, VkF2);
        return keys;
    }

    // A row under a range (a page of its cells), a name or a table reference is reached through that reference.
    private static bool StandsForItsChildren(PrecedentKind kind) =>
        kind == PrecedentKind.Range || kind == PrecedentKind.Name || kind == PrecedentKind.Table;

    private static void Press(List<SyntheticKey> keys, int virtualKey)
    {
        keys.Add(SyntheticKey.Down(virtualKey));
        keys.Add(SyntheticKey.Up(virtualKey));
    }
}
