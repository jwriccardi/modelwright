using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelModelingToolkit.Core.Trace;

/// <summary>
/// Colors for the Trace In formula header (research/07): each reference is colored as Excel colors references in
/// edit mode. The first distinct reference gets the first <see cref="Palette"/> color, the next distinct one the
/// next, wrapping round; a reference written again (ignoring case and <c>$</c>) keeps its color.
/// </summary>
public static class FormulaColoring
{
    /// <summary>
    /// The reference colors as 0xRRGGBB, in the order Excel assigns them (blue, red, purple, green, orange, maroon,
    /// teal, olive): close to Excel's edit-mode colors, and dark enough to read on a white background.
    /// </summary>
    public static IReadOnlyList<int> Palette { get; } = new[]
    {
        0x0066CC, 0xC00000, 0x7030A0, 0x00803C, 0xC65911, 0x993366, 0x007C80, 0x806000,
    };

    /// <summary>
    /// Splits <paramref name="formula"/>'s text into segments, in order and covering all of it: plain text
    /// (<see cref="FormulaSegment.ColorIndex"/> -1) and references, with the index of their <see cref="Palette"/>
    /// color. LET and LAMBDA local names and <c>#REF!</c> are plain, as in Excel. A formula that could not be parsed
    /// is one plain segment. A reference that overlaps the one before it is left plain.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="formula"/> is null.</exception>
    public static IReadOnlyList<FormulaSegment> Segment(ParsedFormula formula)
    {
        if (formula is null)
        {
            throw new ArgumentNullException(nameof(formula));
        }

        var text = formula.Formula;
        var segments = new List<FormulaSegment>();
        var colors = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var position = 0;
        foreach (var reference in formula.References.OrderBy(r => r.Start))
        {
            if (reference.Kind == FormulaReferenceKind.LocalName || reference.Kind == FormulaReferenceKind.RefError ||
                reference.Start < position || reference.Length == 0 || reference.Start + reference.Length > text.Length)
            {
                continue;
            }

            if (reference.Start > position)
            {
                segments.Add(new FormulaSegment(text.Substring(position, reference.Start - position), position, -1));
            }

            var key = reference.Text.Replace("$", string.Empty);
            if (!colors.TryGetValue(key, out var color))
            {
                color = colors.Count % Palette.Count;
                colors.Add(key, color);
            }

            segments.Add(new FormulaSegment(reference.Text, reference.Start, color));
            position = reference.Start + reference.Length;
        }

        if (position < text.Length)
        {
            segments.Add(new FormulaSegment(text.Substring(position), position, -1));
        }

        return segments;
    }
}

/// <summary>A run of the formula header's text: plain, or a reference in one of <see cref="FormulaColoring.Palette"/>'s colors.</summary>
public sealed class FormulaSegment
{
    /// <summary>Creates a segment.</summary>
    public FormulaSegment(string text, int start, int colorIndex)
    {
        Text = text ?? throw new ArgumentNullException(nameof(text));
        Start = start;
        ColorIndex = colorIndex;
    }

    /// <summary>The segment's text.</summary>
    public string Text { get; }

    /// <summary>The 0-based position of <see cref="Text"/> in the formula.</summary>
    public int Start { get; }

    /// <summary>The index of its color in <see cref="FormulaColoring.Palette"/>, or -1 for plain text.</summary>
    public int ColorIndex { get; }

    /// <summary>True for a reference.</summary>
    public bool IsReference => ColorIndex >= 0;

    /// <summary>The text.</summary>
    public override string ToString() => Text;
}
