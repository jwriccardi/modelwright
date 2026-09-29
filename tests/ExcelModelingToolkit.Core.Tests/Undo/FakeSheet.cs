using System.Collections.Generic;
using System.Linq;
using ExcelModelingToolkit.Core.Formatting;
using ExcelModelingToolkit.Core.Undo;

namespace ExcelModelingToolkit.Core.Tests.Undo;

/// <summary>
/// A worksheet's format property as painted rectangles over a default value (later paint wins), read like
/// Excel's <c>Range.NumberFormat</c>: a value if the range is uniform, else unknown (mixed). Records each read.
/// </summary>
internal sealed class FakeSheet : IFormatReader
{
    private readonly CycleValue _default;
    private readonly List<(CellRect Rect, CycleValue Value)> _paint = new List<(CellRect, CycleValue)>();
    private readonly HashSet<CellRect> _mixedCells = new HashSet<CellRect>();
    private readonly List<CellRect> _unrestorable = new List<CellRect>();

    public FakeSheet(CycleValue defaultValue) => _default = defaultValue;

    public List<CellRect> Reads { get; } = new List<CellRect>();

    public FakeSheet Paint(CellRect rect, CycleValue value)
    {
        _paint.Add((rect, value));
        return this;
    }

    /// <summary>Makes a single cell read as mixed on its own (rich text in several colors).</summary>
    public FakeSheet MixedCell(int row, int column)
    {
        _mixedCells.Add(new CellRect(row, column, 1, 1));
        return this;
    }

    /// <summary>
    /// Gives <paramref name="rect"/> a format the reader cannot restore (a pattern fill): a read inside it throws
    /// <see cref="UnrestorableFormatException"/>, and a read that also covers other cells is mixed.
    /// </summary>
    public FakeSheet Unrestorable(CellRect rect)
    {
        _unrestorable.Add(rect);
        return this;
    }

    public CycleValue ReadUniform(CellRect range)
    {
        Reads.Add(range);
        foreach (var rect in _unrestorable)
        {
            if (rect.Intersect(range) == range)
            {
                throw new UnrestorableFormatException("pattern or gradient fill");
            }

            if (rect.Intersect(range) is not null)
            {
                return CycleValue.Unknown;
            }
        }

        return Value(range);
    }

    /// <summary>The value of <paramref name="range"/> without recording a read.</summary>
    public CycleValue Value(CellRect range)
    {
        if (_mixedCells.Any(c => range.Intersect(c) is not null))
        {
            return CycleValue.Unknown;
        }

        // Pieces of the range and the value each holds, refined by each paint in order.
        var pieces = new List<(CellRect Rect, CycleValue Value)> { (range, _default) };
        foreach (var (rect, value) in _paint)
        {
            var next = new List<(CellRect, CycleValue)>();
            foreach (var piece in pieces)
            {
                if (piece.Rect.Intersect(rect) is CellRect overlap)
                {
                    next.Add((overlap, value));
                    next.AddRange(piece.Rect.Subtract(overlap).Select(r => (r, piece.Value)));
                }
                else
                {
                    next.Add(piece);
                }
            }

            pieces = next;
        }

        var values = pieces.Select(p => p.Value).Distinct().ToArray();
        return values.Length == 1 ? values[0] : CycleValue.Unknown;
    }
}
