using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelModelingToolkit.Core.Formatting;

namespace ExcelModelingToolkit.Core.Undo;

/// <summary>
/// Captures one format property of a selection before a cycle changes it, in as few reads as it can
/// (docs/PLAN.md section 4.4: uniform ranges are run-length compressed, and the capture is capped).
/// </summary>
/// <remarks>
/// <para>
/// Each area of the selection is read whole first. If it is mixed, it is split into its part inside the sheet's
/// used range and the bands outside it.
/// </para>
/// <list type="bullet">
/// <item><b>Inside the used range</b>, where formats vary from line to line: read the whole part; if it is mixed,
/// read it row by row; a mixed row is halved until each piece is uniform (at worst, cell by cell).</item>
/// <item><b>Outside the used range</b>, where cells only carry row or column formats: read each band whole; if it
/// is mixed, split it into lines along its shorter side (a band under a whole-column selection splits into its
/// few columns, not its million rows), then halve mixed lines as above.</item>
/// </list>
/// <para>
/// So a uniform whole column costs one read. Every read counts against the cap; when the next read would exceed
/// it, the plan is unavailable. A single cell that reads as mixed (rich text in several font colors) cannot be
/// restored exactly, so it also makes the plan unavailable.
/// </para>
/// </remarks>
public static class SnapshotPlanner
{
    /// <summary>Captures the value of every cell of <paramref name="areas"/> through <paramref name="reader"/>.</summary>
    /// <param name="areas">The selection's areas (rectangles), in order.</param>
    /// <param name="usedRange">The sheet's used range, or null to treat every cell as inside it.</param>
    /// <param name="cap">The most reads allowed (the <c>undoCellCap</c> setting).</param>
    /// <param name="reader">Reads the property on the selection's sheet.</param>
    /// <exception cref="ArgumentNullException"><paramref name="areas"/> or <paramref name="reader"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="areas"/> is empty.</exception>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="cap"/> is below 1.</exception>
    public static SnapshotPlan Plan(IReadOnlyList<CellRect> areas, CellRect? usedRange, int cap, IFormatReader reader)
    {
        if (areas is null)
        {
            throw new ArgumentNullException(nameof(areas));
        }

        if (areas.Count == 0)
        {
            throw new ArgumentException("A selection has at least one area.", nameof(areas));
        }

        if (cap < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(cap), cap, "The cap must be at least 1.");
        }

        var run = new Run(reader ?? throw new ArgumentNullException(nameof(reader)), cap);
        foreach (var area in areas)
        {
            var inner = usedRange is CellRect used ? area.Intersect(used) : area;
            if (inner == area)
            {
                run.Block(area, byRows: true);
            }
            else if (!run.Whole(area))
            {
                // Mixed, and partly or wholly outside the used range.
                if (inner is CellRect inside)
                {
                    run.Block(inside, byRows: true);
                    foreach (var band in area.Subtract(inside))
                    {
                        run.Block(band, ShorterSide(band));
                    }
                }
                else
                {
                    run.Split(area, ShorterSide(area));
                }
            }

            if (run.Failure is not null)
            {
                return SnapshotPlan.Unavailable(run.Failure, run.Reads);
            }
        }

        return SnapshotPlan.Captured(run.Blocks, run.Reads);
    }

    /// <summary>True to split <paramref name="rect"/> into rows (it is no taller than it is wide), false for columns.</summary>
    private static bool ShorterSide(CellRect rect) => rect.RowCount <= rect.ColumnCount;

    /// <summary>One capture: the blocks so far, the reads made, and the first failure.</summary>
    private sealed class Run
    {
        private readonly IFormatReader _reader;
        private readonly int _cap;

        public Run(IFormatReader reader, int cap)
        {
            _reader = reader;
            _cap = cap;
        }

        public List<SnapshotBlock> Blocks { get; } = new List<SnapshotBlock>();

        public int Reads { get; private set; }

        public string? Failure { get; private set; }

        /// <summary>Reads <paramref name="rect"/>; if mixed, reads it line by line (rows or columns).</summary>
        public void Block(CellRect rect, bool byRows)
        {
            if (!Whole(rect))
            {
                Split(rect, byRows);
            }
        }

        /// <summary>
        /// Reads <paramref name="rect"/> whole. True if that settled it (recorded as uniform, or the run failed);
        /// false if it is mixed and must be split.
        /// </summary>
        public bool Whole(CellRect rect) => !TryRead(rect, out var value) || Settle(rect, value);

        /// <summary>Reads a mixed rectangle line by line (rows or columns); a single row or column is halved.</summary>
        public void Split(CellRect rect, bool byRows)
        {
            if (rect.RowCount == 1 || rect.ColumnCount == 1)
            {
                Halve(rect);
                return;
            }

            var lines = byRows ? rect.RowCount : rect.ColumnCount;
            for (var i = 0; i < lines && Failure is null; i++)
            {
                Line(byRows
                    ? new CellRect(rect.Row + i, rect.Column, 1, rect.ColumnCount)
                    : new CellRect(rect.Row, rect.Column + i, rect.RowCount, 1));
            }
        }

        /// <summary>Reads a single row or column; if mixed, halves it.</summary>
        private void Line(CellRect line)
        {
            if (TryRead(line, out var value) && !Settle(line, value))
            {
                Halve(line);
            }
        }

        /// <summary>Splits a mixed row or column (of two cells or more) in two and reads each half.</summary>
        private void Halve(CellRect line)
        {
            if (line.RowCount > 1)
            {
                var top = line.RowCount / 2;
                Line(new CellRect(line.Row, line.Column, top, 1));
                if (Failure is null)
                {
                    Line(new CellRect(line.Row + top, line.Column, line.RowCount - top, 1));
                }
            }
            else
            {
                var left = line.ColumnCount / 2;
                Line(new CellRect(line.Row, line.Column, 1, left));
                if (Failure is null)
                {
                    Line(new CellRect(line.Row, line.Column + left, 1, line.ColumnCount - left));
                }
            }
        }

        /// <summary>
        /// Records <paramref name="rect"/> if <paramref name="value"/> is uniform and returns true. A mixed single
        /// cell fails the run (and returns true: nothing more to do). Otherwise returns false: split it.
        /// </summary>
        private bool Settle(CellRect rect, CycleValue value)
        {
            if (!value.IsUnknown)
            {
                Blocks.Add(new SnapshotBlock(rect, value, CycleValue.Unknown));
                return true;
            }

            if (rect.IsSingleCell)
            {
                Failure = $"cell {rect.Address} has more than one value (for example, rich text in several colors)";
                return true;
            }

            return false;
        }

        private bool TryRead(CellRect rect, out CycleValue value)
        {
            value = CycleValue.Unknown;
            if (Failure is not null)
            {
                return false;
            }

            if (Reads >= _cap)
            {
                Failure = string.Format(
                    CultureInfo.InvariantCulture,
                    "the selection has too many different formats to record (the undoCellCap setting allows {0} reads)",
                    _cap);
                return false;
            }

            Reads++;
            value = _reader.ReadUniform(rect);
            return true;
        }
    }
}
