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
/// used range and the bands outside it, and each is read whole. A mixed rectangle is then halved and each half
/// read the same way (recursive bisection), so k uniform regions in n cells cost about k * log2(n) reads:
/// </para>
/// <list type="bullet">
/// <item><b>Inside the used range</b> a mixed rectangle is cut across its longer side (a tall block into a top and
/// a bottom half).</item>
/// <item><b>Outside the used range</b>, where cells only carry row or column formats, a mixed band is cut across
/// its shorter side (a band under a whole-column selection into its few columns, not its million rows), so the
/// halves soon line up with those formats.</item>
/// </list>
/// <para>
/// A single row or column is always halved along its length. So a uniform whole column costs one read, and n cells
/// whose values all differ cost at most 2n - 1 reads. The weak case is a tall area inside the used range whose
/// columns each have their own format: it is cut into rows first, so it costs about two reads per row (a long one
/// reaches the cap). Every read counts against the cap; when the next read would
/// exceed it, the plan is unavailable. A single cell that reads as mixed (rich text in several font colors), or a
/// cell the reader cannot restore (<see cref="UnrestorableFormatException"/>), also makes the plan unavailable.
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
                run.Read(area, acrossShorter: false);
            }
            else if (!run.Whole(area))
            {
                // Mixed, and partly or wholly outside the used range.
                if (inner is CellRect inside)
                {
                    run.Read(inside, acrossShorter: false);
                    foreach (var band in area.Subtract(inside))
                    {
                        run.Read(band, acrossShorter: true);
                    }
                }
                else
                {
                    run.Split(area, acrossShorter: true);
                }
            }

            if (run.Failure is not null)
            {
                return SnapshotPlan.Unavailable(run.Failure, run.Reads);
            }
        }

        return SnapshotPlan.Captured(run.Blocks, run.Reads);
    }

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

        /// <summary>Reads <paramref name="rect"/> whole; if it is mixed, halves it (see <see cref="Split"/>).</summary>
        public void Read(CellRect rect, bool acrossShorter)
        {
            if (!Whole(rect))
            {
                Split(rect, acrossShorter);
            }
        }

        /// <summary>
        /// Reads <paramref name="rect"/> whole. True if that settled it (recorded as uniform, or the run failed);
        /// false if it is mixed and must be split.
        /// </summary>
        public bool Whole(CellRect rect) => !TryRead(rect, out var value) || Settle(rect, value);

        /// <summary>
        /// Halves a mixed rectangle (of two cells or more) and reads each half. A single row or column is halved
        /// along its length; otherwise the cut goes across the longer side (rows first when square), or across the
        /// shorter side if <paramref name="acrossShorter"/> (a band outside the used range).
        /// </summary>
        public void Split(CellRect rect, bool acrossShorter)
        {
            var cutRows = rect.ColumnCount == 1 ||
                (rect.RowCount > 1 &&
                 (acrossShorter ? rect.RowCount <= rect.ColumnCount : rect.RowCount >= rect.ColumnCount));
            if (cutRows)
            {
                var top = rect.RowCount / 2;
                Read(new CellRect(rect.Row, rect.Column, top, rect.ColumnCount), acrossShorter);
                Read(new CellRect(rect.Row + top, rect.Column, rect.RowCount - top, rect.ColumnCount), acrossShorter);
            }
            else
            {
                var left = rect.ColumnCount / 2;
                Read(new CellRect(rect.Row, rect.Column, rect.RowCount, left), acrossShorter);
                Read(new CellRect(rect.Row, rect.Column + left, rect.RowCount, rect.ColumnCount - left), acrossShorter);
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

        /// <summary>
        /// Reads <paramref name="rect"/> into <paramref name="value"/>. False, with <see cref="Failure"/> set, if
        /// the run has failed, the cap is reached, or the cells cannot be restored.
        /// </summary>
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
            try
            {
                value = _reader.ReadUniform(rect);
            }
            catch (UnrestorableFormatException ex)
            {
                Failure = ex.Message;
                return false;
            }

            return true;
        }
    }
}
