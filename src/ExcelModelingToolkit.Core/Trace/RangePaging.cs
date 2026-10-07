using System;
using System.Collections.Generic;
using ExcelModelingToolkit.Core.Undo;

namespace ExcelModelingToolkit.Core.Trace;

/// <summary>
/// Plans how to read a page of a range's cells (<see cref="IPrecedentProvider.GetRangeCells"/>) with as few Excel
/// calls as possible: the cells are numbered row by row, area after area (a range like <c>A1:B3,D1:D5</c> has two
/// areas), and each area's share of the page is one rectangular block read with a single <c>Value2</c> call.
/// </summary>
public static class RangePaging
{
    /// <summary>
    /// The blocks to read for cells <paramref name="start"/> to <paramref name="start"/> + <paramref name="count"/>
    /// - 1 (0-based, row-major within each area, areas in order), at most one per area. Fewer cells are planned if
    /// the range ends first; none if <paramref name="start"/> is past the end.
    /// </summary>
    /// <param name="areas">The range's areas.</param>
    /// <param name="start">The index of the first cell wanted.</param>
    /// <param name="count">How many cells are wanted.</param>
    /// <exception cref="ArgumentNullException"><paramref name="areas"/> is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="start"/> or <paramref name="count"/> is negative.</exception>
    public static IReadOnlyList<RangeBlock> Plan(IReadOnlyList<CellRect> areas, long start, int count)
    {
        if (areas is null)
        {
            throw new ArgumentNullException(nameof(areas));
        }

        if (start < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(start), start, "The first cell index cannot be negative.");
        }

        if (count < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(count), count, "The cell count cannot be negative.");
        }

        var blocks = new List<RangeBlock>();
        var offset = start; // into the current area
        var remaining = count;
        for (var index = 0; index < areas.Count && remaining > 0; index++)
        {
            var area = areas[index];
            if (offset >= area.CellCount)
            {
                offset -= area.CellCount;
                continue;
            }

            var take = (int)Math.Min(remaining, area.CellCount - offset);
            var firstRow = (int)(offset / area.ColumnCount);
            var firstColumn = (int)(offset % area.ColumnCount);
            var lastCell = offset + take - 1;
            var lastRow = (int)(lastCell / area.ColumnCount);
            CellRect block;
            int skip;
            if (firstRow == lastRow)
            {
                // Within one row: read just those cells.
                block = new CellRect(area.Row + firstRow, area.Column + firstColumn, 1, take);
                skip = 0;
            }
            else
            {
                // Several rows: read them at full width and skip the cells before the first one wanted.
                block = new CellRect(area.Row + firstRow, area.Column, lastRow - firstRow + 1, area.ColumnCount);
                skip = firstColumn;
            }

            blocks.Add(new RangeBlock(index, block, skip, take));
            remaining -= take;
            offset = 0;
        }

        return blocks;
    }
}

/// <summary>A rectangle of one area to read, and which of its cells (row by row) belong to the page.</summary>
public sealed class RangeBlock
{
    internal RangeBlock(int area, CellRect block, int skip, int take)
    {
        Area = area;
        Block = block;
        Skip = skip;
        Take = take;
    }

    /// <summary>The index of the area the block is in.</summary>
    public int Area { get; }

    /// <summary>The cells to read, in sheet coordinates.</summary>
    public CellRect Block { get; }

    /// <summary>How many of the block's cells, counted row by row, come before the page's first cell.</summary>
    public int Skip { get; }

    /// <summary>How many of the block's cells belong to the page, starting after <see cref="Skip"/>.</summary>
    public int Take { get; }

    /// <summary>
    /// The 0-based row and column, within <see cref="Block"/>, of the page's cell number <paramref name="index"/>
    /// (0 to <see cref="Take"/> - 1): what to index the block's <c>Value2</c> array with (adding its lower bounds).
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is outside the page's cells.</exception>
    public (int Row, int Column) Offset(int index)
    {
        if (index < 0 || index >= Take)
        {
            throw new ArgumentOutOfRangeException(nameof(index), index, "The index must be within the block's page cells.");
        }

        var cell = Skip + index;
        return (cell / Block.ColumnCount, cell % Block.ColumnCount);
    }
}
