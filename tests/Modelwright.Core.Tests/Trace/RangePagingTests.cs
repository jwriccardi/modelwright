using System;
using System.Collections.Generic;
using System.Linq;
using Modelwright.Core.Trace;
using Modelwright.Core.Undo;
using Xunit;

namespace Modelwright.Core.Tests.Trace;

public class RangePagingTests
{
    // B2:D5 (4 rows x 3 columns, 12 cells) and F1:F3 (3 cells).
    private static readonly CellRect[] Areas = { new CellRect(2, 2, 4, 3), new CellRect(1, 6, 3, 1) };

    // The cells a plan selects, as addresses, in order.
    private static List<string> Cells(IReadOnlyList<CellRect> areas, long start, int count)
    {
        var cells = new List<string>();
        foreach (var block in RangePaging.Plan(areas, start, count))
        {
            for (var i = 0; i < block.Take; i++)
            {
                var (row, column) = block.Offset(i);
                cells.Add(new CellRect(block.Block.Row + row, block.Block.Column + column, 1, 1).Address.Replace("$", string.Empty));
            }
        }

        return cells;
    }

    // Every cell of the areas, row by row, area after area.
    private static List<string> AllCells(IReadOnlyList<CellRect> areas) =>
        areas.SelectMany(a => Enumerable.Range(0, a.RowCount).SelectMany(r => Enumerable.Range(0, a.ColumnCount)
            .Select(c => CellRect.ColumnName(a.Column + c) + (a.Row + r)))).ToList();

    [Fact]
    public void Pages_follow_rows_then_areas()
    {
        Assert.Equal(new[] { "B2", "C2", "D2", "B3", "C3" }, Cells(Areas, 0, 5));
        Assert.Equal(new[] { "D3", "B4", "C4", "D4", "B5" }, Cells(Areas, 5, 5));
        Assert.Equal(new[] { "C5", "D5", "F1", "F2", "F3" }, Cells(Areas, 10, 5));
    }

    [Fact]
    public void Paging_through_every_size_visits_each_cell_once()
    {
        var expected = AllCells(Areas);
        for (var pageSize = 1; pageSize <= 16; pageSize++)
        {
            var seen = new List<string>();
            for (long start = 0; start < 15; start += pageSize)
            {
                seen.AddRange(Cells(Areas, start, pageSize));
            }

            Assert.Equal(expected, seen);
        }
    }

    [Fact]
    public void A_page_within_one_row_reads_only_its_cells()
    {
        var wholeRow = new[] { new CellRect(1, 1, 1, CellRect.MaxColumns) };

        var block = Assert.Single(RangePaging.Plan(wholeRow, 200, 100));

        Assert.Equal(new CellRect(1, 201, 1, 100), block.Block);
        Assert.Equal(0, block.Skip);
        Assert.Equal(100, block.Take);
    }

    [Fact]
    public void A_page_over_several_rows_reads_them_once_at_full_width()
    {
        var block = Assert.Single(RangePaging.Plan(Areas, 2, 5));

        Assert.Equal(new CellRect(2, 2, 3, 3), block.Block);
        Assert.Equal(2, block.Skip);
        Assert.Equal(5, block.Take);
        Assert.Equal((0, 2), block.Offset(0));
        Assert.Equal((2, 0), block.Offset(4));
        Assert.Throws<ArgumentOutOfRangeException>(() => block.Offset(5));
        Assert.Throws<ArgumentOutOfRangeException>(() => block.Offset(-1));
    }

    [Fact]
    public void A_whole_column_pages_deep_down()
    {
        var column = new[] { new CellRect(1, 1, CellRect.MaxRows, 1) };

        var block = Assert.Single(RangePaging.Plan(column, 1048500, 100));

        Assert.Equal(new CellRect(1048501, 1, 76, 1), block.Block);
        Assert.Equal(76, block.Take);
    }

    [Fact]
    public void Nothing_is_planned_past_the_end_or_for_no_cells()
    {
        Assert.Empty(RangePaging.Plan(Areas, 15, 10));
        Assert.Empty(RangePaging.Plan(Areas, 0, 0));
        Assert.Empty(RangePaging.Plan(new CellRect[0], 0, 10));
        Assert.Equal(new[] { "F3" }, Cells(Areas, 14, 10));
    }

    [Fact]
    public void Arguments_are_checked()
    {
        Assert.Throws<ArgumentNullException>(() => RangePaging.Plan(null!, 0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => RangePaging.Plan(Areas, -1, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => RangePaging.Plan(Areas, 0, -1));
        Assert.Equal(1, RangePaging.Plan(Areas, 12, 1).Single().Area);
    }
}
