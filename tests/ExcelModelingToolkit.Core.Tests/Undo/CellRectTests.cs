using System;
using System.Linq;
using ExcelModelingToolkit.Core.Undo;
using Xunit;

namespace ExcelModelingToolkit.Core.Tests.Undo;

public class CellRectTests
{
    [Theory]
    [InlineData(1, "A")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(52, "AZ")]
    [InlineData(53, "BA")]
    [InlineData(702, "ZZ")]
    [InlineData(703, "AAA")]
    [InlineData(16384, "XFD")]
    public void ColumnName_uses_excel_letters(int column, string expected)
    {
        Assert.Equal(expected, CellRect.ColumnName(column));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(16385)]
    public void ColumnName_rejects_columns_outside_the_sheet(int column)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => CellRect.ColumnName(column));
    }

    [Fact]
    public void Address_is_absolute_a1()
    {
        Assert.Equal("$B$3", new CellRect(3, 2, 1, 1).Address);
        Assert.Equal("$B$3:$D$10", new CellRect(3, 2, 8, 3).Address);
        Assert.Equal("$A$1:$A$1048576", new CellRect(1, 1, CellRect.MaxRows, 1).Address);
        Assert.Equal("$A$5:$XFD$6", new CellRect(5, 1, 2, CellRect.MaxColumns).Address);
    }

    [Fact]
    public void Derived_properties()
    {
        var rect = new CellRect(3, 2, 8, 3);

        Assert.Equal(10, rect.LastRow);
        Assert.Equal(4, rect.LastColumn);
        Assert.Equal(24, rect.CellCount);
        Assert.False(rect.IsSingleCell);
        Assert.True(new CellRect(1, 1, 1, 1).IsSingleCell);
        Assert.Equal(17179869184L, new CellRect(1, 1, CellRect.MaxRows, CellRect.MaxColumns).CellCount);
    }

    [Theory]
    [InlineData(0, 1, 1, 1)]
    [InlineData(1, 0, 1, 1)]
    [InlineData(1, 1, 0, 1)]
    [InlineData(1, 1, 1, 0)]
    [InlineData(1048576, 1, 2, 1)]
    [InlineData(1, 16384, 1, 2)]
    [InlineData(1048577, 1, 1, 1)]
    public void Constructor_rejects_rectangles_outside_the_sheet(int row, int column, int rows, int columns)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new CellRect(row, column, rows, columns));
    }

    [Fact]
    public void Intersect_returns_the_overlap_or_null()
    {
        var a = new CellRect(1, 1, 10, 3);

        Assert.Equal(new CellRect(5, 2, 6, 2), a.Intersect(new CellRect(5, 2, 20, 20)));
        Assert.Equal(a, a.Intersect(new CellRect(1, 1, CellRect.MaxRows, 5)));
        Assert.Null(a.Intersect(new CellRect(11, 1, 1, 1)));
        Assert.Null(a.Intersect(new CellRect(1, 4, 1, 1)));
    }

    [Fact]
    public void Subtract_returns_bands_above_below_left_and_right()
    {
        var outer = new CellRect(1, 1, 10, 10); // A1:J10
        var inner = new CellRect(4, 3, 2, 5);   // C4:G5

        var bands = outer.Subtract(inner).Select(b => b.Address).ToArray();

        Assert.Equal(new[] { "$A$1:$J$3", "$A$6:$J$10", "$A$4:$B$5", "$H$4:$J$5" }, bands);
        Assert.Equal(outer.CellCount - inner.CellCount, outer.Subtract(inner).Sum(b => b.CellCount));
    }

    [Fact]
    public void Subtract_of_a_whole_column_by_the_used_range_leaves_the_rows_below()
    {
        var column = new CellRect(1, 2, CellRect.MaxRows, 1); // B:B
        var used = new CellRect(1, 1, 200, 6);                // A1:F200

        Assert.Equal(new[] { "$B$201:$B$1048576" }, column.Subtract(used).Select(b => b.Address));
    }

    [Fact]
    public void Subtract_is_empty_when_covered_and_the_whole_rectangle_when_disjoint()
    {
        var rect = new CellRect(2, 2, 2, 2);

        Assert.Empty(rect.Subtract(new CellRect(1, 1, 5, 5)));
        Assert.Equal(new[] { rect }, rect.Subtract(new CellRect(10, 10, 1, 1)));
    }

    [Fact]
    public void Equality_is_by_value()
    {
        Assert.Equal(new CellRect(1, 2, 3, 4), new CellRect(1, 2, 3, 4));
        Assert.True(new CellRect(1, 2, 3, 4) == new CellRect(1, 2, 3, 4));
        Assert.True(new CellRect(1, 2, 3, 4) != new CellRect(1, 2, 3, 5));
        Assert.Equal(new CellRect(1, 2, 3, 4).GetHashCode(), new CellRect(1, 2, 3, 4).GetHashCode());
        Assert.Equal("$B$1:$E$3", new CellRect(1, 2, 3, 4).ToString());
    }
}
