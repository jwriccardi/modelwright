using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelModelingToolkit.Core.Formatting;
using ExcelModelingToolkit.Core.Undo;
using Xunit;

namespace ExcelModelingToolkit.Core.Tests.Undo;

public class SnapshotPlannerTests
{
    private static readonly CycleValue General = CycleValue.FromNumberFormat("General");
    private static readonly CycleValue OneDecimal = CycleValue.FromNumberFormat("0.0");
    private static readonly CycleValue Percent = CycleValue.FromNumberFormat("0%");
    private static readonly CycleValue Blue = CycleValue.FromColor(OleColor.FromRgb(0, 0, 255));

    /// <summary>Parses <c>B3</c>, <c>B3:D10</c>, <c>A:C</c> (whole columns) or <c>5:6</c> (whole rows).</summary>
    internal static CellRect A1(string address)
    {
        var columns = Regex.Match(address, "^([A-Z]+):([A-Z]+)$");
        if (columns.Success)
        {
            var first = Column(columns.Groups[1].Value);
            return new CellRect(1, first, CellRect.MaxRows, Column(columns.Groups[2].Value) - first + 1);
        }

        var rows = Regex.Match(address, "^([0-9]+):([0-9]+)$");
        if (rows.Success)
        {
            var first = int.Parse(rows.Groups[1].Value, CultureInfo.InvariantCulture);
            return new CellRect(first, 1, int.Parse(rows.Groups[2].Value, CultureInfo.InvariantCulture) - first + 1, CellRect.MaxColumns);
        }

        var cells = Regex.Match(address, "^([A-Z]+)([0-9]+)(?::([A-Z]+)([0-9]+))?$");
        var row = int.Parse(cells.Groups[2].Value, CultureInfo.InvariantCulture);
        var column = Column(cells.Groups[1].Value);
        if (!cells.Groups[3].Success)
        {
            return new CellRect(row, column, 1, 1);
        }

        var lastRow = int.Parse(cells.Groups[4].Value, CultureInfo.InvariantCulture);
        return new CellRect(row, column, lastRow - row + 1, Column(cells.Groups[3].Value) - column + 1);
    }

    private static int Column(string letters) => letters.Aggregate(0, (n, c) => (n * 26) + (c - 'A' + 1));

    private static SnapshotPlan Plan(FakeSheet sheet, string used, int cap, params string[] areas) =>
        SnapshotPlanner.Plan(areas.Select(A1).ToArray(), used is null ? null : A1(used), cap, sheet);

    /// <summary>The blocks tile the areas exactly (for non-overlapping areas), each holding its captured value.</summary>
    private static void AssertCaptures(SnapshotPlan plan, FakeSheet sheet, params string[] areas)
    {
        Assert.True(plan.IsAvailable, plan.UnavailableReason);
        Assert.Equal(plan.Reads, sheet.Reads.Count);
        Assert.Equal(areas.Select(A1).Sum(a => a.CellCount), plan.CellCount);
        Assert.Equal(plan.CellCount, plan.Blocks.Sum(b => b.Range.CellCount));
        foreach (var block in plan.Blocks)
        {
            Assert.Equal(sheet.Value(block.Range), block.Captured);
            Assert.True(block.Applied.IsUnknown);
            Assert.Contains(areas.Select(A1), a => a.Intersect(block.Range) == block.Range);
        }

        for (var i = 0; i < plan.Blocks.Count; i++)
        {
            for (var j = i + 1; j < plan.Blocks.Count; j++)
            {
                Assert.Null(plan.Blocks[i].Range.Intersect(plan.Blocks[j].Range));
            }
        }
    }

    [Fact]
    public void Uniform_whole_column_is_one_read()
    {
        var sheet = new FakeSheet(General);

        var plan = Plan(sheet, "A1:F200", 10000, "B:B");

        AssertCaptures(plan, sheet, "B:B");
        Assert.Equal(1, plan.Reads);
        var block = Assert.Single(plan.Blocks);
        Assert.Equal("$B$1:$B$1048576", block.Address);
        Assert.Equal(General, block.Captured);
        Assert.Equal(1048576, plan.CellCount);
    }

    [Fact]
    public void Uniform_area_inside_the_used_range_is_one_read()
    {
        var sheet = new FakeSheet(Percent);

        var plan = Plan(sheet, "A1:Z100", 10000, "B2:D40");

        AssertCaptures(plan, sheet, "B2:D40");
        Assert.Equal(1, plan.Reads);
    }

    [Fact]
    public void Whole_column_with_a_mixed_used_range_reads_the_rows_below_as_one_block()
    {
        var sheet = new FakeSheet(General)
            .Paint(A1("B3"), OneDecimal)
            .Paint(A1("B7:B9"), Percent);

        var plan = Plan(sheet, "A1:F10", 10000, "B:B");

        AssertCaptures(plan, sheet, "B:B");
        Assert.Contains(plan.Blocks, b => b.Address == "$B$11:$B$1048576" && b.Captured == General);
        Assert.Contains(plan.Blocks, b => b.Address == "$B$3" && b.Captured == OneDecimal);
        Assert.True(plan.Reads < 20, $"{plan.Reads} reads");

        // Nothing below the used range is read in pieces.
        Assert.DoesNotContain(sheet.Reads, r => r.Row > 10 && r.RowCount < CellRect.MaxRows - 10);
    }

    [Fact]
    public void Mixed_block_inside_the_used_range_is_read_row_by_row()
    {
        var sheet = new FakeSheet(General).Paint(A1("A2:C2"), OneDecimal);

        var plan = Plan(sheet, "A1:Z100", 10000, "A1:C3");

        AssertCaptures(plan, sheet, "A1:C3");
        Assert.Equal(new[] { "$A$1:$C$3", "$A$1:$C$1", "$A$2:$C$2", "$A$3:$C$3" }, sheet.Reads.Select(r => r.Address));
        Assert.Equal(new[] { "$A$1:$C$1", "$A$2:$C$2", "$A$3:$C$3" }, plan.Blocks.Select(b => b.Address));
    }

    [Fact]
    public void Mixed_row_is_halved_down_to_cells()
    {
        var sheet = new FakeSheet(General)
            .Paint(A1("B1"), OneDecimal)
            .Paint(A1("C1"), Percent)
            .Paint(A1("D1"), CycleValue.FromNumberFormat("0.00"));

        var plan = Plan(sheet, "A1:Z100", 10000, "A1:D1");

        AssertCaptures(plan, sheet, "A1:D1");
        Assert.Equal(new[] { "$A$1", "$B$1", "$C$1", "$D$1" }, plan.Blocks.Select(b => b.Address));
        Assert.Equal(7, plan.Reads); // A1:D1, A1:B1, A1, B1, C1:D1, C1, D1
    }

    [Fact]
    public void Runs_in_a_long_column_cost_few_reads()
    {
        var sheet = new FakeSheet(General).Paint(A1("A500:A999"), OneDecimal);

        var plan = Plan(sheet, "A1:A1000", 10000, "A1:A1000");

        AssertCaptures(plan, sheet, "A1:A1000");
        Assert.True(plan.Reads < 60, $"{plan.Reads} reads"); // cell by cell would be 1,000
    }

    [Fact]
    public void Whole_columns_with_different_column_formats_split_into_columns_below_the_used_range()
    {
        var sheet = new FakeSheet(General).Paint(A1("B:B"), Percent);

        var plan = Plan(sheet, "A1:C2", 10000, "A:C");

        AssertCaptures(plan, sheet, "A:C");
        Assert.Contains(plan.Blocks, b => b.Address == "$A$3:$A$1048576" && b.Captured == General);
        Assert.Contains(plan.Blocks, b => b.Address == "$B$3:$B$1048576" && b.Captured == Percent);
        Assert.Contains(plan.Blocks, b => b.Address == "$C$3:$C$1048576" && b.Captured == General);
        Assert.Equal(16, plan.Reads); // A:C, A1:C2, 2 rows x (row, A, B:C, B, C), band, its 3 columns
    }

    [Fact]
    public void Whole_rows_with_different_row_formats_split_into_rows_right_of_the_used_range()
    {
        var sheet = new FakeSheet(General).Paint(A1("6:6"), Percent);

        var plan = Plan(sheet, "A1:D10", 10000, "5:6");

        AssertCaptures(plan, sheet, "5:6");
        Assert.Contains(plan.Blocks, b => b.Address == "$E$5:$XFD$5" && b.Captured == General);
        Assert.Contains(plan.Blocks, b => b.Address == "$E$6:$XFD$6" && b.Captured == Percent);
        Assert.True(plan.Reads < 20, $"{plan.Reads} reads");
    }

    [Fact]
    public void Mixed_area_entirely_outside_the_used_range_is_split_along_its_shorter_side()
    {
        var sheet = new FakeSheet(General).Paint(A1("K:K"), Percent);

        var plan = Plan(sheet, "A1:C10", 10000, "J20:L100000");

        AssertCaptures(plan, sheet, "J20:L100000");
        Assert.Equal(new[] { "$J$20:$J$100000", "$K$20:$K$100000", "$L$20:$L$100000" }, plan.Blocks.Select(b => b.Address));
        Assert.Equal(4, plan.Reads);
    }

    [Fact]
    public void Without_a_used_range_every_cell_counts_as_inside()
    {
        var sheet = new FakeSheet(General).Paint(A1("A2"), Percent);

        var plan = Plan(sheet, null!, 10000, "A1:B2");

        AssertCaptures(plan, sheet, "A1:B2");
        Assert.Equal(new[] { "$A$1:$B$1", "$A$2", "$B$2" }, plan.Blocks.Select(b => b.Address));
    }

    [Fact]
    public void Multi_area_selection_captures_every_area_in_order()
    {
        var sheet = new FakeSheet(General).Paint(A1("D5"), OneDecimal);

        var plan = Plan(sheet, "A1:Z100", 10000, "A1:A3", "D4:D5", "F:F");

        AssertCaptures(plan, sheet, "A1:A3", "D4:D5", "F:F");
        Assert.Equal(new[] { "$A$1:$A$3", "$D$4", "$D$5", "$F$1:$F$1048576" }, plan.Blocks.Select(b => b.Address));
    }

    [Fact]
    public void Colors_are_captured_like_number_formats()
    {
        var sheet = new FakeSheet(CycleValue.FromColor(OleColor.NoFill)).Paint(A1("B2"), Blue);

        var plan = Plan(sheet, "A1:C3", 10000, "A1:C3");

        AssertCaptures(plan, sheet, "A1:C3");
        Assert.Contains(plan.Blocks, b => b.Address == "$B$2" && b.Captured == Blue);
    }

    [Fact]
    public void Exceeding_the_cap_makes_the_plan_unavailable()
    {
        var sheet = new FakeSheet(General);
        for (var row = 1; row <= 20; row += 2)
        {
            sheet.Paint(new CellRect(row, 1, 1, 1), Percent);
        }

        var plan = Plan(sheet, "A1:A20", 10, "A1:A20");

        Assert.False(plan.IsAvailable);
        Assert.Empty(plan.Blocks);
        Assert.Equal(10, plan.Reads);
        Assert.Equal(10, sheet.Reads.Count);
        Assert.Contains("undoCellCap", plan.UnavailableReason);
        Assert.Contains("10 reads", plan.UnavailableReason);
    }

    [Fact]
    public void A_capture_that_needs_exactly_the_cap_is_available()
    {
        var sheet = new FakeSheet(General).Paint(A1("A2:C2"), OneDecimal);

        var plan = Plan(sheet, "A1:Z100", 4, "A1:C3");

        AssertCaptures(plan, sheet, "A1:C3");
        Assert.Equal(4, plan.Reads);
    }

    [Fact]
    public void The_cap_counts_reads_across_areas()
    {
        var sheet = new FakeSheet(General);

        var plan = Plan(sheet, "A1:Z100", 2, "A1", "B2", "C3");

        Assert.False(plan.IsAvailable);
        Assert.Equal(2, plan.Reads);
    }

    [Fact]
    public void A_single_cell_that_is_mixed_makes_the_plan_unavailable()
    {
        var sheet = new FakeSheet(Blue).MixedCell(2, 2);

        var plan = Plan(sheet, "A1:Z100", 10000, "A1:C3");

        Assert.False(plan.IsAvailable);
        Assert.Contains("$B$2", plan.UnavailableReason);
    }

    [Fact]
    public void Arguments_are_checked()
    {
        var sheet = new FakeSheet(General);
        var areas = new[] { A1("A1") };

        Assert.Throws<ArgumentNullException>(() => SnapshotPlanner.Plan(null!, null, 1, sheet));
        Assert.Throws<ArgumentNullException>(() => SnapshotPlanner.Plan(areas, null, 1, null!));
        Assert.Throws<ArgumentException>(() => SnapshotPlanner.Plan(new List<CellRect>(), null, 1, sheet));
        Assert.Throws<ArgumentOutOfRangeException>(() => SnapshotPlanner.Plan(areas, null, 0, sheet));
    }

    [Fact]
    public void Unavailable_requires_a_reason()
    {
        Assert.Throws<ArgumentNullException>(() => SnapshotPlan.Unavailable(null!, 0));
        Assert.Equal("because", SnapshotPlan.Unavailable("because", 3).UnavailableReason);
    }
}
