using System;
using System.Linq;
using Modelwright.Core.Formatting;
using Modelwright.Core.Undo;
using Xunit;

namespace Modelwright.Core.Tests.Undo;

public class FormatSnapshotTests
{
    private static readonly CycleValue General = CycleValue.FromNumberFormat("General");
    private static readonly CycleValue Percent = CycleValue.FromNumberFormat("0%");
    private static readonly CycleValue Applied = CycleValue.FromNumberFormat("#,##0.0_);(#,##0.0)");
    private static readonly CycleValue Blue = CycleValue.FromColor(OleColor.FromRgb(0, 0, 255));
    private static readonly CycleValue NoFill = CycleValue.FromColor(OleColor.NoFill);

    private static SnapshotBlock Block(string address, CycleValue captured, CycleValue applied) =>
        new SnapshotBlock(SnapshotPlannerTests.A1(address), captured, applied);

    [Fact]
    public void Create_keeps_the_blocks_and_counts_cells()
    {
        var snapshot = FormatSnapshot.Create("General Number", CycleKind.NumberFormat, "Model.xlsx", "Calc", new[]
        {
            Block("A1:B2", General, Applied),
            Block("C:C", Percent, Applied),
        });

        Assert.True(snapshot.IsAvailable);
        Assert.Null(snapshot.UnavailableReason);
        Assert.Equal("General Number", snapshot.Label);
        Assert.Equal(CycleKind.NumberFormat, snapshot.Kind);
        Assert.Equal("Model.xlsx", snapshot.Workbook);
        Assert.Equal("Calc", snapshot.Sheet);
        Assert.Equal(2, snapshot.Blocks.Count);
        Assert.Equal(4 + 1048576, snapshot.CellCount);
        Assert.Equal("General Number on Model.xlsx|Calc: 2 blocks, 1048580 cells", snapshot.ToString());
    }

    [Fact]
    public void Unavailable_has_a_reason_and_no_blocks()
    {
        var snapshot = FormatSnapshot.Unavailable("Fill", CycleKind.FillColor, "Model.xlsx", "Calc", "too many formats");

        Assert.False(snapshot.IsAvailable);
        Assert.Equal("too many formats", snapshot.UnavailableReason);
        Assert.Empty(snapshot.Blocks);
        Assert.Equal(0, snapshot.CellCount);
        Assert.Empty(snapshot.WriteGroups(UndoKey.Undo));
        Assert.Equal("Fill on Model.xlsx|Calc: unavailable (too many formats)", snapshot.ToString());
        Assert.Throws<ArgumentNullException>(() => FormatSnapshot.Unavailable("Fill", CycleKind.FillColor, "B", "S", null!));
    }

    [Fact]
    public void Create_rejects_missing_or_invalid_blocks()
    {
        Assert.Throws<ArgumentNullException>(() => FormatSnapshot.Create("N", CycleKind.NumberFormat, "B", "S", null!));
        Assert.Throws<ArgumentNullException>(() => FormatSnapshot.Create(null!, CycleKind.NumberFormat, "B", "S", new[] { Block("A1", General, Applied) }));
        Assert.Throws<ArgumentNullException>(() => FormatSnapshot.Create("N", CycleKind.NumberFormat, null!, "S", new[] { Block("A1", General, Applied) }));
        Assert.Throws<ArgumentNullException>(() => FormatSnapshot.Create("N", CycleKind.NumberFormat, "B", null!, new[] { Block("A1", General, Applied) }));
        Assert.Throws<ArgumentException>(() => FormatSnapshot.Create("N", CycleKind.NumberFormat, "B", "S", Array.Empty<SnapshotBlock>()));
        Assert.Throws<ArgumentException>(() => FormatSnapshot.Create("N", CycleKind.NumberFormat, "B", "S", new SnapshotBlock[] { null! }));

        // Applied not yet known (the write has not happened).
        Assert.Throws<ArgumentException>(() => FormatSnapshot.Create("N", CycleKind.NumberFormat, "B", "S", new[] { Block("A1", General, CycleValue.Unknown) }));

        // Values of the wrong kind.
        Assert.Throws<ArgumentException>(() => FormatSnapshot.Create("N", CycleKind.FontColor, "B", "S", new[] { Block("A1", General, Applied) }));
        Assert.Throws<ArgumentException>(() => FormatSnapshot.Create("N", CycleKind.NumberFormat, "B", "S", new[] { Block("A1", Blue, Blue) }));
        Assert.Throws<ArgumentException>(() => FormatSnapshot.Create("N", CycleKind.FontColor, "B", "S", new[] { Block("A1", NoFill, Blue) }));
    }

    [Fact]
    public void Fill_snapshots_accept_no_fill()
    {
        var snapshot = FormatSnapshot.Create("Fill", CycleKind.FillColor, "B", "S", new[] { Block("A1", NoFill, Blue) });

        Assert.Equal(NoFill, snapshot.Blocks[0].Captured);
    }

    [Fact]
    public void Font_snapshots_accept_automatic_and_other_kinds_do_not()
    {
        var automatic = CycleValue.Automatic;
        var snapshot = FormatSnapshot.Create("Font", CycleKind.FontColor, "B", "S", new[] { Block("A1", automatic, Blue) });

        Assert.Equal(automatic, snapshot.Blocks[0].Captured);
        Assert.Equal(automatic, Assert.Single(snapshot.WriteGroups(UndoKey.Undo)).Value);
        Assert.Throws<ArgumentException>(() => FormatSnapshot.Create("Fill", CycleKind.FillColor, "B", "S", new[] { Block("A1", automatic, Blue) }));
        Assert.Throws<ArgumentException>(() => FormatSnapshot.Create("N", CycleKind.NumberFormat, "B", "S", new[] { Block("A1", automatic, Applied) }));
    }

    [Fact]
    public void Automatic_and_explicit_black_stay_separate_blocks_when_undone()
    {
        var black = CycleValue.FromColor(OleColor.FromRgb(0, 0, 0));
        var snapshot = FormatSnapshot.Create("Font", CycleKind.FontColor, "B", "S", new[]
        {
            Block("A1", CycleValue.Automatic, Blue),
            Block("A2", black, Blue),
        });

        var groups = snapshot.WriteGroups(UndoKey.Undo);

        Assert.Equal(2, groups.Count);
        Assert.Equal(CycleValue.Automatic, groups[0].Value);
        Assert.Equal(black, groups[1].Value);
    }

    [Theory]
    [InlineData(@"C:\Deals\Model.xlsx", "Model.xlsx")]
    [InlineData("https://contoso.sharepoint.com/sites/x/Shared Documents/Model.xlsx", "Model.xlsx")]
    [InlineData("Book1", "Book1")]
    public void WorkbookName_is_the_file_name_of_the_full_name(string fullName, string name)
    {
        var snapshot = FormatSnapshot.Unavailable("N", CycleKind.NumberFormat, fullName, "S", "reason");

        Assert.Equal(fullName, snapshot.Workbook);
        Assert.Equal(name, snapshot.WorkbookName);
    }

    [Fact]
    public void Block_captured_value_must_be_known()
    {
        Assert.Throws<ArgumentException>(() => Block("A1", CycleValue.Unknown, Applied));
    }

    [Fact]
    public void Blocks_share_one_string_per_number_format_code()
    {
        // Strings read over COM are new instances each time.
        var first = new SnapshotBlock(new CellRect(1, 1, 1, 1), CycleValue.FromNumberFormat(new string("0.0%x".ToCharArray(), 0, 4)), CycleValue.Unknown);
        var second = new SnapshotBlock(new CellRect(2, 1, 1, 1), CycleValue.FromNumberFormat(new string("0.0%y".ToCharArray(), 0, 4)), CycleValue.Unknown)
            .WithApplied(CycleValue.FromNumberFormat(new string("0.0%z".ToCharArray(), 0, 4)));

        Assert.Same(first.Captured.NumberFormat, second.Captured.NumberFormat);
        Assert.Same(first.Captured.NumberFormat, second.Applied.NumberFormat);
    }

    [Fact]
    public void Block_expected_and_target_depend_on_the_direction()
    {
        var block = Block("A1", General, CycleValue.Unknown).WithApplied(Applied);

        Assert.Equal(Applied, block.Expected(UndoKey.Undo));
        Assert.Equal(General, block.Target(UndoKey.Undo));
        Assert.Equal(General, block.Expected(UndoKey.Redo));
        Assert.Equal(Applied, block.Target(UndoKey.Redo));
        Assert.Equal("$A$1", block.Address);
        Assert.Equal("$A$1 General -> #,##0.0_);(#,##0.0)", block.ToString());
    }

    [Fact]
    public void Undo_groups_blocks_by_captured_value()
    {
        var snapshot = FormatSnapshot.Create("N", CycleKind.NumberFormat, "B", "S", new[]
        {
            Block("A1", General, Applied),
            Block("A2", Percent, Applied),
            Block("A3:B4", General, Applied),
        });

        var groups = snapshot.WriteGroups(UndoKey.Undo);

        Assert.Equal(2, groups.Count);
        Assert.Equal("$A$1,$A$3:$B$4", groups[0].Address);
        Assert.Equal(General, groups[0].Value);
        Assert.Equal(2, groups[0].Blocks.Count);
        Assert.Equal("$A$2", groups[1].Address);
        Assert.Equal(Percent, groups[1].Value);
    }

    [Fact]
    public void Redo_writes_the_applied_value_to_every_block_at_once()
    {
        var snapshot = FormatSnapshot.Create("N", CycleKind.NumberFormat, "B", "S", new[]
        {
            Block("A1", General, Applied),
            Block("A2", Percent, Applied),
        });

        var group = Assert.Single(snapshot.WriteGroups(UndoKey.Redo));

        Assert.Equal("$A$1,$A$2", group.Address);
        Assert.Equal(Applied, group.Value);
    }

    [Fact]
    public void Groups_stay_within_the_range_address_limit_and_cover_every_block_once()
    {
        var blocks = Enumerable.Range(1, 200)
            .Select(row => new SnapshotBlock(new CellRect(row * 2, 3, 1, 1), row % 3 == 0 ? Percent : General, Applied))
            .ToArray();
        var snapshot = FormatSnapshot.Create("N", CycleKind.NumberFormat, "B", "S", blocks);

        foreach (var key in new[] { UndoKey.Undo, UndoKey.Redo })
        {
            var groups = snapshot.WriteGroups(key);

            Assert.All(groups, g => Assert.True(g.Address.Length <= FormatSnapshot.MaxAddressLength, g.Address));
            Assert.All(groups, g => Assert.Equal(string.Join(",", g.Blocks.Select(b => b.Address)), g.Address));
            Assert.All(groups, g => Assert.All(g.Blocks, b => Assert.Equal(g.Value, b.Target(key))));
            var covered = groups.SelectMany(g => g.Blocks).ToList();
            Assert.Equal(blocks.Length, covered.Count);
            Assert.Equal(blocks.Length, covered.Distinct().Count());
            Assert.True(groups.Count > 2);
        }
    }
}
