using System;
using System.Linq;
using Modelwright.Core.Trace;
using Xunit;
using static Modelwright.Core.Tests.Trace.FakePrecedentProvider;

namespace Modelwright.Core.Tests.Trace;

public class PrecedentTreeTests
{
    private static string[] Labels(PrecedentTree tree) => tree.VisibleNodes.Select(node => node.Item.Label).ToArray();

    [Fact]
    public void Opening_shows_the_audited_cells_precedents_in_order_with_duplicates()
    {
        var root = Cell("A2");
        var provider = new FakePrecedentProvider().Has(root, Cell("B2"), Cell("C2"), Cell("B2"));

        var tree = new PrecedentTree(provider, root);

        Assert.Equal(new[] { "Calc!A2", "Calc!B2", "Calc!C2", "Calc!B2" }, Labels(tree));
        Assert.Same(tree.Root, tree.Selected);
        Assert.Equal(0, tree.SelectedIndex);
        Assert.Equal(4, tree.NodeCount);
        Assert.Equal(new[] { "precedents Calc!A2" }, provider.Calls);
    }

    [Fact]
    public void Children_load_only_when_a_node_is_first_expanded()
    {
        var root = Cell("A1");
        var b1 = Cell("B1");
        var provider = new FakePrecedentProvider().Has(root, b1).Has(b1, Cell("C1"));
        var tree = new PrecedentTree(provider, root);
        var node = tree.Root.Children[0];

        Assert.False(node.IsLoaded);
        Assert.True(tree.Expand(node));
        Assert.True(node.IsLoaded);
        Assert.True(tree.Collapse(node));
        Assert.True(tree.Expand(node));

        Assert.Equal(new[] { "precedents Calc!A1", "precedents Calc!B1" }, provider.Calls);
        Assert.Equal(new[] { "Calc!A1", "Calc!B1", "Calc!C1" }, Labels(tree));
    }

    [Fact]
    public void Up_and_down_move_through_visible_rows_and_stop_at_the_ends()
    {
        var root = Cell("A1");
        var tree = new PrecedentTree(new FakePrecedentProvider().Has(root, Cell("B1"), Cell("C1")), root);

        Assert.Equal(TreeMove.None, tree.MoveUp());
        Assert.Equal(TreeMove.Moved, tree.MoveDown());
        Assert.Equal(TreeMove.Moved, tree.MoveDown());
        Assert.Equal("Calc!C1", tree.Selected.Item.Label);
        Assert.Equal(TreeMove.None, tree.MoveDown());
        Assert.Equal(TreeMove.Moved, tree.MoveUp());
        Assert.Equal(1, tree.SelectedIndex);
    }

    [Fact]
    public void Right_expands_then_moves_to_the_first_child()
    {
        var root = Cell("A1");
        var b1 = Cell("B1");
        var tree = new PrecedentTree(new FakePrecedentProvider().Has(root, b1).Has(b1, Cell("C1"), Cell("D1")), root);
        tree.MoveDown();

        Assert.Equal(TreeMove.Expanded, tree.MoveRight());
        Assert.Equal("Calc!B1", tree.Selected.Item.Label);
        Assert.Equal(TreeMove.Moved, tree.MoveRight());
        Assert.Equal("Calc!C1", tree.Selected.Item.Label);
        Assert.Equal(2, tree.Selected.Depth);
    }

    [Fact]
    public void Right_on_a_cell_without_precedents_does_nothing_more()
    {
        var root = Cell("A1");
        var constant = Cell("B1", canExpand: false);
        var empty = Cell("C1");
        var tree = new PrecedentTree(new FakePrecedentProvider().Has(root, constant, empty), root);

        tree.MoveDown();
        Assert.Equal(TreeMove.None, tree.MoveRight());
        tree.MoveDown();
        Assert.Equal(TreeMove.Expanded, tree.MoveRight());
        Assert.False(tree.Selected.CanExpand);
        Assert.Equal(TreeMove.None, tree.MoveRight());
    }

    [Fact]
    public void Left_collapses_an_expanded_node_else_moves_to_the_parent()
    {
        var root = Cell("A1");
        var b1 = Cell("B1");
        var tree = new PrecedentTree(new FakePrecedentProvider().Has(root, b1).Has(b1, Cell("C1")), root);
        tree.MoveDown();
        tree.MoveRight();
        tree.MoveRight();

        Assert.Equal(TreeMove.Moved, tree.MoveLeft());
        Assert.Equal("Calc!B1", tree.Selected.Item.Label);
        Assert.Equal(TreeMove.Collapsed, tree.MoveLeft());
        Assert.False(tree.Selected.IsExpanded);
        Assert.Equal(TreeMove.Moved, tree.MoveLeft());
        Assert.Same(tree.Root, tree.Selected);
        Assert.Equal(TreeMove.Collapsed, tree.MoveLeft());
        Assert.Equal(new[] { "Calc!A1" }, Labels(tree));
        Assert.Equal(TreeMove.None, tree.MoveLeft());
    }

    [Fact]
    public void Left_on_an_expanded_node_without_children_moves_to_the_parent()
    {
        var root = Cell("A1");
        var tree = new PrecedentTree(new FakePrecedentProvider().Has(root, Cell("B1")), root);
        tree.MoveDown();
        tree.MoveRight();

        Assert.True(tree.Selected.IsExpanded);
        Assert.Equal(TreeMove.Moved, tree.MoveLeft());
        Assert.Same(tree.Root, tree.Selected);
    }

    [Fact]
    public void Collapsing_an_ancestor_of_the_selection_selects_it()
    {
        var root = Cell("A1");
        var b1 = Cell("B1");
        var tree = new PrecedentTree(new FakePrecedentProvider().Has(root, b1).Has(b1, Cell("C1")), root);
        tree.MoveDown();
        tree.MoveRight();
        tree.MoveRight();

        Assert.True(tree.Collapse(tree.Root));

        Assert.Same(tree.Root, tree.Selected);
        Assert.False(tree.Collapse(tree.Root));
    }

    [Fact]
    public void Visible_rows_are_depth_first()
    {
        var root = Cell("A1");
        var b1 = Cell("B1");
        var c1 = Cell("C1");
        var provider = new FakePrecedentProvider().Has(root, b1, c1).Has(b1, Cell("B2"), Cell("B3")).Has(c1, Cell("C2"));
        var tree = new PrecedentTree(provider, root);

        tree.Expand(tree.Root.Children[1]);
        tree.Expand(tree.Root.Children[0]);

        Assert.Equal(new[] { "Calc!A1", "Calc!B1", "Calc!B2", "Calc!B3", "Calc!C1", "Calc!C2" }, Labels(tree));
        Assert.Equal(new[] { 0, 1, 2, 2, 1, 2 }, tree.VisibleNodes.Select(node => node.Depth));
    }

    [Fact]
    public void Select_picks_a_visible_node_and_rejects_others()
    {
        var root = Cell("A1");
        var b1 = Cell("B1");
        var tree = new PrecedentTree(new FakePrecedentProvider().Has(root, b1).Has(b1, Cell("C1")), root);
        var b1Node = tree.Root.Children[0];
        tree.Expand(b1Node);
        var hidden = b1Node.Children[0];
        tree.Collapse(b1Node);

        Assert.True(tree.Select(b1Node));
        Assert.False(tree.Select(b1Node));
        Assert.Throws<ArgumentException>(() => tree.Select(hidden));
    }

    [Fact]
    public void Small_range_expands_into_all_its_cells()
    {
        var root = Cell("A1");
        var range = Range("B1:B50", 50);
        var provider = new FakePrecedentProvider().Has(root, range);
        var tree = new PrecedentTree(provider, root);
        var node = tree.Root.Children[0];

        Assert.False(node.IsSummarizedRange);
        tree.Expand(node);

        Assert.Equal(50, node.Children.Count);
        Assert.All(node.Children, child => Assert.Equal(PrecedentKind.Cell, child.Item.Kind));
        Assert.Equal("cells Calc!B1:B50 0 50", provider.Calls.Last());
    }

    [Fact]
    public void Large_range_is_summarized_and_paged_100_cells_at_a_time()
    {
        var root = Cell("A1");
        var range = Range("B1:B250", 250);
        var provider = new FakePrecedentProvider().Has(root, range);
        var tree = new PrecedentTree(provider, root);
        var node = tree.Root.Children[0];
        Assert.True(node.IsSummarizedRange);

        tree.Expand(node);
        Assert.Equal(101, node.Children.Count);
        var more = node.Children[100];
        Assert.Equal(PrecedentKind.MoreCells, more.Item.Kind);
        Assert.Equal("Next 100 of 250 cells (101-200)", more.Item.Label);
        Assert.Null(more.Item.Id);

        tree.Expand(more);
        Assert.Equal(201, node.Children.Count);
        Assert.Equal("Calc!R101", node.Children[100].Item.Label);
        Assert.Equal("Next 50 of 250 cells (201-250)", node.Children[200].Item.Label);

        tree.Expand(node.Children[200]);
        Assert.Equal(250, node.Children.Count);
        Assert.Equal("Calc!R250", node.Children[249].Item.Label);
        Assert.DoesNotContain(node.Children, child => child.Item.Kind == PrecedentKind.MoreCells);
        Assert.Equal(new[] { "cells Calc!B1:B250 0 100", "cells Calc!B1:B250 100 100", "cells Calc!B1:B250 200 50" },
            provider.Calls.Skip(1));
    }

    [Fact]
    public void Right_on_more_cells_loads_the_page_and_selects_its_first_cell()
    {
        var root = Cell("A1");
        var tree = new PrecedentTree(new FakePrecedentProvider().Has(root, Range("B:B", 1048576)), root);
        tree.MoveDown();
        tree.MoveRight();
        var range = tree.Selected;
        tree.Select(range.Children[100]);

        Assert.Equal(TreeMove.Moved, tree.MoveRight());

        Assert.Equal("Calc!R101", tree.Selected.Item.Label);
        Assert.Equal("Next 100 of 1,048,576 cells (201-300)", range.Children[200].Item.Label);
    }

    [Fact]
    public void Range_that_ends_early_gets_no_more_cells_row()
    {
        var root = Cell("A1");
        var range = Range("B1:B500", 500);
        var provider = new FakePrecedentProvider().Has(root, range).RangeHas(range, Cell("B1"), Cell("B2"));
        var tree = new PrecedentTree(provider, root);

        tree.Expand(tree.Root.Children[0]);

        Assert.Equal(new[] { "Calc!B1", "Calc!B2" }, tree.Root.Children[0].Children.Select(child => child.Item.Label));
    }

    [Fact]
    public void Circular_reference_is_marked_and_not_expanded()
    {
        var a1 = Cell("A1");
        var b1 = Cell("B1");
        var provider = new FakePrecedentProvider().Has(a1, b1).Has(b1, Cell("C1"), new PrecedentItem(PrecedentKind.Cell, "back", "model.XLSX", "CALC", "a1"));
        var tree = new PrecedentTree(provider, a1);
        tree.MoveDown();
        tree.MoveRight();
        tree.MoveRight();
        tree.MoveDown();

        var cycle = tree.Selected;
        Assert.Equal("back", cycle.Item.Label);
        Assert.True(cycle.IsCycle);
        Assert.False(cycle.CanExpand);
        Assert.False(tree.Root.Children[0].Children[0].IsCycle);
        Assert.Equal(TreeMove.None, tree.MoveRight());
        Assert.Equal(new[] { "precedents Calc!A1", "precedents Calc!B1" }, provider.Calls);
    }

    [Fact]
    public void Circular_reference_through_a_range_is_found_among_its_cells()
    {
        var a1 = Cell("A1");
        var range = Range("A1:A3", 3);
        var provider = new FakePrecedentProvider().Has(a1, range).RangeHas(range, Cell("A1"), Cell("A2"), Cell("A3"));
        var tree = new PrecedentTree(provider, a1);

        tree.Expand(tree.Root.Children[0]);

        Assert.Equal(new[] { true, false, false }, tree.Root.Children[0].Children.Select(child => child.IsCycle));
    }

    [Fact]
    public void Same_cell_in_another_branch_is_not_a_cycle()
    {
        var a1 = Cell("A1");
        var b1 = Cell("B1");
        var c1 = Cell("C1");
        var provider = new FakePrecedentProvider().Has(a1, b1, c1).Has(c1, b1);
        var tree = new PrecedentTree(provider, a1);

        tree.Expand(tree.Root.Children[1]);

        Assert.False(tree.Root.Children[1].Children[0].IsCycle);
    }

    [Fact]
    public void Node_limit_truncates_with_a_marker_and_stops_reading()
    {
        var root = Cell("A1");
        var many = Enumerable.Range(1, 20).Select(i => Cell("B" + i)).ToArray();
        var provider = new FakePrecedentProvider().Has(root, many);
        var tree = new PrecedentTree(provider, root, maxNodes: 10);

        Assert.Equal(10, tree.NodeCount);
        Assert.Equal(10, tree.Root.Children.Count);
        var marker = tree.Root.Children[9];
        Assert.Equal(PrecedentKind.Truncated, marker.Item.Kind);
        Assert.Equal("Truncated: the node limit (10) was reached", marker.Item.Label);
        Assert.False(marker.CanExpand);

        tree.Expand(tree.Root.Children[0]);

        Assert.Equal(PrecedentKind.Truncated, tree.Root.Children[0].Children.Single().Item.Kind);
        Assert.Single(provider.Calls);
    }

    [Fact]
    public void Node_limit_cuts_a_range_page_short_without_a_more_row()
    {
        var root = Cell("A1");
        var range = Range("B1:B250", 250);
        var provider = new FakePrecedentProvider().Has(root, range);
        var tree = new PrecedentTree(provider, root, maxNodes: 50);

        tree.Expand(tree.Root.Children[0]);

        var children = tree.Root.Children[0].Children;
        Assert.Equal(48, children.Count(child => child.Item.Kind == PrecedentKind.Cell));
        Assert.Equal(PrecedentKind.Truncated, children.Last().Item.Kind);
        Assert.Equal(50, tree.NodeCount);
        Assert.Equal("cells Calc!B1:B250 0 49", provider.Calls.Last());
    }

    [Fact]
    public void Range_expanded_after_the_node_limit_shows_only_a_marker()
    {
        var root = Cell("A1");
        var provider = new FakePrecedentProvider().Has(root, Range("B1:B9", 9), Cell("C1"));
        var tree = new PrecedentTree(provider, root, maxNodes: 3);

        tree.Expand(tree.Root.Children[0]);

        Assert.Equal(PrecedentKind.Truncated, tree.Root.Children[0].Children.Single().Item.Kind);
        Assert.Single(provider.Calls);
    }

    [Fact]
    public void Default_limits_match_the_plan()
    {
        var root = Cell("A1");
        var tree = new PrecedentTree(new FakePrecedentProvider(), root);

        Assert.Equal(5000, tree.MaxNodes);
        Assert.Equal(20, tree.MaxDepth);
        Assert.Equal(50, PrecedentTree.SummaryThreshold);
        Assert.Equal(100, PrecedentTree.PageSize);
    }

    [Fact]
    public void Depth_limit_stops_with_a_marker_without_reading()
    {
        // A chain C0 <- C1 <- C2 <- ... where each cell's precedent is the next.
        var provider = new FakePrecedentProvider();
        for (var i = 0; i < 10; i++)
        {
            provider.Has(Cell("C" + i), Cell("C" + (i + 1)));
        }

        var tree = new PrecedentTree(provider, Cell("C0"), maxDepth: 3);
        while (tree.MoveRight() != TreeMove.None)
        {
        }

        // Right goes as deep as it can: onto the marker below the deepest cell.
        var marker = tree.Selected;
        Assert.Equal(PrecedentKind.Truncated, marker.Item.Kind);
        Assert.Equal("Truncated: the depth limit (3 levels) was reached", marker.Item.Label);
        Assert.False(marker.CanExpand);
        var deepest = marker.Parent!;
        Assert.Equal(3, deepest.Depth);
        Assert.Equal("Calc!C3", deepest.Item.Label);
        Assert.DoesNotContain("precedents Calc!C3", provider.Calls);
    }

    [Fact]
    public void Provider_failure_leaves_the_tree_unchanged_and_can_be_retried()
    {
        var root = Cell("A1");
        var b1 = Cell("B1");
        var provider = new FakePrecedentProvider().Has(root, b1).Has(b1, Cell("C1"));
        var tree = new PrecedentTree(provider, root);
        var node = tree.Root.Children[0];
        provider.ThrowNext = new InvalidOperationException("Excel is busy");

        Assert.Throws<InvalidOperationException>(() => tree.Expand(node));

        Assert.False(node.IsLoaded);
        Assert.False(node.IsExpanded);
        Assert.Equal(2, tree.NodeCount);
        Assert.True(tree.Expand(node));
        Assert.Equal("Calc!C1", node.Children.Single().Item.Label);
    }

    [Fact]
    public void Excel_being_busy_leaves_the_node_unloaded_so_a_later_expand_reads_it()
    {
        var root = Cell("A1");
        var b1 = Cell("B1");
        var provider = new FakePrecedentProvider().Has(root, b1).Has(b1, Cell("C1"));
        var tree = new PrecedentTree(provider, root);
        tree.MoveDown();
        provider.ThrowNext = new PrecedentsUnavailableException();

        Assert.Throws<PrecedentsUnavailableException>(() => tree.MoveRight());

        var node = tree.Root.Children[0];
        Assert.False(node.IsLoaded);
        Assert.True(node.CanExpand);
        Assert.Same(node, tree.Selected);
        Assert.Equal(TreeMove.Expanded, tree.MoveRight());
        Assert.Equal("Calc!C1", node.Children.Single().Item.Label);
    }

    [Fact]
    public void Excel_being_busy_on_a_next_page_keeps_the_more_cells_row()
    {
        var root = Cell("A1");
        var range = Range("A1:A250", 250);
        var provider = new FakePrecedentProvider().Has(root, range);
        var tree = new PrecedentTree(provider, root);
        var rangeNode = tree.Root.Children[0];
        tree.Expand(rangeNode);
        var more = rangeNode.Children.Last();
        provider.ThrowNext = new PrecedentsUnavailableException();

        Assert.Throws<PrecedentsUnavailableException>(() => tree.Expand(more));

        Assert.Same(more, rangeNode.Children.Last());
        Assert.Equal(PrecedentTree.PageSize + 1, rangeNode.Children.Count);
        Assert.True(tree.Expand(more));
        Assert.Equal("Calc!R101", rangeNode.Children[PrecedentTree.PageSize].Item.Label);
    }

    [Theory]
    [InlineData(unchecked((int)0x800AC472), true)]
    [InlineData(unchecked((int)0x80010001), true)]
    [InlineData(unchecked((int)0x8001010A), true)]
    [InlineData(unchecked((int)0x800A03EC), false)]
    [InlineData(0, false)]
    public void Only_the_busy_hresults_count_as_excel_being_busy(int hresult, bool busy)
    {
        Assert.Equal(busy, PrecedentsUnavailableException.IsExcelBusy(hresult));
    }

    [Fact]
    public void Root_that_cannot_expand_shows_alone()
    {
        var provider = new FakePrecedentProvider();
        var tree = new PrecedentTree(provider, Cell("A1", canExpand: false));

        Assert.Single(tree.VisibleNodes);
        Assert.Empty(provider.Calls);
    }

    [Fact]
    public void Constructor_checks_its_arguments()
    {
        var provider = new FakePrecedentProvider();
        Assert.Throws<ArgumentNullException>(() => new PrecedentTree(null!, Cell("A1")));
        Assert.Throws<ArgumentNullException>(() => new PrecedentTree(provider, null!));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PrecedentTree(provider, Cell("A1"), maxNodes: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PrecedentTree(provider, Cell("A1"), maxDepth: 0));
        var tree = new PrecedentTree(provider, Cell("A1"));
        Assert.Throws<ArgumentNullException>(() => tree.Expand(null!));
        Assert.Throws<ArgumentNullException>(() => tree.Collapse(null!));
    }

    [Fact]
    public void Item_ids_are_case_insensitive_locations_and_null_for_structure_nodes()
    {
        Assert.Equal("MODEL.XLSX|CALC|A1", new PrecedentItem(PrecedentKind.Cell, "x", "Model.xlsx", "Calc", "a1").Id);
        Assert.Null(new PrecedentItem(PrecedentKind.Name, "Revenue", "Model.xlsx").Id);
        Assert.Null(new PrecedentItem(PrecedentKind.Table, "[@Amount]", "Model.xlsx", "Calc").Id);
        Assert.Null(new PrecedentItem(PrecedentKind.Function, "SUM(...)").Id);
        Assert.Null(new PrecedentItem(PrecedentKind.Group, "(x)").Id);
        Assert.Throws<ArgumentOutOfRangeException>(() => new PrecedentItem(PrecedentKind.Range, "r", cellCount: 0));
        Assert.Throws<ArgumentNullException>(() => new PrecedentItem(PrecedentKind.Cell, null!));
    }

    [Fact]
    public void Expanding_a_more_cells_row_that_was_already_replaced_does_nothing()
    {
        var root = Cell("A1");
        var provider = new FakePrecedentProvider().Has(root, Range("B1:B250", 250));
        var tree = new PrecedentTree(provider, root);
        var range = tree.Root.Children[0];
        tree.Expand(range);
        var more = range.Children[100];
        Assert.True(tree.Expand(more));
        var calls = provider.Calls.Count;
        var nodes = tree.NodeCount;
        var children = range.Children.ToList();

        Assert.False(tree.Expand(more));

        Assert.Equal(calls, provider.Calls.Count);
        Assert.Equal(nodes, tree.NodeCount);
        Assert.Equal(children, range.Children);
    }

    [Fact]
    public void Item_ids_ignore_dollar_signs_in_addresses()
    {
        Assert.Equal("MODEL.XLSX|CALC|A1", new PrecedentItem(PrecedentKind.Cell, "x", "Model.xlsx", "Calc", "$A$1").Id);
        Assert.Equal("MODEL.XLSX|CALC|A1:B2", new PrecedentItem(PrecedentKind.Range, "x", "Model.xlsx", "Calc", "$A$1:B$2", 4).Id);
        Assert.Equal("MODEL.XLSX|SHEET$1|A1", new PrecedentItem(PrecedentKind.Cell, "x", "Model.xlsx", "Sheet$1", "A1").Id);
    }

    [Fact]
    public void Item_id_of_a_one_cell_range_is_its_cell_as_excel_writes_it()
    {
        var context = new FormulaContext("Model.xlsx", "Calc");
        var references = FormulaParser.Parse("=SUM(A1:A1,$B$2:B2,A:A,3:3,C1:D1)", context).References;

        Assert.Equal(
            new[] { "MODEL.XLSX|CALC|A1", "MODEL.XLSX|CALC|B2", "MODEL.XLSX|CALC|A:A", "MODEL.XLSX|CALC|3:3", "MODEL.XLSX|CALC|C1:D1" },
            references.Select(reference => PrecedentItem.FromReference(reference, context).Id));
        Assert.Equal("MODEL.XLSX|CALC|A1", new PrecedentItem(PrecedentKind.Range, "x", "Model.xlsx", "Calc", "$a$1:A1").Id);
    }

    [Fact]
    public void One_cell_range_of_an_ancestor_cell_is_a_cycle()
    {
        var root = new PrecedentItem(PrecedentKind.Cell, "Calc!$A$1", "Model.xlsx", "Calc", "$A$1");
        var provider = new FakePrecedentProvider().Has(root, Range("A1:A1", 1));

        var tree = new PrecedentTree(provider, root);

        Assert.True(tree.Root.Children[0].IsCycle);
    }

    [Fact]
    public void Names_and_structured_references_without_an_address_never_form_a_false_cycle()
    {
        // [@Amount] in the Sales table and [@Amount] in the Costs table, or a sheet-scoped Rate on two sheets, are
        // different cells written the same way. A real cycle through them still shows, at the cells they expand to.
        var context = new FormulaContext("Model.xlsx", "Calc");
        var amount = FormulaParser.Parse("=[@Amount]", context).References.Single();
        var sales = PrecedentItem.FromReference(amount, context);
        var costs = PrecedentItem.FromReference(amount, context);
        var rate = new PrecedentItem(PrecedentKind.Name, "Rate", "Model.xlsx");
        var sheet2Rate = new PrecedentItem(PrecedentKind.Name, "Rate", "Model.xlsx");
        var root = Cell("A1");
        var b2 = Cell("B2");
        var c1 = Cell("C1", "Sheet2");

        // The fake looks precedents up by label, so costs [@Amount] has the same precedent as sales [@Amount]: B2.
        var provider = new FakePrecedentProvider().Has(root, sales, rate).Has(sales, b2).Has(b2, costs).Has(rate, c1).Has(c1, sheet2Rate);
        var tree = new PrecedentTree(provider, root);
        tree.Expand(tree.Root.Children[0]);
        var b2Node = tree.Root.Children[0].Children[0];
        tree.Expand(b2Node);
        tree.Expand(b2Node.Children[0]);
        tree.Expand(tree.Root.Children[1]);
        tree.Expand(tree.Root.Children[1].Children[0]);

        Assert.Null(sales.Id);
        Assert.Null(PrecedentItem.FromReference(amount, context, "Sales[@Amount]").Id);
        Assert.False(b2Node.Children[0].IsCycle);
        Assert.True(b2Node.Children[0].Children[0].IsCycle);
        Assert.False(tree.Root.Children[1].Children[0].Children[0].IsCycle);
    }

    [Fact]
    public void Absolute_and_relative_addresses_of_the_same_cell_form_a_cycle()
    {
        var root = new PrecedentItem(PrecedentKind.Cell, "Calc!$A$1", "Model.xlsx", "Calc", "$A$1");
        var provider = new FakePrecedentProvider().Has(root, Cell("A1"));

        var tree = new PrecedentTree(provider, root);

        Assert.True(tree.Root.Children[0].IsCycle);
    }

    [Fact]
    public void Items_from_references_fill_in_the_formulas_workbook_and_sheet()
    {
        var context = new FormulaContext("Model.xlsx", "Calc");
        var parsed = FormulaParser.Parse(
            "=$A$1+Calc!A1+Sheet2!$B$2+[Other.xlsx]Data!C3+SUM(D1:E5)+B:B+Rate+Inputs!Tax+Book2.xlsx!Fx+Sales[Amount]+#REF!+Jan:Dec!A1",
            context);

        var items = parsed.References.Select(reference => PrecedentItem.FromReference(reference, context)).ToList();

        Assert.Equal(
            new[]
            {
                "MODEL.XLSX|CALC|A1", "MODEL.XLSX|CALC|A1", "MODEL.XLSX|SHEET2|B2", "OTHER.XLSX|DATA|C3", "MODEL.XLSX|CALC|D1:E5",
                "MODEL.XLSX|CALC|B:B", null, null, null, null, null, "MODEL.XLSX|JAN:DEC|A1",
            },
            items.Select(item => item.Id));
        Assert.Equal(
            new[]
            {
                PrecedentKind.Cell, PrecedentKind.Cell, PrecedentKind.Cell, PrecedentKind.Cell, PrecedentKind.Range, PrecedentKind.Range,
                PrecedentKind.Name, PrecedentKind.Name, PrecedentKind.Name, PrecedentKind.Table, PrecedentKind.Error, PrecedentKind.Cell,
            },
            items.Select(item => item.Kind));
        Assert.Equal("$A$1", items[0].Label);
        Assert.Equal(10, items[4].CellCount);
        Assert.Equal(1048576, items[5].CellCount);
        Assert.False(items[10].CanExpand);
        Assert.Equal("Inputs", items[7].Sheet);
        Assert.Equal("Book2.xlsx", items[8].Workbook);
        Assert.Equal("number1", PrecedentItem.FromReference(parsed.References[0], context, "Calc!A1", "5", argument: "number1").Argument);
        Assert.Equal("Calc!A1", PrecedentItem.FromReference(parsed.References[0], context, "Calc!A1").Label);
    }

    [Fact]
    public void Try_item_from_a_reference_skips_local_names()
    {
        var context = new FormulaContext("Model.xlsx", "Calc");
        var references = FormulaParser.Parse("=LET(x,A1,x+Rate)", context).References;

        var items = references.Select(reference => PrecedentItem.TryFromReference(reference, context)).ToList();

        Assert.Equal(new[] { null, "A1", null, "Rate" }, items.Select(item => item?.Label));
        Assert.Throws<ArgumentNullException>(() => PrecedentItem.TryFromReference(null!, context));
        Assert.Throws<ArgumentNullException>(() => PrecedentItem.TryFromReference(references[1], null!));
    }

    [Fact]
    public void Item_from_a_reference_checks_its_arguments()
    {
        var context = new FormulaContext("Model.xlsx", "Calc");
        var local = FormulaParser.Parse("=LET(x,1,x)", context).References[0];
        var cell = FormulaParser.Parse("=A1", context).References[0];

        Assert.Throws<ArgumentException>(() => PrecedentItem.FromReference(local, context));
        Assert.Throws<ArgumentNullException>(() => PrecedentItem.FromReference(null!, context));
        Assert.Throws<ArgumentNullException>(() => PrecedentItem.FromReference(cell, null!));
    }

    [Fact]
    public void Evaluate_mode_function_nodes_never_count_as_cycles()
    {
        var root = Cell("A1");
        var sum = new PrecedentItem(PrecedentKind.Function, "SUM(...)", argument: "number1");
        var provider = new FakePrecedentProvider().Has(root, sum).Has(sum, sum);
        var tree = new PrecedentTree(provider, root);

        tree.Expand(tree.Root.Children[0]);

        Assert.False(tree.Root.Children[0].Children[0].IsCycle);
        Assert.Equal("number1", tree.Root.Children[0].Item.Argument);
    }

    [Fact]
    public void A_path_brings_the_selection_back_in_a_rebuilt_tree()
    {
        var root = Cell("A1");
        var c1 = Cell("C1");
        var provider = new FakePrecedentProvider().Has(root, Cell("B1"), c1).Has(c1, Cell("D1"), Cell("E1"));
        var tree = new PrecedentTree(provider, root);
        tree.Select(tree.Root.Children[1]);
        tree.MoveRight();
        tree.MoveRight();
        tree.MoveDown(); // E1, under C1

        var path = tree.PathOf(tree.Selected);
        var rebuilt = new PrecedentTree(provider, root);
        var selected = rebuilt.SelectPath(path);

        Assert.Equal(new[] { 1, 1 }, path);
        Assert.Equal("Calc!E1", selected.Item.Label);
        Assert.Same(selected, rebuilt.Selected);
        Assert.Equal(tree.SelectedIndex, rebuilt.SelectedIndex);
        Assert.Empty(rebuilt.PathOf(rebuilt.Root));
        Assert.Same(rebuilt.Root, rebuilt.SelectPath(new int[0]));
    }

    [Fact]
    public void A_path_that_no_longer_exists_stops_at_the_last_node_it_reached()
    {
        // The edited formula now has one reference fewer, or the cell no longer has precedents.
        var root = Cell("A1");
        var b1 = Cell("B1");
        var tree = new PrecedentTree(new FakePrecedentProvider().Has(root, b1).Has(b1, Cell("C1", canExpand: false)), root);

        Assert.Equal("Calc!A1", tree.SelectPath(new[] { 3, 0 }).Item.Label);
        Assert.Equal("Calc!B1", tree.SelectPath(new[] { 0, 5 }).Item.Label);
        Assert.Equal("Calc!C1", tree.SelectPath(new[] { 0, 0, 0 }).Item.Label);
        Assert.Throws<ArgumentNullException>(() => tree.SelectPath(null!));
        Assert.Throws<ArgumentNullException>(() => tree.PathOf(null!));
    }

    [Fact]
    public void A_path_stops_where_the_provider_fails_and_the_failure_propagates()
    {
        var root = Cell("A1");
        var b1 = Cell("B1");
        var provider = new FakePrecedentProvider().Has(root, b1).Has(b1, Cell("C1"));
        var tree = new PrecedentTree(provider, root);
        provider.ThrowNext = new PrecedentsUnavailableException("busy");

        Assert.Throws<PrecedentsUnavailableException>(() => tree.SelectPath(new[] { 0, 0 }));
        Assert.Equal("Calc!B1", tree.Selected.Item.Label);
        Assert.False(tree.Selected.IsLoaded);
    }
}
