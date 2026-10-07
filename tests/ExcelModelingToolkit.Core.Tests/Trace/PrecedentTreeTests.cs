using System;
using System.Linq;
using ExcelModelingToolkit.Core.Trace;
using Xunit;
using static ExcelModelingToolkit.Core.Tests.Trace.FakePrecedentProvider;

namespace ExcelModelingToolkit.Core.Tests.Trace;

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
        Assert.Equal("MODEL.XLSX||REVENUE", new PrecedentItem(PrecedentKind.Name, "Revenue", "Model.xlsx").Id);
        Assert.Null(new PrecedentItem(PrecedentKind.Function, "SUM(...)").Id);
        Assert.Null(new PrecedentItem(PrecedentKind.Group, "(x)").Id);
        Assert.Throws<ArgumentOutOfRangeException>(() => new PrecedentItem(PrecedentKind.Range, "r", cellCount: 0));
        Assert.Throws<ArgumentNullException>(() => new PrecedentItem(PrecedentKind.Cell, null!));
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
}
