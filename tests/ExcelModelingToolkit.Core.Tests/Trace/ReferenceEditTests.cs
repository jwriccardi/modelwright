using System;
using System.Collections.Generic;
using System.Linq;
using ExcelModelingToolkit.Core.Trace;
using Xunit;
using static ExcelModelingToolkit.Core.Tests.Trace.FakePrecedentProvider;

namespace ExcelModelingToolkit.Core.Tests.Trace;

public class ReferenceEditTests
{
    private const int Shift = 0x10;
    private const int Ctrl = 0x11;
    private static readonly Dictionary<string, int> Vk = new Dictionary<string, int>
    {
        ["F2"] = 0x71, ["Home"] = 0x24, ["End"] = 0x23, ["Left"] = 0x25, ["Right"] = 0x27,
    };

    // The events for "F2", "Ctrl+Home", "Right*5", "Shift+Right*2": each press is a key-down and key-up; a modifier
    // is held around the presses after it.
    private static List<SyntheticKey> Keys(params string[] parts)
    {
        var keys = new List<SyntheticKey>();
        foreach (var part in parts)
        {
            var modifier = part.StartsWith("Ctrl+", StringComparison.Ordinal) ? Ctrl : part.StartsWith("Shift+", StringComparison.Ordinal) ? Shift : 0;
            var rest = modifier == 0 ? part : part.Substring(part.IndexOf('+') + 1);
            var star = rest.IndexOf('*');
            var count = star < 0 ? 1 : int.Parse(rest.Substring(star + 1), System.Globalization.CultureInfo.InvariantCulture);
            var key = Vk[star < 0 ? rest : rest.Substring(0, star)];
            if (modifier != 0)
            {
                keys.Add(SyntheticKey.Down(modifier));
            }

            for (var i = 0; i < count; i++)
            {
                keys.Add(SyntheticKey.Down(key));
                keys.Add(SyntheticKey.Up(key));
            }

            if (modifier != 0)
            {
                keys.Add(SyntheticKey.Up(modifier));
            }
        }

        return keys;
    }

    private static ReferenceSpan Span(string formula, string reference, string address = "A1") =>
        new ReferenceSpan("Model.xlsx", "Calc", address, formula, formula.IndexOf(reference, StringComparison.Ordinal), reference.Length);

    private static PrecedentItem WithSpan(PrecedentItem item, ReferenceSpan span) => item.WithSpan(span);

    [Fact]
    public void A_reference_near_the_start_is_reached_from_the_start()
    {
        // =A2+A14: F2 (caret at the end), Ctrl+Home (before the =), Right past "=", Shift+Right over "A2", F2 (Point).
        Assert.Equal(Keys("F2", "Ctrl+Home", "Right*1", "Shift+Right*2", "F2"), ReferenceEdit.Keys(Span("=A2+A14", "A2")));
    }

    [Fact]
    public void A_reference_near_the_end_is_reached_from_the_end_and_still_selected_forward()
    {
        // =A14+A2 (7 characters), A2 at 5: Ctrl+End, Left twice to its start, Shift+Right over it.
        Assert.Equal(Keys("F2", "Ctrl+End", "Left*2", "Shift+Right*2", "F2"), ReferenceEdit.Keys(Span("=A14+A2", "A2")));
    }

    [Fact]
    public void Halfway_is_reached_from_the_start()
    {
        Assert.Equal(Keys("F2", "Ctrl+Home", "Right*3", "Shift+Right*1", "F2"), ReferenceEdit.Keys(6, 3, 1));
        Assert.Equal(Keys("F2", "Ctrl+End", "Left*2", "Shift+Right*1", "F2"), ReferenceEdit.Keys(6, 4, 1));
    }

    [Fact]
    public void The_whole_sequence_is_capped_at_the_key_event_limit()
    {
        // 10 events for F2, Ctrl+Home, Shift and F2, plus 2 per Right: 1,994 moves and 1 selected character is 4,000.
        var atLimit = ReferenceEdit.Keys(8000, 1994, 1);

        Assert.NotNull(atLimit);
        Assert.Equal(ReferenceEdit.MaxKeyEvents, atLimit!.Count);
        Assert.Null(ReferenceEdit.Keys(8000, 1995, 1));
        Assert.Null(ReferenceEdit.Keys(8000, 4000, 10));
    }

    [Theory]
    [InlineData(7, -1, 2)]
    [InlineData(7, 0, 0)]
    [InlineData(7, 6, 2)]
    [InlineData(7, 8, 1)]
    public void Text_outside_the_formula_is_rejected(int textLength, int start, int length)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => ReferenceEdit.Keys(textLength, start, length));
    }

    [Fact]
    public void A_plain_F2_is_one_press()
    {
        Assert.Equal(Keys("F2"), ReferenceEdit.PlainF2);
    }

    [Fact]
    public void A_span_knows_its_cell_and_text_and_must_lie_in_the_formula()
    {
        var span = Span("=Inputs!B2*(1+Growth)", "Growth", "B2");

        Assert.Equal("Growth", span.Text);
        Assert.Equal(14, span.Start);
        Assert.Equal("Calc!B2[14+6]", span.ToString());
        Assert.Throws<ArgumentOutOfRangeException>(() => new ReferenceSpan("M.xlsx", "Calc", "A1", "=A1", 2, 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => new ReferenceSpan("M.xlsx", "Calc", "A1", "=A1", 1, 0));
        Assert.Throws<ArgumentNullException>(() => new ReferenceSpan("M.xlsx", "Calc", "A1", null!, 0, 1));
    }

    [Fact]
    public void With_span_copies_every_other_property()
    {
        var item = new PrecedentItem(PrecedentKind.Range, "Data!A1:A60", "Model.xlsx", "Data", "A1:A60", 60, "1 (60 cells)",
            canExpand: true, argument: "[number1]", hiddenNote: "hidden rows");
        var span = Span("=SUM(Data!A1:A60)", "Data!A1:A60");

        var copy = item.WithSpan(span);

        Assert.NotSame(item, copy);
        Assert.Null(item.Span);
        Assert.Same(span, copy.Span);
        Assert.Equal(
            new object?[] { item.Kind, item.Label, item.Workbook, item.Sheet, item.Address, item.CellCount, item.ValueText, item.CanExpand, item.Argument, item.HiddenNote, item.Id },
            new object?[] { copy.Kind, copy.Label, copy.Workbook, copy.Sheet, copy.Address, copy.CellCount, copy.ValueText, copy.CanExpand, copy.Argument, copy.HiddenNote, copy.Id });
        Assert.Throws<ArgumentNullException>(() => item.WithSpan(null!));
    }

    [Fact]
    public void The_audited_cell_has_no_reference_to_edit()
    {
        var root = Cell("A1");
        var tree = new PrecedentTree(new FakePrecedentProvider().Has(root, Cell("B1")), root);

        Assert.Null(ReferenceEdit.SpanOf(tree.Root));
    }

    [Fact]
    public void A_reference_in_a_cells_formula_edits_itself_and_a_nested_one_its_own_cells_formula()
    {
        var root = Cell("A1");
        var b2Span = Span("=B2+C2", "B2");
        var b2 = WithSpan(Cell("B2"), b2Span);
        var d2Span = Span("=D2*2", "D2", "B2");
        var d2 = WithSpan(Cell("D2"), d2Span);
        var tree = new PrecedentTree(new FakePrecedentProvider().Has(root, b2).Has(b2, d2), root);
        tree.MoveDown();
        tree.MoveRight();
        tree.MoveRight();

        Assert.Same(d2Span, ReferenceEdit.SpanOf(tree.Selected));
        Assert.Same(b2Span, ReferenceEdit.SpanOf(tree.Selected.Parent!));
    }

    [Fact]
    public void A_cell_of_a_range_and_a_names_target_edit_the_reference_they_came_from()
    {
        var root = Cell("A1");
        var rangeSpan = Span("=SUM(Data!A1:A3)+Rate", "Data!A1:A3");
        var nameSpan = Span("=SUM(Data!A1:A3)+Rate", "Rate");
        var range = WithSpan(Range("A1:A3", 3, "Data"), rangeSpan);
        var name = WithSpan(new PrecedentItem(PrecedentKind.Name, "Rate"), nameSpan);
        var target = Cell("B3", "Inputs");
        var tree = new PrecedentTree(new FakePrecedentProvider().Has(root, range, name).Has(name, target), root);
        tree.MoveDown();
        tree.MoveRight();
        tree.MoveRight(); // the range's first cell

        Assert.Equal(PrecedentKind.Cell, tree.Selected.Item.Kind);
        Assert.Same(rangeSpan, ReferenceEdit.SpanOf(tree.Selected));

        tree.Select(tree.Root.Children[1]);
        tree.MoveRight();
        tree.MoveRight(); // the name's target

        Assert.Same(target, tree.Selected.Item);
        Assert.Same(nameSpan, ReferenceEdit.SpanOf(tree.Selected));
    }

    [Fact]
    public void Markers_errors_and_computed_references_have_no_reference_to_edit()
    {
        var root = Cell("A1");
        var formula = "=SUM(Data!A1:A150)+INDEX(B1:B9,2)+Sheet9!A1";
        var big = WithSpan(Range("A1:A150", 150, "Data"), Span(formula, "Data!A1:A150"));
        var dynamic = new PrecedentItem(PrecedentKind.DynamicReference, "INDEX(B1:B9,2)");
        var error = WithSpan(new PrecedentItem(PrecedentKind.Error, "Sheet9!A1", canExpand: false), Span(formula, "Sheet9!A1"));
        var provider = new FakePrecedentProvider().Has(root, big, dynamic, error).Has(dynamic, Cell("B2"));
        var tree = new PrecedentTree(provider, root);
        tree.Expand(tree.Root.Children[0]);
        tree.Expand(tree.Root.Children[1]);
        var more = tree.Root.Children[0].Children.Last();

        Assert.Equal(PrecedentKind.MoreCells, more.Item.Kind);
        Assert.Null(ReferenceEdit.SpanOf(more));
        Assert.Null(ReferenceEdit.SpanOf(tree.Root.Children[1]));
        Assert.Null(ReferenceEdit.SpanOf(tree.Root.Children[1].Children[0]));
        Assert.Null(ReferenceEdit.SpanOf(tree.Root.Children[2]));
    }

    [Fact]
    public void A_precedent_without_a_span_under_a_cell_does_not_edit_that_cells_reference()
    {
        // Excel's same-sheet precedents of a formula that could not be parsed: not written anywhere we know.
        var root = Cell("A1");
        var b2 = WithSpan(Cell("B2"), Span("=B2", "B2"));
        var fallback = Cell("C5");
        var tree = new PrecedentTree(new FakePrecedentProvider().Has(root, b2).Has(b2, fallback), root);
        tree.Expand(tree.Root.Children[0]);

        Assert.Null(ReferenceEdit.SpanOf(tree.Root.Children[0].Children[0]));
    }

    [Fact]
    public void A_truncation_marker_has_no_reference_to_edit()
    {
        var root = Cell("A1");
        var tree = new PrecedentTree(new FakePrecedentProvider().Has(root, WithSpan(Cell("B1"), Span("=B1+C1", "B1")), Cell("C1")), root, maxNodes: 2);

        Assert.Equal(PrecedentKind.Truncated, tree.Root.Children.Last().Item.Kind);
        Assert.Null(ReferenceEdit.SpanOf(tree.Root.Children.Last()));
        Assert.Throws<ArgumentNullException>(() => ReferenceEdit.SpanOf(null!));
    }
}
