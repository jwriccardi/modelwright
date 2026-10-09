using System;
using System.Collections.Generic;
using System.Linq;
using Modelwright.Core.Trace;
using Xunit;
using static Modelwright.Core.Tests.Trace.FakePrecedentProvider;

namespace Modelwright.Core.Tests.Trace;

public class ReferenceEditTests
{
    private const int Shift = 0x10;
    private const int Ctrl = 0x11;
    private static readonly Dictionary<string, int> Vk = new Dictionary<string, int>
    {
        ["F2"] = 0x71, ["Home"] = 0x24, ["End"] = 0x23, ["Left"] = 0x25, ["Right"] = 0x27, ["F5"] = 0x74,
        ["Enter"] = 0x0D, ["Tab"] = 0x09,
    };

    // Each character typed: its press and release.
    private static List<SyntheticKey> Typed(string text)
    {
        var keys = new List<SyntheticKey>();
        foreach (var character in text)
        {
            keys.Add(SyntheticKey.CharacterDown(character));
            keys.Add(SyntheticKey.CharacterUp(character));
        }

        return keys;
    }

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
    public void The_go_to_step_follows_point_mode_with_F5_the_target_typed_and_Enter()
    {
        // =A14+A2, A2 at 5: select it, Point mode, then Go To A2 (Point mode now moves from A2, not from the edited cell).
        var expected = Keys("F2", "Ctrl+End", "Left*2", "Shift+Right*2", "F2", "F5");
        expected.AddRange(Typed("A2"));
        expected.AddRange(Keys("Enter"));

        Assert.Equal(expected, ReferenceEdit.Keys(Span("=A14+A2", "A2"), "A2"));
        Assert.Equal(expected, ReferenceEdit.Keys(7, 5, 2, "A2"));
        Assert.Equal(SyntheticKey.VkPacket, expected[20].VirtualKey);
        Assert.Equal('A', expected[20].Character);
        Assert.Equal('2', expected[22].Character);
    }

    [Fact]
    public void Another_workbooks_target_is_reached_with_ctrl_tab_then_a_go_to_in_its_window()
    {
        // Go To across windows is unreliable in Point mode: Ctrl+Tab to the target's window (made Excel's previously
        // active one), then Go To on its sheet. Point mode writes it absolute ([Ext.xlsx]Rates!$B$3).
        var span = Span("=[Ext.xlsx]Rates!B3*2", "[Ext.xlsx]Rates!B3");
        var target = new PrecedentItem(PrecedentKind.Cell, "[Ext.xlsx]Rates!B3", "Ext.xlsx", "Rates", "B3");
        var goTo = ReferenceEdit.GoToText(span, target);

        Assert.Equal("'Rates'!B3", goTo);
        Assert.True(ReferenceEdit.SwitchesWindow(span, target));

        var keys = ReferenceEdit.Keys(span, goTo, switchWindow: true)!;
        var expected = Keys("F2", "Ctrl+Home", "Right*1", "Shift+Right*18", "F2", "Ctrl+Tab", "F5");
        expected.AddRange(Typed(goTo!));
        expected.AddRange(Keys("Enter"));

        Assert.Equal(expected, keys);
        Assert.Equal(ReferenceEdit.CtrlTab, Keys("Ctrl+Tab"));
        Assert.Equal(SyntheticKey.Up(Vk["Enter"]), keys[keys.Count - 1]);
        Assert.Single(keys, key => key.Equals(SyntheticKey.Down(Vk["F5"])));
        Assert.Single(keys, key => key.Equals(SyntheticKey.Down(Vk["Tab"])));
    }

    [Fact]
    public void Only_another_workbooks_target_switches_windows()
    {
        var span = Span("=A14+A2", "A2");

        Assert.False(ReferenceEdit.SwitchesWindow(span, new PrecedentItem(PrecedentKind.Cell, "A2", "Model.xlsx", "Calc", "A2")));
        Assert.False(ReferenceEdit.SwitchesWindow(span, new PrecedentItem(PrecedentKind.Cell, "A2", "model.xlsx", "Inputs", "B2")));
        Assert.True(ReferenceEdit.SwitchesWindow(span, new PrecedentItem(PrecedentKind.Range, "x", "Ext.xlsx", "Calc", "B2:B5")));

        // No Go To, no switch.
        Assert.False(ReferenceEdit.SwitchesWindow(span, new PrecedentItem(PrecedentKind.Name, "Rate", "Ext.xlsx", "Rates", "B3")));
        Assert.False(ReferenceEdit.SwitchesWindow(span, new PrecedentItem(PrecedentKind.Cell, "B3", "Ext.xlsx", null, "B3")));
        Assert.Throws<ArgumentNullException>(() => ReferenceEdit.SwitchesWindow(null!, Cell("A2")));
        Assert.Throws<ArgumentNullException>(() => ReferenceEdit.SwitchesWindow(span, null!));

        // Same workbook: no Ctrl+Tab.
        var keys = ReferenceEdit.Keys(span, "A2")!;
        Assert.DoesNotContain(keys, key => key.Equals(SyntheticKey.Down(Vk["Tab"])));
    }

    [Fact]
    public void The_keys_are_sent_in_batches_that_end_after_ctrl_tab_and_after_F5()
    {
        // Up to Ctrl+Tab; F5 alone (once the target's window is in front); the text and Enter (once Go To has the focus).
        var span = Span("=[Ext.xlsx]Rates!B3*2", "[Ext.xlsx]Rates!B3");
        var keys = ReferenceEdit.Keys(span, "'Rates'!B3", switchWindow: true)!;

        var batches = ReferenceEdit.Batches(keys);

        Assert.Equal(3, batches.Count);
        Assert.Equal(Keys("F2", "Ctrl+Home", "Right*1", "Shift+Right*18", "F2", "Ctrl+Tab"), batches[0]);
        Assert.True(ReferenceEdit.EndsWithCtrlTab(batches[0]));
        Assert.Equal(Keys("F5"), batches[1]);
        var third = Typed("'Rates'!B3");
        third.AddRange(Keys("Enter"));
        Assert.Equal(third, batches[2]);
        Assert.Equal(keys, batches.SelectMany(batch => batch));
    }

    [Fact]
    public void Only_a_tab_release_then_a_ctrl_release_ends_with_ctrl_tab()
    {
        Assert.True(ReferenceEdit.EndsWithCtrlTab(ReferenceEdit.CtrlTab));
        Assert.False(ReferenceEdit.EndsWithCtrlTab(Keys("Ctrl+Home")));
        Assert.False(ReferenceEdit.EndsWithCtrlTab(Keys("Tab")));
        Assert.False(ReferenceEdit.EndsWithCtrlTab(new SyntheticKey[0]));
        Assert.Throws<ArgumentNullException>(() => ReferenceEdit.EndsWithCtrlTab(null!));

        // Ctrl+Home in the edit keys does not end a batch.
        Assert.Single(ReferenceEdit.Batches(ReferenceEdit.Keys(Span("=A2+A14", "A2"))!));
    }

    [Fact]
    public void Keys_without_a_go_to_are_one_batch_and_a_final_F5_ends_the_last()
    {
        var keys = ReferenceEdit.Keys(Span("=A14+A2", "A2"))!;

        Assert.Equal(keys, Assert.Single(ReferenceEdit.Batches(keys)));
        Assert.Equal(ReferenceEdit.PlainF2, Assert.Single(ReferenceEdit.Batches(ReferenceEdit.PlainF2)));
        Assert.Single(ReferenceEdit.Batches(Keys("F2", "F5")));
        Assert.Equal(2, ReferenceEdit.Batches(ReferenceEdit.Keys(Span("=A14+A2", "A2"), "A2")!).Count);
        Assert.Empty(ReferenceEdit.Batches(new SyntheticKey[0]));
        Assert.Throws<ArgumentNullException>(() => ReferenceEdit.Batches(null!));
    }

    [Fact]
    public void Ctrl_tab_counts_toward_the_key_event_limit()
    {
        // 10 + 2 x 1,968 moves + 2 x 1 selected = 3,948; Ctrl+Tab 4; F5 and Enter 4; 2 x 22 characters = 44: 4,000.
        var goTo = new string('A', 22);

        Assert.Equal(ReferenceEdit.MaxKeyEvents, ReferenceEdit.Keys(8000, 1968, 1, goTo, switchWindow: true)!.Count);
        Assert.Null(ReferenceEdit.Keys(8000, 1969, 1, goTo, switchWindow: true));
        Assert.NotNull(ReferenceEdit.Keys(8000, 1969, 1, goTo));
    }

    [Fact]
    public void Switching_windows_needs_a_go_to()
    {
        Assert.Throws<ArgumentException>(() => ReferenceEdit.Keys(7, 5, 2, null, switchWindow: true));
        Assert.Throws<ArgumentException>(() => ReferenceEdit.Keys(Span("=A14+A2", "A2"), null, switchWindow: true));
    }

    [Fact]
    public void The_go_to_step_counts_toward_the_key_event_limit()
    {
        // 10 + 2 x 1,970 moves + 2 x 1 selected = 3,952; F5 and Enter 4; 2 x 22 characters = 44: 4,000.
        var goTo = new string('A', 22);

        Assert.Equal(ReferenceEdit.MaxKeyEvents, ReferenceEdit.Keys(8000, 1970, 1, goTo)!.Count);
        Assert.Null(ReferenceEdit.Keys(8000, 1970, 1, goTo + "A"));
        Assert.NotNull(ReferenceEdit.Keys(8000, 1970, 1));
    }

    [Fact]
    public void An_empty_go_to_is_rejected()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => ReferenceEdit.Keys(7, 5, 2, string.Empty));
    }

    [Theory]
    // On the formula's own sheet: the address alone, a range as a range.
    [InlineData("=A14+A2", "A2", "Model.xlsx", "Calc", "A2", "A2")]
    [InlineData("=SUM(B2:B5)", "B2:B5", "Model.xlsx", "Calc", "B2:B5", "B2:B5")]
    [InlineData("=SUM(A:A)", "A:A", "Model.xlsx", "calc", "A:A", "A:A")]
    // Another sheet of the same workbook: always quoted, ' doubled.
    [InlineData("=Inputs!B2*(1+Growth)", "Inputs!B2", "Model.xlsx", "Inputs", "B2", "'Inputs'!B2")]
    [InlineData("='Bob''s Data'!B4", "'Bob''s Data'!B4", "model.xlsx", "Bob's Data", "B4", "'Bob''s Data'!B4")]
    // Another open workbook: its sheet, quoted, with no workbook (the keys switch to its window first), even a sheet
    // named like the formula's.
    [InlineData("=[Ext.xlsx]Rates!B3*2", "[Ext.xlsx]Rates!B3", "Ext.xlsx", "Rates", "B3", "'Rates'!B3")]
    [InlineData("=[Ext.xlsx]Rates!$B$3*2", "[Ext.xlsx]Rates!$B$3", "Ext.xlsx", "Rates", "B3", "'Rates'!B3")]
    [InlineData("='[Bob''s.xlsx]My Rates'!B$3", "'[Bob''s.xlsx]My Rates'!B$3", "Bob's.xlsx", "My Rates", "B3", "'My Rates'!B3")]
    [InlineData("=[Ext.xlsx]Calc!A2", "[Ext.xlsx]Calc!A2", "Ext.xlsx", "Calc", "A2", "'Calc'!A2")]
    // Anchored ($), even the two ends of a range anchored differently: the same Go To (Point mode's anchoring results).
    [InlineData("=$A$2+1", "$A$2", "Model.xlsx", "Calc", "A2", "A2")]
    [InlineData("=A$2+1", "A$2", "Model.xlsx", "Calc", "A2", "A2")]
    [InlineData("=SUM(Data!$B$2:$B$5)", "Data!$B$2:$B$5", "Model.xlsx", "Data", "B2:B5", "'Data'!B2:B5")]
    [InlineData("=SUM($B$2:B5)", "$B$2:B5", "Model.xlsx", "Calc", "B2:B5", "B2:B5")]
    [InlineData("=SUM($A:$A)", "$A:$A", "Model.xlsx", "Calc", "A:A", "A:A")]
    [InlineData("=SUM($1:$1)", "$1:$1", "Model.xlsx", "Calc", "1:1", "1:1")]
    [InlineData("=SUM([Ext.xlsx]Rates!A:A)", "[Ext.xlsx]Rates!A:A", "Ext.xlsx", "Rates", "A:A", "'Rates'!A:A")]
    public void The_go_to_text_is_the_resolved_target(
        string formula, string reference, string workbook, string sheet, string address, string goTo)
    {
        var target = new PrecedentItem(address.IndexOf(':') < 0 ? PrecedentKind.Cell : PrecedentKind.Range, reference, workbook, sheet, address);

        Assert.Equal(goTo, ReferenceEdit.GoToText(Span(formula, reference), target));
    }

    [Theory]
    // A spill: Go To would write the spill range.
    [InlineData("=SUM(A1#)", "A1#", PrecedentKind.Range, "Calc", "A1:A5")]
    // 3-D, and several areas.
    [InlineData("=SUM(S1:S3!A1)", "S1:S3!A1", PrecedentKind.Cell, "S1:S3", "A1")]
    [InlineData("=SUM(Calc!A1)", "Calc!A1", PrecedentKind.Range, "Calc", "A1:A5,C1:C5")]
    // A name or table reference: Go To would replace it with an address.
    [InlineData("=Rate*2", "Rate", PrecedentKind.Name, "Calc", "B3")]
    [InlineData("=SUM(Sales[Amount])", "Sales[Amount]", PrecedentKind.Table, "Data", "G2:G6")]
    public void There_is_no_go_to_where_it_would_not_write_back_the_same_reference(
        string formula, string reference, PrecedentKind kind, string sheet, string address)
    {
        var target = new PrecedentItem(kind, reference, "Model.xlsx", sheet, address);

        Assert.Null(ReferenceEdit.GoToText(Span(formula, reference), target));
    }

    [Fact]
    public void There_is_no_go_to_without_a_place_or_longer_than_the_dialog_holds()
    {
        var span = Span("=A2+1", "A2");

        Assert.Null(ReferenceEdit.GoToText(span, new PrecedentItem(PrecedentKind.Cell, "A2")));
        Assert.Null(ReferenceEdit.GoToText(span, new PrecedentItem(PrecedentKind.Cell, "A2", "Model.xlsx", null, "A2")));
        Assert.Null(ReferenceEdit.GoToText(span, new PrecedentItem(PrecedentKind.Cell, "A2", "Model.xlsx", new string('S', 300), "A2")));
        Assert.Throws<ArgumentNullException>(() => ReferenceEdit.GoToText(null!, Cell("A2")));
        Assert.Throws<ArgumentNullException>(() => ReferenceEdit.GoToText(span, null!));
    }

    [Fact]
    public void The_target_of_a_cell_in_a_range_is_the_range()
    {
        var root = Cell("A1");
        var range = WithSpan(Range("A1:A3", 3, "Data"), Span("=SUM(Data!A1:A3)", "Data!A1:A3"));
        var tree = new PrecedentTree(new FakePrecedentProvider().Has(root, range), root);
        tree.MoveDown();
        tree.MoveRight();
        tree.MoveRight(); // the range's first cell

        var node = ReferenceEdit.SpanNodeOf(tree.Selected);

        Assert.Equal(PrecedentKind.Cell, tree.Selected.Item.Kind);
        Assert.Same(range, node!.Item);
        Assert.Equal("'Data'!A1:A3", ReferenceEdit.GoToText(node.Item.Span!, node.Item));
        Assert.Null(ReferenceEdit.SpanNodeOf(tree.Root));
    }

    [Fact]
    public void A_reference_is_found_again_once_its_workbook_path_is_no_longer_written()
    {
        // B11 as traced with the external workbook closed, and as Excel writes it once the trace has opened it.
        const string traced = @"='C:\Temp\fix\[EMT_TraceExternal.xlsx]Rates'!B3*2+A1";
        const string open = "=[EMT_TraceExternal.xlsx]Rates!B3*2+A1";
        var span = Span(traced, @"'C:\Temp\fix\[EMT_TraceExternal.xlsx]Rates'!B3", "B11");

        var moved = ReferenceEdit.Relocate(span, open);

        Assert.NotNull(moved);
        Assert.Equal("[EMT_TraceExternal.xlsx]Rates!B3", moved!.Text);
        Assert.Equal(open, moved.Formula);
        Assert.Equal(("Model.xlsx", "Calc", "B11"), (moved.Workbook, moved.Sheet, moved.Address));

        var after = ReferenceEdit.Relocate(Span(traced, "A1", "B11"), open);
        Assert.Equal("A1", after!.Text);
        Assert.Equal(open.LastIndexOf("A1", StringComparison.Ordinal), after.Start);
    }

    [Fact]
    public void A_formula_changed_otherwise_has_no_reference_to_find()
    {
        var span = Span("=A14+A2", "A2");

        Assert.Same(span, ReferenceEdit.Relocate(span, "=A14+A2"));
        Assert.Null(ReferenceEdit.Relocate(span, "=A14+A3"));
        Assert.Null(ReferenceEdit.Relocate(span, "=A14-A2"));
        Assert.Null(ReferenceEdit.Relocate(span, "=A14+A2+A5"));
        Assert.Null(ReferenceEdit.Relocate(span, "=A14+$A$2"));
        Assert.Null(ReferenceEdit.Relocate(span, "=A14+A2)"));
        Assert.Null(ReferenceEdit.Relocate(span, "2"));
        Assert.Throws<ArgumentNullException>(() => ReferenceEdit.Relocate(span, null!));
        Assert.Throws<ArgumentNullException>(() => ReferenceEdit.Relocate(null!, "=A1"));
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

    [Fact]
    public void Keys_that_would_cross_a_surrogate_pair_are_not_used()
    {
        const string Emoji = "😀";

        // From the start (the reference is in the first half): the caret crosses the emoji before it.
        var before = "=\"" + Emoji + "\"&B2&\"a long text after the reference\"";
        Assert.True(ReferenceEdit.CrossesSurrogatePair(Span(before, "B2", "B2")));

        // From the end: the caret crosses only what follows the reference's start.
        Assert.False(ReferenceEdit.CrossesSurrogatePair(Span("=\"" + Emoji + "\"&B2", "B2", "B2")));
        Assert.True(ReferenceEdit.CrossesSurrogatePair(Span("=\"a long text before the reference\"&B2&\"" + Emoji + "\"", "B2", "B2")));

        // In the reference itself (a sheet name).
        Assert.True(ReferenceEdit.CrossesSurrogatePair(Span("='" + Emoji + "'!A1", "'" + Emoji + "'!A1")));
        Assert.False(ReferenceEdit.CrossesSurrogatePair(Span("=A1+B2", "B2", "B2")));
        Assert.Throws<ArgumentNullException>(() => ReferenceEdit.CrossesSurrogatePair(null!));
    }
}
