using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Modelwright.Core.Trace;
using Xunit;

namespace Modelwright.Core.Tests.Trace;

public class EvaluatePrecedentsTests
{
    private static readonly FormulaContext Context = new FormulaContext("Model.xlsx", "Calc");

    // The formula from Macabacus's Trace In help page (research/07), with inputs that give its values there.
    private const string MacabacusExample =
        "=(B2+C2/D2)+IF(E2>0,F2+G2,SUM(H2:J2))*(K2+L2-ABS(M2))+PRODUCT(N2:T2,U2)+V2";

    // What Worksheet.Evaluate would give for each subexpression of the example (B2=4, C2=3, D2=2, E2=1, F2=10, G2=4.9,
    // H2:J2=5,7,8, K2=1, L2=2, M2=-7, N2:T2=1..7, U2=3).
    private static readonly Dictionary<string, string> ExampleValues = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        ["(B2+C2/D2)"] = "5.5",
        ["IF(E2>0,F2+G2,SUM(H2:J2))"] = "14.9",
        ["E2>0"] = "TRUE",
        ["F2+G2"] = "14.9",
        ["SUM(H2:J2)"] = "20",
        ["(K2+L2-ABS(M2))"] = "-4",
        ["ABS(M2)"] = "7",
        ["PRODUCT(N2:T2,U2)"] = "15120",
    };

    private static ParsedFormula Parse(string formula)
    {
        var parsed = FormulaParser.Parse(formula, Context);
        Assert.True(parsed.IsParsed, parsed.Error);
        return parsed;
    }

    private static string? Value(FormulaNode node) => ExampleValues.TryGetValue(node.Text, out var value) ? value : "?";

    // The classic mapping, as the add-in's would give it for input cells (without Excel: no value, nothing below).
    private static PrecedentItem Reference(FormulaNode node, string? argument) =>
        PrecedentItem.FromReference(node.Reference!, Context, argument: argument, canExpand: false);

    // The whole tree below a cell, every row expanded: "kind label [argument] = value", indented two spaces a level,
    // with "+" before a row that can be expanded.
    private static string Tree(string formula)
    {
        var text = new StringBuilder();
        Add(EvaluatePrecedents.Of(Parse(formula), Value, Reference), 0, text);
        return text.ToString();
    }

    private static void Add(IReadOnlyList<EvaluatePrecedent> rows, int depth, StringBuilder text)
    {
        foreach (var row in rows)
        {
            var item = row.Item;
            text.Append(' ', depth * 2).Append(item.CanExpand ? "+ " : "  ").Append(item.Kind).Append(' ').Append(item.Label);
            if (item.Argument is not null)
            {
                text.Append(" [").Append(item.Argument).Append(']');
            }

            if (item.ValueText is not null)
            {
                text.Append(" = ").Append(item.ValueText);
            }

            text.Append('\n');
            if (item.Kind == PrecedentKind.Function || item.Kind == PrecedentKind.Group)
            {
                Add(EvaluatePrecedents.Of(row.Node, Value, Reference), depth + 1, text);
            }
        }
    }

    [Fact]
    public void Macabacus_example_gives_the_documented_tree()
    {
        // research/07: (x) (B2+C2/D2) 5.5; ƒx IF(...) 14.9 with (x) E2>0 logical_test True, (x) F2+G2
        // [value_if_true] 14.9, ƒx SUM(...) [value_if_false] 20; (x) (K2+L2-ABS(M2)) (4); ƒx PRODUCT(...) 15,120; V2.
        Assert.Equal(
            "+ Group (B2+C2/D2) = 5.5\n" +
            "    Cell B2\n" +
            "    Cell C2\n" +
            "    Cell D2\n" +
            "+ Function IF(...) = 14.9\n" +
            "  + Group E2>0 [logical_test] = TRUE\n" +
            "      Cell E2\n" +
            "  + Group F2+G2 [[value_if_true]] = 14.9\n" +
            "      Cell F2\n" +
            "      Cell G2\n" +
            "  + Function SUM(...) [[value_if_false]] = 20\n" +
            "      Range H2:J2 [number1]\n" +
            "+ Group (K2+L2-ABS(M2)) = -4\n" +
            "    Cell K2\n" +
            "    Cell L2\n" +
            "  + Function ABS(...) = 7\n" +
            "      Cell M2 [number]\n" +
            "+ Function PRODUCT(...) = 15120\n" +
            "    Range N2:T2 [number1]\n" +
            "    Cell U2 [[number2]]\n" +
            "  Cell V2\n",
            Tree(MacabacusExample));
    }

    [Fact]
    public void Top_level_rows_follow_the_formula_structure()
    {
        var rows = EvaluatePrecedents.Of(Parse(MacabacusExample), Value, Reference);

        Assert.Equal(new[] { "(B2+C2/D2)", "IF(...)", "(K2+L2-ABS(M2))", "PRODUCT(...)", "V2" }, rows.Select(r => r.Item.Label));
        Assert.Equal(
            new[] { PrecedentKind.Group, PrecedentKind.Function, PrecedentKind.Group, PrecedentKind.Function, PrecedentKind.Cell },
            rows.Select(r => r.Item.Kind));
        Assert.All(rows, r => Assert.Null(r.Item.Argument));
        Assert.All(rows.Take(4), r => Assert.True(r.Item.CanExpand));
        Assert.Equal("IF(E2>0,F2+G2,SUM(H2:J2))", rows[1].Node.Text);
        Assert.Equal("V2", rows[4].Item.Address);
        Assert.Equal("Calc", rows[4].Item.Sheet);
    }

    [Fact]
    public void Function_and_group_rows_have_no_place_and_are_never_cycles()
    {
        var rows = EvaluatePrecedents.Of(Parse(MacabacusExample), Value, Reference);

        Assert.All(rows.Take(4), r =>
        {
            Assert.Null(r.Item.Id);
            Assert.Null(r.Item.Address);
            Assert.Null(r.Item.Span);
        });
    }

    [Fact]
    public void Values_are_asked_for_only_the_rows_being_built()
    {
        var asked = new List<string>();
        var rows = EvaluatePrecedents.Of(Parse(MacabacusExample), node =>
        {
            asked.Add(node.Text);
            return "1";
        }, Reference);

        Assert.Equal(new[] { "(B2+C2/D2)", "IF(E2>0,F2+G2,SUM(H2:J2))", "(K2+L2-ABS(M2))", "PRODUCT(N2:T2,U2)" }, asked);

        asked.Clear();
        EvaluatePrecedents.Of(rows[1].Node, node =>
        {
            asked.Add(node.Text);
            return "1";
        }, Reference);

        Assert.Equal(new[] { "E2>0", "F2+G2", "SUM(H2:J2)" }, asked);
    }

    [Fact]
    public void References_are_mapped_by_the_caller_with_their_argument_name()
    {
        var mapped = new List<string>();
        var sum = EvaluatePrecedents.Of(Parse("=SUM(A1,Sheet2!B2:B5,Rate)"), Value, (node, argument) =>
        {
            mapped.Add(node.Text + "|" + argument);
            return Reference(node, argument);
        }).Single();

        Assert.Empty(mapped); // the function row itself needs no reference
        var arguments = EvaluatePrecedents.Of(sum.Node, Value, (node, argument) =>
        {
            mapped.Add(node.Text + "|" + argument);
            return Reference(node, argument);
        });

        Assert.Equal(new[] { "A1|number1", "Sheet2!B2:B5|[number2]", "Rate|[number3]" }, mapped);
        Assert.Equal(new[] { PrecedentKind.Cell, PrecedentKind.Range, PrecedentKind.Name }, arguments.Select(r => r.Item.Kind));
        Assert.Equal(new[] { "number1", "[number2]", "[number3]" }, arguments.Select(r => r.Item.Argument));
    }

    [Fact]
    public void Repeating_parameters_continue_their_numbering()
    {
        var sumifs = EvaluatePrecedents.Of(Parse("=SUMIFS(D:D,A:A,\"x\",B:B,\">0\",C:C,F1)"), Value, Reference).Single();

        Assert.Equal(
            new[] { "sum_range", "criteria_range1", "criteria1", "[criteria_range2]", "[criteria2]", "[criteria_range3]", "[criteria3]" },
            EvaluatePrecedents.Of(sumifs.Node, Value, Reference).Select(r => r.Item.Argument));
    }

    [Fact]
    public void Constant_and_omitted_arguments_are_rows_that_cannot_be_expanded()
    {
        var call = EvaluatePrecedents.Of(Parse("=IF(A1,,\"say \"\"hi\"\"\")+ROUND(B1,2)"), Value, Reference);
        var ifArguments = EvaluatePrecedents.Of(call[0].Node, _ => throw new InvalidOperationException("a constant is not evaluated"), Reference);
        var roundArguments = EvaluatePrecedents.Of(call[1].Node, Value, Reference);

        Assert.Equal(new[] { "A1", EvaluatePrecedents.OmittedLabel, "\"say \"\"hi\"\"\"" }, ifArguments.Select(r => r.Item.Label));
        Assert.Equal(new[] { "logical_test", "[value_if_true]", "[value_if_false]" }, ifArguments.Select(r => r.Item.Argument));
        Assert.Equal(string.Empty, ifArguments[1].Item.ValueText);
        Assert.Equal("say \"hi\"", ifArguments[2].Item.ValueText);
        Assert.False(ifArguments[1].Item.CanExpand);
        Assert.False(ifArguments[2].Item.CanExpand);
        Assert.Equal(PrecedentKind.Group, ifArguments[2].Item.Kind);
        Assert.Equal("2", roundArguments[1].Item.ValueText);
        Assert.Equal("num_digits", roundArguments[1].Item.Argument);
    }

    [Fact]
    public void A_group_of_constants_cannot_be_expanded()
    {
        var rows = EvaluatePrecedents.Of(Parse("=A1*(1+2)"), _ => "3", Reference);

        Assert.Equal(new[] { "A1", "(1+2)" }, rows.Select(r => r.Item.Label));
        Assert.False(rows[1].Item.CanExpand);
        Assert.Equal("3", rows[1].Item.ValueText);
    }

    [Fact]
    public void Local_names_are_left_out()
    {
        var let = EvaluatePrecedents.Of(Parse("=LET(x,A1,x*2)"), Value, Reference).Single();
        var arguments = EvaluatePrecedents.Of(let.Node, Value, Reference);

        Assert.Equal(new[] { "A1", "x*2" }, arguments.Select(r => r.Item.Label));
        Assert.Equal(new[] { "name_value1", "calculation_or_name2" }, arguments.Select(r => r.Item.Argument));
        Assert.False(arguments[1].Item.CanExpand); // x*2 holds only the local name
    }

    [Fact]
    public void A_function_without_arguments_is_written_with_empty_parentheses()
    {
        var rows = EvaluatePrecedents.Of(Parse("=TODAY()-A1"), _ => "45000", Reference);

        Assert.Equal("TODAY()", rows[0].Item.Label);
        Assert.False(rows[0].Item.CanExpand);
        Assert.Equal("45000", rows[0].Item.ValueText);
    }

    [Fact]
    public void Reference_returning_functions_are_function_rows_whose_children_are_their_arguments()
    {
        var index = EvaluatePrecedents.Of(Parse("=INDEX(Data!A1:A60,5)"), Value, Reference).Single();

        Assert.True(index.Node.ReturnsReference);
        Assert.Equal(PrecedentKind.Function, index.Item.Kind);
        Assert.Equal("INDEX(...)", index.Item.Label);
        Assert.Equal(new[] { "array", "row_num" }, EvaluatePrecedents.Of(index.Node, Value, Reference).Select(r => r.Item.Argument));
    }

    [Fact]
    public void Newer_functions_are_labelled_without_their_prefix_and_labels_are_one_line()
    {
        var rows = EvaluatePrecedents.Of(Parse("=_xlfn.XLOOKUP(A1,B:B,C:C)+(A2\n+A3)"), _ => "1", Reference);

        Assert.Equal(new[] { "XLOOKUP(...)", "(A2 +A3)" }, rows.Select(r => r.Item.Label));
    }

    [Fact]
    public void A_formula_that_was_not_parsed_has_no_rows()
    {
        Assert.Empty(EvaluatePrecedents.Of(FormulaParser.Parse("=SUM(", Context), Value, Reference));
        Assert.Empty(EvaluatePrecedents.Of(Parse("=1+2"), Value, Reference));
    }

    [Fact]
    public void Arguments_are_checked()
    {
        var parsed = Parse("=A1");
        var node = parsed.TopLevelNodes[0];

        Assert.Throws<ArgumentNullException>(() => EvaluatePrecedents.Of((ParsedFormula)null!, Value, Reference));
        Assert.Throws<ArgumentNullException>(() => EvaluatePrecedents.Of((FormulaNode)null!, Value, Reference));
        Assert.Throws<ArgumentNullException>(() => EvaluatePrecedents.Of(parsed, null!, Reference));
        Assert.Throws<ArgumentNullException>(() => EvaluatePrecedents.Of(node, Value, null!));
        Assert.Throws<ArgumentNullException>(() => EvaluatePrecedents.Label(null!));
    }

    [Fact]
    public void With_argument_changes_only_the_argument()
    {
        var span = new ReferenceSpan("Model.xlsx", "Calc", "A1", "=B2", 1, 2);
        var item = new PrecedentItem(PrecedentKind.Cell, "B2", "Model.xlsx", "Calc", "B2", valueText: "5", hiddenNote: "hidden rows")
            .WithSpan(span);

        var copy = item.WithArgument("[value_if_true]");

        Assert.Equal("[value_if_true]", copy.Argument);
        Assert.Null(copy.WithArgument(null).Argument);
        Assert.Equal(item.Label, copy.Label);
        Assert.Equal(item.Id, copy.Id);
        Assert.Equal(item.ValueText, copy.ValueText);
        Assert.Equal(item.HiddenNote, copy.HiddenNote);
        Assert.Same(span, copy.Span);
        Assert.NotSame(item, copy);
    }
}
