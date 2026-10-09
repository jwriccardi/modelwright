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

    // The rows under a formula, then under each of its top-level rows, with each value asked recorded ("asked") and
    // given as "=text".
    private static (IReadOnlyList<EvaluatePrecedent> Top, IReadOnlyList<IReadOnlyList<EvaluatePrecedent>> Children) Expand(
        string formula, List<string> asked, Func<string>? cellValue = null)
    {
        string? ValueOf(FormulaNode node)
        {
            asked.Add(node.Text);
            return "=" + node.Text;
        }

        var top = EvaluatePrecedents.Of(Parse(formula), ValueOf, Reference, cellValue);
        var children = top.Select(row => row.Node.Kind == FormulaNodeKind.Reference
            ? (IReadOnlyList<EvaluatePrecedent>)new EvaluatePrecedent[0]
            : EvaluatePrecedents.Of(row.Node, ValueOf, Reference)).ToList();
        return (top, children);
    }

    [Fact]
    public void A_row_using_a_let_name_is_not_evaluated()
    {
        var asked = new List<string>();
        var (top, children) = Expand("=LET(x,A1,x*2)+1", asked);

        // LET(...) declares x itself, so it is evaluated; x*2 alone would be #NAME? (or a defined name x's value).
        Assert.Equal("=LET(x,A1,x*2)", top.Single().Item.ValueText);
        Assert.Equal(new[] { "A1", "x*2" }, children[0].Select(r => r.Item.Label));
        Assert.Equal(EvaluatePrecedents.UsesLocalNames, children[0][1].Item.ValueText);
        Assert.DoesNotContain("x*2", asked);
        Assert.Equal(new[] { "LET(x,A1,x*2)" }, asked);
    }

    [Theory]
    [InlineData("=LET(_xlpm.x,A1,_xlpm.x*2)+1", "_xlpm.x*2")]
    [InlineData("=LET(x,A1,y,B1,SUM(x,y))+1", "SUM(x,y)")]
    [InlineData("=LET(x,A1,y,x+1,(y+A2)*2)+1", "(y+A2)*2")]
    [InlineData("=LET(_xlpm.f,LAMBDA(_xlpm.y,_xlpm.y+1),_xlpm.f(A1))+1", "_xlpm.f(A1)")]
    public void Rows_using_names_declared_outside_them_are_not_evaluated(string formula, string fragment)
    {
        var asked = new List<string>();
        var (top, children) = Expand(formula, asked);

        var row = children[0].Single(r => r.Node.Text == fragment);
        Assert.Equal(EvaluatePrecedents.UsesLocalNames, row.Item.ValueText);
        Assert.DoesNotContain(fragment, asked);
        Assert.Equal(EvaluatePrecedents.UsesLocalNames, EvaluatePrecedents.NotEvaluated(row.Node));
        Assert.Null(EvaluatePrecedents.NotEvaluated(top[0].Node)); // the LET declares them
    }

    [Fact]
    public void A_lambda_called_in_place_is_one_function_row()
    {
        var asked = new List<string>();
        var (top, children) = Expand("=LAMBDA(x,y,x+y)(A1,B1)*2", asked);

        var call = top.Single();
        Assert.True(call.Node.IsLambdaCall);
        Assert.Equal(PrecedentKind.Function, call.Item.Kind);
        Assert.Equal(EvaluatePrecedents.LambdaCallLabel, call.Item.Label);
        Assert.Equal("=LAMBDA(x,y,x+y)(A1,B1)", call.Item.ValueText); // self-contained: evaluated
        Assert.True(call.Item.CanExpand);

        // Its children: the LAMBDA (a function value: not evaluated), then the arguments passed to it.
        Assert.Equal(new[] { "LAMBDA(...)", "A1", "B1" }, children[0].Select(r => r.Item.Label));
        Assert.Equal(EvaluatePrecedents.UsesLocalNames, children[0][0].Item.ValueText);
        Assert.Equal(new[] { "LAMBDA(x,y,x+y)(A1,B1)" }, asked);

        // Inside the LAMBDA: the parameters are left out, the calculation uses them.
        var inside = EvaluatePrecedents.Of(children[0][0].Node, _ => throw new InvalidOperationException("not evaluated"), Reference);
        Assert.Equal(EvaluatePrecedents.UsesLocalNames, inside.Single().Item.ValueText);
    }

    [Fact]
    public void A_lambda_followed_by_a_space_is_an_intersection_not_a_call()
    {
        var rows = EvaluatePrecedents.Of(Parse("=LAMBDA(x,x*2) (A1)"), _ => "1", Reference);

        Assert.Equal(new[] { "LAMBDA(...)", "(A1)" }, rows.Select(r => r.Item.Label));
        Assert.Equal(EvaluatePrecedents.UsesLocalNames, rows[0].Item.ValueText);
    }

    [Fact]
    public void A_lambda_passed_to_a_function_does_not_stop_the_call_being_evaluated()
    {
        var asked = new List<string>();
        var (top, children) = Expand("=SUM(MAP(A1:A3,LAMBDA(x,x*2)))+1", asked);
        var map = EvaluatePrecedents.Of(children[0].Single().Node, node =>
        {
            asked.Add(node.Text);
            return "1";
        }, Reference);

        Assert.Contains("MAP(A1:A3,LAMBDA(x,x*2))", asked);
        Assert.Equal(new[] { "A1:A3", "LAMBDA(...)" }, map.Select(r => r.Item.Label));
        Assert.Equal(EvaluatePrecedents.UsesLocalNames, map[1].Item.ValueText);
        Assert.DoesNotContain("LAMBDA(x,x*2)", asked);
    }

    [Theory]
    [InlineData("=MyUdf(A1)+1", "MyUdf(A1)")]
    [InlineData("=Book.xlsx!MyUdf(A1)+1", "Book.xlsx!MyUdf(A1)")]
    [InlineData("=SUM(A1,MyUdf(B1))+1", "SUM(A1,MyUdf(B1))")]
    [InlineData("=(A1+_xll.Price(B1))*2", "(A1+_xll.Price(B1))")]
    [InlineData("=NamedLambda(A1)+1", "NamedLambda(A1)")]
    public void Rows_calling_a_function_that_is_not_excels_are_not_evaluated(string formula, string fragment)
    {
        var asked = new List<string>();
        var (top, _) = Expand(formula, asked);

        Assert.Equal(EvaluatePrecedents.UnknownFunction, top.Single(r => r.Node.Text == fragment).Item.ValueText);
        Assert.Empty(asked);
    }

    [Theory]
    [InlineData("=WEBSERVICE(A1)&\"\"", "WEBSERVICE(A1)", "WEBSERVICE")]
    [InlineData("=_xlfn.STOCKHISTORY(A1,B1)+0", "_xlfn.STOCKHISTORY(A1,B1)", "STOCKHISTORY")]
    [InlineData("=RTD(\"srv\",,A1)+0", "RTD(\"srv\",,A1)", "RTD")]
    [InlineData("=(CUBEVALUE(\"c\",A1)+1)*2", "(CUBEVALUE(\"c\",A1)+1)", "CUBEVALUE")]
    public void Rows_reaching_outside_the_workbook_are_not_evaluated(string formula, string fragment, string function)
    {
        var asked = new List<string>();
        var (top, _) = Expand(formula, asked);

        var value = top.Single(r => r.Node.Text == fragment).Item.ValueText;
        Assert.Equal(EvaluatePrecedents.ReachesOutside(function), value);
        Assert.Contains(function + " reads external data", value);
        Assert.Empty(asked);
    }

    [Fact]
    public void Excel_functions_old_and_new_are_evaluated()
    {
        var asked = new List<string>();
        Expand("=XLOOKUP(A1,B:B,C:C)+_xlfn.TEXTSPLIT(A2,\",\")+IFERROR(A3,0)+NOW()", asked);

        Assert.Equal(new[] { "XLOOKUP(A1,B:B,C:C)", "_xlfn.TEXTSPLIT(A2,\",\")", "IFERROR(A3,0)", "NOW()" }, asked);
    }

    [Fact]
    public void The_whole_formula_shows_the_cell_value_without_evaluating()
    {
        var asked = new List<string>();
        var (top, children) = Expand("=IF(A1>0,B1+1,C1)", asked, () => "$15,066");

        Assert.Equal("$15,066", top.Single().Item.ValueText);
        Assert.Equal(new[] { "A1>0", "B1+1" }, asked); // only its arguments, when expanded
        Assert.Equal("=A1>0", children[0][0].Item.ValueText);

        asked.Clear();
        Assert.Equal("cell", Expand("=MyUdf(A1)", asked, () => "cell").Top.Single().Item.ValueText);
        Assert.Equal("cell", Expand("=LET(x,A1,x*2)", asked, () => "cell").Top.Single().Item.ValueText);
        Assert.Equal("cell", Expand("=(A1+B1)", asked, () => "cell").Top.Single().Item.ValueText);
        Assert.Empty(asked);

        // Without the cell's value, it is evaluated like any other row.
        Assert.Equal("=IF(A1>0,B1+1,C1)", Expand("=IF(A1>0,B1+1,C1)", asked).Top.Single().Item.ValueText);
    }

    [Fact]
    public void A_whole_formula_returning_a_reference_is_evaluated_to_find_its_range()
    {
        var asked = new List<string>();
        var index = Expand("=INDEX(A1:A9,2)", asked, () => "cell").Top.Single();

        Assert.Equal("=INDEX(A1:A9,2)", index.Item.ValueText);
        Assert.Equal(new[] { "INDEX(A1:A9,2)" }, asked);

        // Too long to evaluate: the cell's value.
        var tooLong = EvaluatePrecedents.Of(Parse("=INDEX(A1:A9,2)"), _ => EvaluatePrecedents.TooLong, Reference, () => "cell").Single();
        Assert.Equal("cell", tooLong.Item.ValueText);
    }

    private static FormulaNode Node(string formula, string fragment)
    {
        var pending = new Stack<FormulaNode>();
        pending.Push(Parse(formula).Structure!);
        while (pending.Count > 0)
        {
            var node = pending.Pop();
            if (node.Text == fragment)
            {
                return node;
            }

            foreach (var child in node.Children)
            {
                pending.Push(child);
            }
        }

        throw new InvalidOperationException(fragment + " is not a node of " + formula);
    }

    private static readonly Func<string, bool> AllOpen = _ => true;

    private static readonly Func<FormulaReference, string?> AsWritten = reference => reference.Text;

    [Fact]
    public void Prepare_gives_the_text_with_row_and_column_made_the_cells()
    {
        var request = EvaluatePrecedents.Prepare(Node("=(ROW()+COLUMN()+A1)*2", "(ROW()+COLUMN()+A1)"), 7, 3, AllOpen, AsWritten);

        Assert.Equal("(7+3+A1)", request.Text);
        Assert.Null(request.NotEvaluated);
        Assert.False(request.IsVolatile);
        Assert.Equal("5", request.Finish("5"));
    }

    [Theory]
    [InlineData("=LET(_xlpm.x,A1,_xlpm.x*2)+1", "LET(_xlpm.x,A1,_xlpm.x*2)", "LET(x,A1,x*2)")]
    [InlineData("=LET(_xlpm.f,LAMBDA(_xlpm.y,_xlpm.y+1),_xlpm.f(A1))+1", "LET(_xlpm.f,LAMBDA(_xlpm.y,_xlpm.y+1),_xlpm.f(A1))",
        "LET(f,LAMBDA(y,y+1),f(A1))")]
    [InlineData("=LET(x,A1,x*2)+1", "LET(x,A1,x*2)", "LET(x,A1,x*2)")]
    public void Prepare_writes_let_and_lambda_names_as_typed(string formula, string fragment, string expected)
    {
        Assert.Equal(expected, EvaluatePrecedents.Prepare(Node(formula, fragment), 1, 1, AllOpen, AsWritten).Text);
    }

    [Fact]
    public void Prepare_marks_volatile_rows()
    {
        var request = EvaluatePrecedents.Prepare(Node("=(NOW()+A1)*2", "(NOW()+A1)"), 1, 1, AllOpen, AsWritten);

        Assert.True(request.IsVolatile);
        Assert.Equal("45000 (volatile)", request.Finish("45000"));
        Assert.True(EvaluatePrecedents.Prepare(Node("=_xlfn.RANDARRAY(2)+1", "_xlfn.RANDARRAY(2)"), 1, 1, AllOpen, AsWritten).IsVolatile);
    }

    [Fact]
    public void Prepare_does_not_evaluate_a_closed_workbook()
    {
        const string Formula = "=SUM('C:\\x\\[Ext.xlsx]Rates'!B3:B5)+1";
        var node = Node(Formula, "SUM('C:\\x\\[Ext.xlsx]Rates'!B3:B5)");
        var checkedBooks = new List<string>();

        var closed = EvaluatePrecedents.Prepare(node, 1, 1, name =>
        {
            checkedBooks.Add(name);
            return false;
        }, AsWritten);
        var open = EvaluatePrecedents.Prepare(node, 1, 1, AllOpen, AsWritten);

        Assert.Null(closed.Text);
        Assert.Equal(EvaluatePrecedents.ClosedWorkbook("Ext.xlsx"), closed.NotEvaluated);
        Assert.Contains("[Ext.xlsx] is closed", closed.NotEvaluated);
        Assert.Contains("trace again", closed.NotEvaluated);
        Assert.Equal(new[] { "Ext.xlsx" }, checkedBooks);
        Assert.Equal("SUM('C:\\x\\[Ext.xlsx]Rates'!B3:B5)", open.Text);
    }

    [Fact]
    public void Prepare_writes_table_references_to_the_current_row_as_their_cells()
    {
        var asked = new List<string>();
        var request = EvaluatePrecedents.Prepare(Node("=([@Qty]*[@Price])+1", "([@Qty]*[@Price])"), 7, 3, AllOpen, reference =>
        {
            asked.Add(reference.Text);
            return reference.TableColumns[0] == "Qty" ? "'Eval'!$A$7" : "'Eval'!$B$7";
        });

        Assert.Equal("('Eval'!$A$7*'Eval'!$B$7)", request.Text);
        Assert.Equal(new[] { "[@Qty]", "[@Price]" }, asked);
    }

    [Fact]
    public void Prepare_asks_only_for_table_references_that_need_the_cell_and_for_names()
    {
        var asked = new List<string>();
        var request = EvaluatePrecedents.Prepare(
            Node("=SUM(Sales[Qty],Sales[@Qty],Sales[[#This Row],[Qty]],[Qty],Rate,A1)+1", "SUM(Sales[Qty],Sales[@Qty],Sales[[#This Row],[Qty]],[Qty],Rate,A1)"),
            3, 1, AllOpen, reference =>
            {
                asked.Add(reference.Text);
                return reference.Kind == FormulaReferenceKind.Name ? reference.Text : "X";
            });

        Assert.Equal(new[] { "Sales[@Qty]", "Sales[[#This Row],[Qty]]", "[Qty]", "Rate" }, asked);
        Assert.Equal("SUM(Sales[Qty],X,X,X,Rate,A1)", request.Text);
    }

    [Theory]
    [InlineData("=([@Qty]*2)+1", "([@Qty]*2)")]
    [InlineData("=(RelativeName*2)+1", "(RelativeName*2)")]
    [InlineData("=(@A1:A10+1)*2", "(@A1:A10+1)")]
    [InlineData("=(_xlfn.SINGLE(A1:A10)+1)*2", "(_xlfn.SINGLE(A1:A10)+1)")]
    public void Prepare_does_not_evaluate_what_depends_on_its_cell(string formula, string fragment)
    {
        var request = EvaluatePrecedents.Prepare(Node(formula, fragment), 1, 1, AllOpen, _ => null);

        Assert.Null(request.Text);
        Assert.Equal(EvaluatePrecedents.DependsOnItsCell, request.NotEvaluated);
    }

    [Fact]
    public void Prepare_reports_text_longer_than_excel_evaluates()
    {
        var sum = "SUM(" + string.Join(",", Enumerable.Repeat("B2", 90)) + ")";
        var request = EvaluatePrecedents.Prepare(Node("=" + sum + "+1", sum), 1, 1, AllOpen, AsWritten);

        Assert.Null(request.Text);
        Assert.Equal(EvaluatePrecedents.TooLong, request.NotEvaluated);
        Assert.Contains("255-character", EvaluatePrecedents.TooLong);
    }

    [Fact]
    public void Prepare_gives_the_reasons_rows_are_never_evaluated()
    {
        Assert.Equal(EvaluatePrecedents.UsesLocalNames,
            EvaluatePrecedents.Prepare(Node("=LET(x,A1,x*2)", "x*2"), 1, 1, AllOpen, AsWritten).NotEvaluated);
        Assert.Equal(EvaluatePrecedents.UnknownFunction,
            EvaluatePrecedents.Prepare(Node("=MyUdf(A1)+1", "MyUdf(A1)"), 1, 1, AllOpen, AsWritten).NotEvaluated);
    }

    [Fact]
    public void Prepare_checks_its_arguments()
    {
        var node = Node("=SUM(A1)", "SUM(A1)");

        Assert.Throws<ArgumentNullException>(() => EvaluatePrecedents.Prepare(null!, 1, 1, AllOpen, AsWritten));
        Assert.Throws<ArgumentNullException>(() => EvaluatePrecedents.Prepare(node, 1, 1, null!, AsWritten));
        Assert.Throws<ArgumentNullException>(() => EvaluatePrecedents.Prepare(node, 1, 1, AllOpen, null!));
        Assert.Throws<ArgumentNullException>(() => EvaluatePrecedents.NotEvaluated(null!));
        Assert.Throws<ArgumentNullException>(() => EvaluatePrecedents.ClosedWorkbook(null!));
        Assert.Throws<ArgumentNullException>(() => EvaluatePrecedents.Prepare(node, 1, 1, AllOpen, AsWritten).Finish(null!));
    }
}
