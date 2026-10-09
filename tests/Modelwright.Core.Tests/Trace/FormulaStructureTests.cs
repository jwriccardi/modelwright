using System.Linq;
using Modelwright.Core.Trace;
using Xunit;

namespace Modelwright.Core.Tests.Trace;

public class FormulaStructureTests
{
    private static readonly FormulaContext Context = new FormulaContext("Model.xlsx", "Calc");

    // The formula from Macabacus's Trace In help page (research/07).
    private const string MacabacusExample =
        "=(B2+C2/D2)+IF(E2>0,F2+G2,SUM(H2:J2))*(K2+L2-ABS(M2))+PRODUCT(N2:T2,U2)+V2";

    private static FormulaNode Structure(string formula)
    {
        var parsed = FormulaParser.Parse(formula, Context);
        Assert.True(parsed.IsParsed, parsed.Error);
        return parsed.Structure!;
    }

    [Fact]
    public void Macabacus_example_has_the_documented_top_level_nodes()
    {
        var top = FormulaParser.Parse(MacabacusExample, Context).TopLevelNodes;

        Assert.Equal(new[] { "(B2+C2/D2)", "IF(E2>0,F2+G2,SUM(H2:J2))", "(K2+L2-ABS(M2))", "PRODUCT(N2:T2,U2)", "V2" },
            top.Select(node => node.Text));
        Assert.Equal(
            new[] { FormulaNodeKind.Group, FormulaNodeKind.Function, FormulaNodeKind.Group, FormulaNodeKind.Function, FormulaNodeKind.Reference },
            top.Select(node => node.Kind));
    }

    [Theory]
    [InlineData("=A1", "A1", FormulaNodeKind.Reference)]
    [InlineData("=Rate", "Rate", FormulaNodeKind.Reference)]
    [InlineData("=A1#", "A1#", FormulaNodeKind.Reference)]
    [InlineData("=SUM(A1,B1)", "SUM(A1,B1)", FormulaNodeKind.Function)]
    [InlineData("=(A1+B1)", "(A1+B1)", FormulaNodeKind.Group)]
    [InlineData("=SUM(A1)+0", "SUM(A1)", FormulaNodeKind.Function)]
    [InlineData("=-A1", "A1", FormulaNodeKind.Reference)]
    public void Top_level_node_is_the_root_itself_unless_it_is_an_operator(string formula, string text, FormulaNodeKind kind)
    {
        var top = FormulaParser.Parse(formula, Context).TopLevelNodes;

        Assert.Equal(text, top.Single().Text);
        Assert.Equal(kind, top.Single().Kind);
    }

    [Fact]
    public void Top_level_nodes_of_an_operator_root_are_its_operands()
    {
        Assert.Equal(new[] { "A1", "B1" }, FormulaParser.Parse("=A1+B1", Context).TopLevelNodes.Select(node => node.Text));
        Assert.Empty(FormulaParser.Parse("=5", Context).TopLevelNodes);
        Assert.Empty(FormulaParser.Parse("=SUM(A1", Context).TopLevelNodes);
    }

    [Fact]
    public void Function_root_traces_to_its_arguments()
    {
        var sum = FormulaParser.Parse("=SUM(A1,B1)", Context).TopLevelNodes.Single();

        Assert.Equal(new[] { "A1", "B1" }, sum.TraceChildren.Select(node => node.Text));
    }

    [Fact]
    public void Macabacus_example_IF_arguments_carry_tooltip_parameter_names()
    {
        var function = FormulaParser.Parse(MacabacusExample, Context).TopLevelNodes[1];

        Assert.Equal("IF", function.FunctionName);
        Assert.Equal(new[] { "E2>0", "F2+G2", "SUM(H2:J2)" }, function.TraceChildren.Select(node => node.Text));
        Assert.Equal(new[] { "logical_test", "[value_if_true]", "[value_if_false]" }, function.TraceChildren.Select(node => node.ParameterName));
        Assert.Equal(new[] { 0, 1, 2 }, function.TraceChildren.Select(node => node.ArgumentIndex));
        Assert.Equal(new[] { FormulaNodeKind.Operator, FormulaNodeKind.Operator, FormulaNodeKind.Function },
            function.TraceChildren.Select(node => node.Kind));
    }

    [Fact]
    public void Operator_argument_traces_to_its_references_not_its_constants()
    {
        var test = FormulaParser.Parse(MacabacusExample, Context).TopLevelNodes[1].Children[0];

        Assert.Equal(">", test.Operator);
        Assert.Equal(OperatorFixity.Infix, test.Fixity);
        Assert.Equal(new[] { "E2" }, test.TraceChildren.Select(node => node.Text));
    }

    [Fact]
    public void Group_traces_to_the_operands_inside_it()
    {
        var group = FormulaParser.Parse(MacabacusExample, Context).TopLevelNodes[2];

        Assert.Equal(new[] { "K2", "L2", "ABS(M2)" }, group.TraceChildren.Select(node => node.Text));
        Assert.Single(group.Children);
        Assert.Equal("K2+L2-ABS(M2)", group.Children[0].Text);
    }

    [Fact]
    public void Node_spans_point_into_the_formula()
    {
        var root = Structure(MacabacusExample);

        foreach (var node in Flatten(root))
        {
            Assert.Equal(node.Text, MacabacusExample.Substring(node.Start, node.Length));
        }

        Assert.Equal(1, root.Start);
        Assert.Equal(MacabacusExample.Length - 1, root.Length);
    }

    [Fact]
    public void Nested_IF_names_every_level()
    {
        var root = Structure("=IF(A1>100,\"High\",IF(A1>50,\"Mid\",IF(A1>0,\"Low\",\"None\")))");

        var level = root;
        for (var depth = 0; depth < 3; depth++)
        {
            Assert.Equal("IF", level.FunctionName);
            Assert.Equal(3, level.Children.Count);
            Assert.Equal(FormulaNodeKind.Constant, level.Children[1].Kind);
            Assert.Equal("[value_if_true]", level.Children[1].ParameterName);
            level = level.Children[2];
        }

        Assert.Equal("\"None\"", level.Text);
        Assert.Equal("[value_if_false]", level.ParameterName);
    }

    [Fact]
    public void Repeating_arguments_continue_their_numbering()
    {
        var sum = Structure("=SUM(A1,B1,C1,D1)");
        var sumifs = Structure("=SUMIFS(D:D,A:A,\"x\",B:B,\">0\",C:C,F1)");

        Assert.Equal(new[] { "number1", "[number2]", "[number3]", "[number4]" }, sum.Children.Select(node => node.ParameterName));
        Assert.Equal(
            new[] { "sum_range", "criteria_range1", "criteria1", "[criteria_range2]", "[criteria2]", "[criteria_range3]", "[criteria3]" },
            sumifs.Children.Select(node => node.ParameterName));
    }

    [Fact]
    public void Future_function_prefixes_are_removed_from_names()
    {
        var root = Structure("=_xlfn.XLOOKUP(A1,B:B,C:C,,0)");

        Assert.Equal("XLOOKUP", root.FunctionName);
        Assert.Equal(new[] { "lookup_value", "lookup_array", "return_array", "[if_not_found]", "[match_mode]" },
            root.Children.Select(node => node.ParameterName));
        Assert.Equal(FormulaNodeKind.MissingArgument, root.Children[3].Kind);
        Assert.Equal(string.Empty, root.Children[3].Text);
    }

    [Fact]
    public void Unknown_function_arguments_have_no_parameter_names()
    {
        var root = Structure("=MyUdf(A1,2)");

        Assert.Equal("MYUDF", root.FunctionName);
        Assert.All(root.Children, node => Assert.Null(node.ParameterName));
        Assert.Equal(new[] { 0, 1 }, root.Children.Select(node => node.ArgumentIndex));
    }

    [Fact]
    public void External_user_defined_function_keeps_its_workbook_prefix()
    {
        var root = Structure("=Book2.xlsx!MyUdf(A1)");

        Assert.Equal(FormulaNodeKind.Function, root.Kind);
        Assert.Equal("Book2.xlsx!MYUDF", root.FunctionName);
        Assert.Equal("A1", root.Children.Single().Text);
    }

    [Fact]
    public void Zero_argument_function_has_no_children()
    {
        var root = Structure("=TODAY()");

        Assert.Equal(FormulaNodeKind.Function, root.Kind);
        Assert.Empty(root.Children);
    }

    [Fact]
    public void Reference_returning_functions_are_marked()
    {
        var functions = FormulaParser.Parse("=INDEX(A1:C3,2,2)+OFFSET(A1,1,1)+SUM(A1)", Context).TopLevelNodes;

        Assert.Equal(new[] { true, true, false }, functions.Select(node => node.ReturnsReference));
    }

    [Fact]
    public void Range_with_a_function_end_is_a_range_operator()
    {
        var root = Structure("=SUM(A1:INDEX(B:B,5))").Children.Single();

        Assert.Equal(FormulaNodeKind.Operator, root.Kind);
        Assert.Equal(":", root.Operator);
        Assert.Equal(new[] { FormulaNodeKind.Reference, FormulaNodeKind.Function }, root.Children.Select(node => node.Kind));
    }

    [Fact]
    public void Union_and_intersection_are_operators()
    {
        var union = Structure("=SUM((A1:A3,C1:C3))").Children.Single();
        var intersection = Structure("=SUM(A1:A3 A2:C2)").Children.Single();

        Assert.Equal(",", union.Operator);
        Assert.Equal(2, union.Children.Count);
        Assert.Equal(" ", intersection.Operator);
        Assert.Equal(new[] { "A1:A3", "A2:C2" }, intersection.Children.Select(node => node.Text));
    }

    [Fact]
    public void Prefix_and_postfix_operators_have_fixity()
    {
        // Excel applies negation before percent: -A1% is (-A1)%.
        var root = Structure("=-A1%");

        Assert.Equal("%", root.Operator);
        Assert.Equal(OperatorFixity.Postfix, root.Fixity);
        Assert.Equal("-", root.Children[0].Operator);
        Assert.Equal(OperatorFixity.Prefix, root.Children[0].Fixity);
        Assert.Equal("A1", root.Children[0].Children[0].Text);
    }

    [Fact]
    public void Implicit_intersection_is_a_prefix_operator()
    {
        var root = Structure("=@A1:A10");

        Assert.Equal("@", root.Operator);
        Assert.Equal(FormulaReferenceKind.Range, root.Children[0].Reference!.Kind);
    }

    [Fact]
    public void Constants_and_array_constants_are_leaves()
    {
        var root = Structure("=SUM({1,2;3,4})+\"A1\"+TRUE+#N/A");

        var kinds = Flatten(root).Where(node => node.Children.Count == 0).Select(node => node.Kind).ToList();

        Assert.Equal(new[] { FormulaNodeKind.ArrayConstant, FormulaNodeKind.Constant, FormulaNodeKind.Constant, FormulaNodeKind.Constant }, kinds);
        Assert.DoesNotContain(root.TraceChildren, node => node.Kind == FormulaNodeKind.Constant);
    }

    [Fact]
    public void Reference_nodes_hold_the_same_objects_as_the_reference_list()
    {
        var parsed = FormulaParser.Parse(MacabacusExample, Context);

        var fromTree = Flatten(parsed.Structure!).Where(node => node.Kind == FormulaNodeKind.Reference).Select(node => node.Reference!).ToList();

        Assert.Equal(13, fromTree.Count);
        Assert.Equal(parsed.References.ToList(), fromTree);
    }

    [Fact]
    public void Parenthesized_reference_is_a_group()
    {
        var root = Structure("=(A1)*2");

        var group = root.Children[0];
        Assert.Equal(FormulaNodeKind.Group, group.Kind);
        Assert.Equal("(A1)", group.Text);
        Assert.Equal(FormulaNodeKind.Reference, group.Children.Single().Kind);
    }

    [Fact]
    public void Chained_range_operators_make_one_bounding_range()
    {
        var parsed = FormulaParser.Parse("=SUM(A1:B2:C3)", Context);
        var range = parsed.Structure!.Children.Single();

        Assert.Equal(FormulaNodeKind.Reference, range.Kind);
        Assert.Equal("A1:B2:C3", range.Text);
        Assert.Same(parsed.References.Single(), range.Reference);
        Assert.Equal("A1:C3", range.Reference!.Address);
    }

    [Fact]
    public void Range_operator_between_sheets_stays_an_operator()
    {
        var range = Structure("=SUM(Sheet2!A1:B2:C3)").Children.Single();

        Assert.Equal(":", range.Operator);
        Assert.Equal(new[] { "Sheet2!A1:B2", "C3" }, range.Children.Select(node => node.Text));
    }

    [Fact]
    public void Let_and_lambda_arguments_are_named()
    {
        var let = Structure("=LET(x,A1,y,B1,x+y)");
        var lambda = Structure("=MAP(A1:A3,LAMBDA(a,b,a+b))").Children[1];

        Assert.Equal(new[] { "name1", "name_value1", "calculation_or_name2", "[name_value2]", "[calculation_or_name3]" },
            let.Children.Select(node => node.ParameterName));
        Assert.Equal(new[] { "[parameter1]", "[parameter2]", "calculation" }, lambda.Children.Select(node => node.ParameterName));
    }

    private static System.Collections.Generic.IEnumerable<FormulaNode> Flatten(FormulaNode node) =>
        new[] { node }.Concat(node.Children.SelectMany(Flatten));
}
