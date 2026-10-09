using System;
using System.Collections.Generic;

namespace Modelwright.Core.Trace;

/// <summary>
/// The rows the Trace In tree shows under a formula in "Evaluate functions &amp; groups" mode (docs/PLAN.md section
/// 4.5, research/07): the formula's own structure. Under a cell, its <see cref="ParsedFormula.TopLevelNodes"/>; under a
/// function or group row, that node's <see cref="FormulaNode.TraceChildren"/>. A parenthesized group or an operator
/// expression is a <see cref="PrecedentKind.Group"/> row labelled with its text (<c>(B2+C2/D2)</c>, <c>E2&gt;0</c>), a
/// function call a <see cref="PrecedentKind.Function"/> row labelled <c>NAME(...)</c> whose children are its
/// arguments, a reference the same row classic mode shows, and a constant argument a row that cannot be expanded. A
/// function's arguments carry Excel's parameter names in the Argument column (<see cref="FormulaNode.ParameterName"/>).
/// LET and LAMBDA local names are left out, as in <see cref="ClassicPrecedents"/>.
/// </summary>
/// <remarks>
/// Pure: the caller supplies each group's and function's value (the add-in evaluates its text with
/// <c>Worksheet.Evaluate</c>) and each reference's item (the add-in resolves it as classic mode does). Both are asked
/// for only the rows being built, so values are read one expansion at a time.
/// </remarks>
public static class EvaluatePrecedents
{
    /// <summary>The Value column of a subexpression longer than <c>Worksheet.Evaluate</c> accepts (255 characters).</summary>
    public const string TooLong = "(too long to evaluate)";

    /// <summary>The Value column of a subexpression Excel could not evaluate.</summary>
    public const string CannotEvaluate = "(cannot evaluate)";

    /// <summary>The label of an omitted argument, as in <c>IF(A1,,0)</c>.</summary>
    public const string OmittedLabel = "(omitted)";

    /// <summary>The rows under a formula cell: its top-level nodes. Empty if the formula could not be parsed.</summary>
    /// <param name="formula">The cell's parsed formula.</param>
    /// <param name="valueOf">The Value column of a group or function node (its evaluated text).</param>
    /// <param name="referenceItem">
    /// The item for a reference node, with the Argument column given (null when the reference is not an argument).
    /// </param>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    public static IReadOnlyList<EvaluatePrecedent> Of(ParsedFormula formula, Func<FormulaNode, string?> valueOf,
        Func<FormulaNode, string?, PrecedentItem> referenceItem)
    {
        if (formula is null)
        {
            throw new ArgumentNullException(nameof(formula));
        }

        return Rows(formula.TopLevelNodes, valueOf, referenceItem);
    }

    /// <summary>The rows under a function or group row: <paramref name="node"/>'s trace children.</summary>
    /// <param name="node">The function or group node the row stands for.</param>
    /// <param name="valueOf">The Value column of a group or function node (its evaluated text).</param>
    /// <param name="referenceItem">
    /// The item for a reference node, with the Argument column given (null when the reference is not an argument).
    /// </param>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    public static IReadOnlyList<EvaluatePrecedent> Of(FormulaNode node, Func<FormulaNode, string?> valueOf,
        Func<FormulaNode, string?, PrecedentItem> referenceItem)
    {
        if (node is null)
        {
            throw new ArgumentNullException(nameof(node));
        }

        return Rows(node.TraceChildren, valueOf, referenceItem);
    }

    /// <summary>
    /// The Precedents column of a group, function or constant row: a function as <c>NAME(...)</c> (Macabacus leaves
    /// the arguments out; a call without arguments is <c>NAME()</c>), an omitted argument as
    /// <see cref="OmittedLabel"/>, anything else as written, on one line.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="node"/> is null.</exception>
    public static string Label(FormulaNode node)
    {
        if (node is null)
        {
            throw new ArgumentNullException(nameof(node));
        }

        switch (node.Kind)
        {
            case FormulaNodeKind.Function:
                return (node.FunctionName ?? string.Empty) + (node.Children.Count == 0 ? "()" : "(...)");
            case FormulaNodeKind.MissingArgument:
                return OmittedLabel;
            default:
                return OneLine(node.Text);
        }
    }

    private static IReadOnlyList<EvaluatePrecedent> Rows(IReadOnlyList<FormulaNode> nodes, Func<FormulaNode, string?> valueOf,
        Func<FormulaNode, string?, PrecedentItem> referenceItem)
    {
        if (valueOf is null)
        {
            throw new ArgumentNullException(nameof(valueOf));
        }

        if (referenceItem is null)
        {
            throw new ArgumentNullException(nameof(referenceItem));
        }

        var rows = new List<EvaluatePrecedent>(nodes.Count);
        foreach (var node in nodes)
        {
            if (!Shows(node))
            {
                continue;
            }

            var argument = node.ParameterName;
            PrecedentItem item;
            switch (node.Kind)
            {
                case FormulaNodeKind.Reference:
                    item = referenceItem(node, argument);
                    break;
                case FormulaNodeKind.Function:
                    item = new PrecedentItem(PrecedentKind.Function, Label(node), valueText: valueOf(node),
                        canExpand: AnyShown(node.Children), argument: argument);
                    break;
                case FormulaNodeKind.Constant:
                case FormulaNodeKind.ArrayConstant:
                case FormulaNodeKind.MissingArgument:
                    item = new PrecedentItem(PrecedentKind.Group, Label(node), valueText: ConstantValue(node), canExpand: false,
                        argument: argument);
                    break;
                default:
                    // A group, an operator expression (an argument such as E2>0), or another construct.
                    item = new PrecedentItem(PrecedentKind.Group, Label(node), valueText: valueOf(node),
                        canExpand: AnyShown(node.TraceChildren), argument: argument);
                    break;
            }

            rows.Add(new EvaluatePrecedent(node, item));
        }

        return rows;
    }

    // A local name (LET, LAMBDA) is not a precedent.
    private static bool Shows(FormulaNode node) =>
        node.Kind != FormulaNodeKind.Reference || node.Reference?.Kind != FormulaReferenceKind.LocalName;

    private static bool AnyShown(IReadOnlyList<FormulaNode> nodes)
    {
        foreach (var node in nodes)
        {
            if (Shows(node))
            {
                return true;
            }
        }

        return false;
    }

    // A constant's value is itself: a string without its quotes, anything else as written. No Excel call is needed.
    private static string ConstantValue(FormulaNode node)
    {
        if (node.Kind == FormulaNodeKind.MissingArgument)
        {
            return string.Empty;
        }

        var text = node.Text;
        return node.Kind == FormulaNodeKind.Constant && text.Length >= 2 && text[0] == '"' && text[text.Length - 1] == '"'
            ? OneLine(text.Substring(1, text.Length - 2).Replace("\"\"", "\""))
            : OneLine(text);
    }

    // Line breaks (Alt+Enter in the formula) as spaces: a row is one line.
    private static string OneLine(string text) =>
        text.IndexOf('\n') < 0 && text.IndexOf('\r') < 0 ? text : text.Replace("\r\n", " ").Replace('\n', ' ').Replace('\r', ' ');
}

/// <summary>A row of <see cref="EvaluatePrecedents"/>: the formula node and the item shown for it.</summary>
public sealed class EvaluatePrecedent
{
    internal EvaluatePrecedent(FormulaNode node, PrecedentItem item)
    {
        Node = node;
        Item = item;
    }

    /// <summary>The part of the formula the row stands for; a function or group row's children come from it.</summary>
    public FormulaNode Node { get; }

    /// <summary>What the row shows.</summary>
    public PrecedentItem Item { get; }

    /// <summary>The item's label.</summary>
    public override string ToString() => Item.Label;
}
