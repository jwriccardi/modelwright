using System.Collections.Generic;

namespace ExcelModelingToolkit.Core.Trace;

/// <summary>
/// A node of a formula's structure (docs/PLAN.md section 4.5, "Evaluate functions &amp; groups"): a group, function
/// call, reference, operator expression or constant, with its place in the formula text.
/// </summary>
public sealed class FormulaNode
{
    private static readonly FormulaNode[] NoChildren = new FormulaNode[0];

    internal FormulaNode(FormulaNodeKind kind, string text, int start)
    {
        Kind = kind;
        Text = text;
        Start = start;
    }

    /// <summary>The kind of node.</summary>
    public FormulaNodeKind Kind { get; }

    /// <summary>The node's source text: for a function, from its name to its closing parenthesis.</summary>
    public string Text { get; internal set; }

    /// <summary>The 0-based position of <see cref="Text"/> in the formula string (which starts with <c>=</c>).</summary>
    public int Start { get; internal set; }

    /// <summary>The length of <see cref="Text"/>.</summary>
    public int Length => Text.Length;

    /// <summary>
    /// The child nodes: a group's inner expression, a function's arguments, an operator's operands. Empty for
    /// leaves.
    /// </summary>
    public IReadOnlyList<FormulaNode> Children { get; internal set; } = NoChildren;

    /// <summary>
    /// A function's name in upper case, without the <c>_xlfn.</c> / <c>_xlws.</c> prefixes Excel stores for newer
    /// functions (<c>XLOOKUP</c>). An external user-defined function keeps its workbook prefix. Null for other
    /// kinds.
    /// </summary>
    public string? FunctionName { get; internal set; }

    /// <summary>
    /// True for INDEX, OFFSET, INDIRECT and CHOOSE, which can return a reference; the caller can evaluate the call to
    /// find the range it points to.
    /// </summary>
    public bool ReturnsReference { get; internal set; }

    /// <summary>
    /// An operator node's operator as written (<c>+</c>, <c>&lt;&gt;</c>, <c>:</c>, <c>,</c>), or a single space for
    /// intersection. Null for other kinds.
    /// </summary>
    public string? Operator { get; internal set; }

    /// <summary>Where the operator sits; <see cref="OperatorFixity.None"/> for other kinds.</summary>
    public OperatorFixity Fixity { get; internal set; }

    /// <summary>The reference of a <see cref="FormulaNodeKind.Reference"/> node, else null.</summary>
    public FormulaReference? Reference { get; internal set; }

    /// <summary>The 0-based argument position when this node is a function argument, else -1.</summary>
    public int ArgumentIndex { get; internal set; } = -1;

    /// <summary>
    /// When this node is a function argument, the parameter's name in Excel's tooltip style
    /// (<c>logical_test</c>, <c>[value_if_true]</c>) from <see cref="FunctionSignatures"/>; null if the function is
    /// not in the table or the node is not an argument.
    /// </summary>
    public string? ParameterName { get; internal set; }

    /// <summary>
    /// The nodes a trace tree shows under this one, the way Macabacus's "Evaluate functions &amp; groups" view does
    /// (research/07): a function's arguments, each as-is; for any other node, its operands with operator
    /// expressions flattened away, keeping groups, function calls and references but not constants. So for
    /// <c>(B2+C2/D2)+IF(...)*(K2+L2)+V2</c> the root's trace children are the group, IF, the second group and V2.
    /// </summary>
    public IReadOnlyList<FormulaNode> TraceChildren
    {
        get
        {
            if (Kind == FormulaNodeKind.Function)
            {
                return Children;
            }

            var result = new List<FormulaNode>();
            foreach (var child in Children)
            {
                AddOperands(child, result);
            }

            return result;
        }
    }

    /// <summary>The node's source text.</summary>
    public override string ToString() => Text;

    private static void AddOperands(FormulaNode node, List<FormulaNode> result)
    {
        switch (node.Kind)
        {
            case FormulaNodeKind.Operator:
            case FormulaNodeKind.Other:
                foreach (var child in node.Children)
                {
                    AddOperands(child, result);
                }

                break;
            case FormulaNodeKind.Group:
            case FormulaNodeKind.Function:
            case FormulaNodeKind.Reference:
                result.Add(node);
                break;
        }
    }
}
