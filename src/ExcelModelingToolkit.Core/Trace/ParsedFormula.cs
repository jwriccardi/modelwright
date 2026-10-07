using System.Collections.Generic;

namespace ExcelModelingToolkit.Core.Trace;

/// <summary>The result of <see cref="FormulaParser.Parse"/>: a formula's references and structure.</summary>
public sealed class ParsedFormula
{
    private static readonly FormulaReference[] NoReferences = new FormulaReference[0];
    private static readonly DynamicReference[] NoDynamicReferences = new DynamicReference[0];
    private static readonly FormulaNode[] NoNodes = new FormulaNode[0];
    private IReadOnlyList<FormulaNode>? _topLevelNodes;

    private ParsedFormula(string formula, FormulaContext context, string? error, FormulaNode? structure,
        IReadOnlyList<FormulaReference> references, IReadOnlyList<DynamicReference> dynamicReferences)
    {
        Formula = formula;
        Context = context;
        Error = error;
        Structure = structure;
        References = references;
        DynamicReferences = dynamicReferences;
    }

    /// <summary>The formula text that was parsed.</summary>
    public string Formula { get; }

    /// <summary>The formula's workbook and sheet.</summary>
    public FormulaContext Context { get; }

    /// <summary>True if the formula was parsed; false if <see cref="Error"/> says why not.</summary>
    public bool IsParsed => Error is null;

    /// <summary>Why the formula could not be parsed, or null.</summary>
    public string? Error { get; }

    /// <summary>
    /// The formula's expression tree (without the leading <c>=</c>), or null if it could not be parsed. Its
    /// reference nodes hold the same objects as <see cref="References"/>.
    /// </summary>
    public FormulaNode? Structure { get; }

    /// <summary>
    /// The nodes a trace tree shows directly under the audited cell in "Evaluate functions &amp; groups" mode: the
    /// root itself when it is a group, function call or reference (<c>=SUM(A1,B1)</c> gives the SUM node,
    /// <c>=(A1+B1)</c> the group, <c>=A1</c> the reference), else its operands as
    /// <see cref="FormulaNode.TraceChildren"/> finds them (<c>=A1+B1</c> gives A1 and B1). Empty for a constant or
    /// if the formula could not be parsed.
    /// </summary>
    public IReadOnlyList<FormulaNode> TopLevelNodes
    {
        get
        {
            if (_topLevelNodes is null)
            {
                var nodes = new List<FormulaNode>();
                if (Structure is not null)
                {
                    FormulaNode.AddOperands(Structure, nodes);
                }

                _topLevelNodes = nodes.Count == 0 ? NoNodes : nodes;
            }

            return _topLevelNodes;
        }
    }

    /// <summary>
    /// Every reference in the order written, duplicates included (Macabacus lists each reference as written).
    /// Text inside string constants and array constants is never a reference. Empty if the formula could not be
    /// parsed.
    /// </summary>
    public IReadOnlyList<FormulaReference> References { get; }

    /// <summary>The calls whose result may be a reference (INDEX, OFFSET, INDIRECT, CHOOSE), in the order written.</summary>
    public IReadOnlyList<DynamicReference> DynamicReferences { get; }

    /// <summary>True if the formula has a call in <see cref="DynamicReferences"/>.</summary>
    public bool HasDynamicReferences => DynamicReferences.Count > 0;

    internal static ParsedFormula Parsed(string formula, FormulaContext context, FormulaNode structure,
        IReadOnlyList<FormulaReference> references, IReadOnlyList<DynamicReference> dynamicReferences) =>
        new ParsedFormula(formula, context, null, structure, references, dynamicReferences);

    internal static ParsedFormula Failed(string formula, FormulaContext context, string error) =>
        new ParsedFormula(formula, context, error, null, NoReferences, NoDynamicReferences);
}
