using System;
using System.Collections.Generic;
using System.Text;

namespace Modelwright.Core.Trace;

/// <summary>
/// The rows the Trace In tree shows under a formula in "Evaluate functions &amp; groups" mode (docs/PLAN.md section
/// 4.5, research/07): the formula's own structure. Under a cell, its <see cref="ParsedFormula.TopLevelNodes"/>; under a
/// function or group row, that node's <see cref="FormulaNode.TraceChildren"/>. A parenthesized group or an operator
/// expression is a <see cref="PrecedentKind.Group"/> row labelled with its text (<c>(B2+C2/D2)</c>, <c>E2&gt;0</c>), a
/// function call a <see cref="PrecedentKind.Function"/> row labelled <c>NAME(...)</c> whose children are its
/// arguments (a LAMBDA called in place, <c>LAMBDA(x,x*2)(A1)</c>, is one <c>LAMBDA(...)(...)</c> row), a reference the
/// same row classic mode shows, and a constant argument a row that cannot be expanded. A function's arguments carry
/// Excel's parameter names in the Argument column (<see cref="FormulaNode.ParameterName"/>). LET and LAMBDA local
/// names are left out, as in <see cref="ClassicPrecedents"/>.
/// </summary>
/// <remarks>
/// <para>
/// Pure: the caller supplies each group's and function's value (the add-in evaluates the text <see cref="Prepare"/>
/// gives it with <c>Worksheet.Evaluate</c>) and each reference's item (the add-in resolves it as classic mode does).
/// Both are asked for only the rows being built, so values are read one expansion at a time.
/// </para>
/// <para>
/// A fragment is not given to the caller to evaluate when that cannot give its value in the formula
/// (<see cref="NotEvaluated"/>): it uses a LET or LAMBDA name declared outside it, or is a LAMBDA that is not called;
/// it calls a function that is not Excel's (a VBA or add-in function may have side effects) or one that reaches
/// outside the workbook (WEBSERVICE, RTD, the CUBE functions...). Its Value column then says why. A fragment that is
/// the whole formula shows the cell's own value instead of being evaluated.
/// </para>
/// </remarks>
public static class EvaluatePrecedents
{
    /// <summary>The longest text <c>Worksheet.Evaluate</c> accepts.</summary>
    public const int MaxEvaluateLength = 255;

    /// <summary>The Value column of a fragment longer than <c>Worksheet.Evaluate</c> accepts (<see cref="MaxEvaluateLength"/>).</summary>
    public const string TooLong = "(not evaluated: longer than Excel's 255-character limit)";

    /// <summary>The Value column of a fragment Excel could not evaluate.</summary>
    public const string CannotEvaluate = "(cannot evaluate)";

    /// <summary>
    /// The Value column of a fragment that uses a LET or LAMBDA name declared outside it (<c>x*2</c> in
    /// <c>LET(x,A1,x*2)</c>), or is a LAMBDA that is not called: evaluated alone it would be <c>#NAME?</c> or
    /// <c>#CALC!</c>, or a defined name's value.
    /// </summary>
    public const string UsesLocalNames = "(uses LET/LAMBDA names: not evaluated)";

    /// <summary>The Value column of a fragment calling a function that is not Excel's (VBA, an add-in, a named LAMBDA).</summary>
    public const string UnknownFunction = "(not evaluated: unknown or user-defined function)";

    /// <summary>
    /// The Value column of a fragment whose meaning depends on the formula's cell and could not be written without it:
    /// the implicit-intersection <c>@</c>, a table reference to the current row that could not be resolved, a defined
    /// name with relative references.
    /// </summary>
    public const string DependsOnItsCell = "(not evaluated: depends on its cell)";

    /// <summary>Appended to the value of a fragment calling NOW, TODAY or a RAND function: evaluating it again gives another value.</summary>
    public const string VolatileNote = " (volatile)";

    /// <summary>The label of an omitted argument, as in <c>IF(A1,,0)</c>.</summary>
    public const string OmittedLabel = "(omitted)";

    /// <summary>The label of a LAMBDA called in place, <c>LAMBDA(x,x*2)(A1)</c>.</summary>
    public const string LambdaCallLabel = "LAMBDA(...)(...)";

    private const string LocalNamePrefix = "_xlpm.";

    /// <summary>
    /// The Value column of a fragment that refers to a closed workbook (<c>Worksheet.Evaluate</c> would give
    /// <c>#REF!</c>). A reference row that opens the workbook does not change rows already built, and collapsing and
    /// expanding a row keeps its children: tracing again does.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="workbookName"/> is null.</exception>
    public static string ClosedWorkbook(string workbookName) =>
        "(not evaluated: [" + (workbookName ?? throw new ArgumentNullException(nameof(workbookName))) +
        "] is closed; trace again once it is open)";

    /// <summary>The Value column of a fragment calling a function that reaches outside the workbook (<see cref="ExcelFunctions.ReachesOutside"/>).</summary>
    /// <exception cref="ArgumentNullException"><paramref name="functionName"/> is null.</exception>
    public static string ReachesOutside(string functionName) =>
        "(not evaluated: " + FunctionSignatures.Normalize(functionName) + " reads external data or runs code)";

    /// <summary>The rows under a formula cell: its top-level nodes. Empty if the formula could not be parsed.</summary>
    /// <param name="formula">The cell's parsed formula.</param>
    /// <param name="valueOf">The Value column of a group or function node (its evaluated text).</param>
    /// <param name="referenceItem">
    /// The item for a reference node, with the Argument column given (null when the reference is not an argument).
    /// </param>
    /// <param name="cellValue">
    /// The formula cell's value as displayed, for a row that is the whole formula (<c>=IF(...)</c>, <c>=(A1+B1)</c>):
    /// that row's value is the cell's, so nothing is evaluated. Only a whole-formula call that can return a reference
    /// (INDEX, OFFSET, INDIRECT, CHOOSE) is still evaluated, to find the range its row goes to, unless it is too long.
    /// Null to treat the whole formula's row like any other.
    /// </param>
    /// <exception cref="ArgumentNullException">An argument other than <paramref name="cellValue"/> is null.</exception>
    public static IReadOnlyList<EvaluatePrecedent> Of(ParsedFormula formula, Func<FormulaNode, string?> valueOf,
        Func<FormulaNode, string?, PrecedentItem> referenceItem, Func<string>? cellValue = null)
    {
        if (formula is null)
        {
            throw new ArgumentNullException(nameof(formula));
        }

        return Rows(formula.TopLevelNodes, valueOf, referenceItem, cellValue is null ? null : formula.Structure, cellValue);
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

        return Rows(node.TraceChildren, valueOf, referenceItem, null, null);
    }

    /// <summary>
    /// The Precedents column of a group, function or constant row: a function as <c>NAME(...)</c> (Macabacus leaves
    /// the arguments out; a call without arguments is <c>NAME()</c>), a LAMBDA called in place as
    /// <see cref="LambdaCallLabel"/>, an omitted argument as <see cref="OmittedLabel"/>, anything else as written, on
    /// one line.
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
            case FormulaNodeKind.Operator when node.IsLambdaCall:
                return LambdaCallLabel;
            default:
                return OneLine(node.Text);
        }
    }

    /// <summary>
    /// Why a group or function row is not evaluated whatever the workbook holds, or null if it can be: it uses a LET or
    /// LAMBDA name declared outside it, or is a LAMBDA that is not called (<see cref="UsesLocalNames"/>); it calls a
    /// function that reaches outside the workbook (<see cref="ReachesOutside(string)"/>) or one that is not Excel's
    /// (<see cref="UnknownFunction"/>). A name declared inside the fragment (<c>LET(x,A1,x*2)</c>,
    /// <c>LAMBDA(x,x*2)(A1)</c>, <c>MAP(A1:A3,LAMBDA(x,x*2))</c>) does not count.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="node"/> is null.</exception>
    public static string? NotEvaluated(FormulaNode node)
    {
        if (node is null)
        {
            throw new ArgumentNullException(nameof(node));
        }

        if (node.Kind == FormulaNodeKind.Function && node.FunctionName == "LAMBDA")
        {
            return UsesLocalNames; // a LAMBDA that is not called is a function value: #CALC!
        }

        // The names the LET and LAMBDA calls inside the fragment declare (a name used in the fragment and declared
        // anywhere in it is taken as bound), and the declarations themselves.
        var nodes = Subtree(node);
        var declared = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var declarations = new HashSet<FormulaNode>();
        foreach (var current in nodes)
        {
            if (current.Kind != FormulaNodeKind.Function || (current.FunctionName != "LET" && current.FunctionName != "LAMBDA"))
            {
                continue;
            }

            var count = current.Children.Count;
            for (var index = 0; index < count - 1; index++)
            {
                if ((current.FunctionName == "LAMBDA" || index % 2 == 0) &&
                    current.Children[index].Reference is { Kind: FormulaReferenceKind.LocalName } declaration)
                {
                    declared.Add(WithoutLocalPrefix(declaration.Name ?? declaration.Text));
                    declarations.Add(current.Children[index]);
                }
            }
        }

        string? function = null;
        foreach (var current in nodes)
        {
            if (current.Kind == FormulaNodeKind.Reference && current.Reference is { Kind: FormulaReferenceKind.LocalName } local &&
                !declarations.Contains(current) && !declared.Contains(WithoutLocalPrefix(local.Name ?? local.Text)))
            {
                return UsesLocalNames;
            }

            if (current.Kind != FormulaNodeKind.Function || current.FunctionName is not string name)
            {
                continue;
            }

            if (name.StartsWith(LocalNamePrefix, StringComparison.OrdinalIgnoreCase))
            {
                // A LAMBDA held by a LET name, called: _xlpm.f(A1).
                if (!declared.Contains(WithoutLocalPrefix(name)))
                {
                    return UsesLocalNames;
                }
            }
            else if (ExcelFunctions.ReachesOutside(name))
            {
                function ??= ReachesOutside(name);
            }
            else if (!ExcelFunctions.IsBuiltIn(name))
            {
                function ??= UnknownFunction;
            }
        }

        return function;
    }

    /// <summary>
    /// What to evaluate for a group or function row, or why it is not evaluated: <see cref="NotEvaluated"/>'s
    /// reasons; then a workbook it refers to that is closed (<see cref="ClosedWorkbook"/>); then what depends on the
    /// formula's cell (<see cref="DependsOnItsCell"/>): the implicit-intersection <c>@</c> (or <c>SINGLE</c>) is not
    /// evaluated, while a table reference to the current row or to the formula's own table (<c>[@Qty]</c>,
    /// <c>Sales[@Qty]</c>, <c>[Qty]</c>) and a defined name are replaced by what <paramref name="outsideItsCell"/>
    /// gives; argument-less <c>ROW()</c> and <c>COLUMN()</c> become the cell's (<see cref="CallingCell"/>); last, the
    /// result must fit <see cref="MaxEvaluateLength"/> (<see cref="TooLong"/>).
    /// </summary>
    /// <param name="node">The group or function node.</param>
    /// <param name="row">The formula cell's row.</param>
    /// <param name="column">The formula cell's column.</param>
    /// <param name="isWorkbookOpen">True if the workbook of that file name is open.</param>
    /// <param name="outsideItsCell">
    /// For a table reference to the current row or the formula's own table, the text of the cells it means for the
    /// formula's cell (<c>'Data'!$E$3</c>); for a defined name, its own text, or null if its formula has relative
    /// references (Excel would evaluate them from the active cell). Null when there is no such text: the row is then
    /// <see cref="DependsOnItsCell"/>.
    /// </param>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    public static EvaluateRequest Prepare(FormulaNode node, int row, int column, Func<string, bool> isWorkbookOpen,
        Func<FormulaReference, string?> outsideItsCell)
    {
        if (node is null)
        {
            throw new ArgumentNullException(nameof(node));
        }

        if (isWorkbookOpen is null)
        {
            throw new ArgumentNullException(nameof(isWorkbookOpen));
        }

        if (outsideItsCell is null)
        {
            throw new ArgumentNullException(nameof(outsideItsCell));
        }

        if (NotEvaluated(node) is string reason)
        {
            return new EvaluateRequest(null, reason, false);
        }

        var nodes = Subtree(node);
        var checkedBooks = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var current in nodes)
        {
            if (current.Reference?.WorkbookName is string book && checkedBooks.Add(book) && !isWorkbookOpen(book))
            {
                return new EvaluateRequest(null, ClosedWorkbook(book), false);
            }
        }

        // The text with each reference that needs the formula's cell replaced, in source order (references do not
        // overlap), and the _xlpm. prefix Range.Formula2 writes before LET and LAMBDA names dropped (they are
        // declared in the fragment: see NotEvaluated), as they are typed.
        var source = node.Text;
        var text = new StringBuilder(source.Length + 16);
        var position = node.Start;
        var isVolatile = false;
        foreach (var current in nodes)
        {
            var isLocal = current.Reference?.Kind == FormulaReferenceKind.LocalName ||
                (current.Kind == FormulaNodeKind.Function &&
                 current.FunctionName?.StartsWith(LocalNamePrefix, StringComparison.OrdinalIgnoreCase) == true);
            if (isLocal && current.Start >= position &&
                string.Compare(source, current.Start - node.Start, LocalNamePrefix, 0, LocalNamePrefix.Length, StringComparison.OrdinalIgnoreCase) == 0)
            {
                text.Append(source, position - node.Start, current.Start - position);
                position = current.Start + LocalNamePrefix.Length;
                continue;
            }

            if ((current.Kind == FormulaNodeKind.Operator && current.Operator == "@") ||
                (current.Kind == FormulaNodeKind.Function && current.FunctionName == "SINGLE"))
            {
                return new EvaluateRequest(null, DependsOnItsCell, false);
            }

            if (current.Kind == FormulaNodeKind.Function && current.FunctionName is string name && ExcelFunctions.IsVolatile(name))
            {
                isVolatile = true;
            }

            if (current.Reference is not FormulaReference reference || !AskOutsideItsCell(reference))
            {
                continue;
            }

            var replacement = outsideItsCell(reference);
            if (replacement is null)
            {
                return new EvaluateRequest(null, DependsOnItsCell, false);
            }

            text.Append(source, position - node.Start, reference.Start - position).Append(replacement);
            position = reference.Start + reference.Length;
        }

        text.Append(source, position - node.Start, node.Start + node.Length - position);
        var evaluated = CallingCell.SubstituteRowAndColumn(text.ToString(), row, column);
        return evaluated.Length > MaxEvaluateLength
            ? new EvaluateRequest(null, TooLong, false)
            : new EvaluateRequest(evaluated, null, isVolatile);
    }

    // A table reference that needs the formula's cell (to find its own table, or its row), or a defined name (whose
    // formula may be relative).
    private static bool AskOutsideItsCell(FormulaReference reference)
    {
        switch (reference.Kind)
        {
            case FormulaReferenceKind.Name:
                return true;
            case FormulaReferenceKind.StructuredReference:
                if (reference.Name is null)
                {
                    return true;
                }

                foreach (var specifier in reference.TableSpecifiers)
                {
                    var trimmed = specifier.Trim();
                    if (trimmed == "@" || string.Equals(trimmed, "#This Row", StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }

                return false;
            default:
                return false;
        }
    }

    private static IReadOnlyList<EvaluatePrecedent> Rows(IReadOnlyList<FormulaNode> nodes, Func<FormulaNode, string?> valueOf,
        Func<FormulaNode, string?, PrecedentItem> referenceItem, FormulaNode? whole, Func<string>? cellValue)
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
                    item = new PrecedentItem(PrecedentKind.Function, Label(node),
                        valueText: Value(node, valueOf, ReferenceEquals(node, whole) ? cellValue : null),
                        canExpand: AnyShown(node.Children), argument: argument);
                    break;
                case FormulaNodeKind.Constant:
                case FormulaNodeKind.ArrayConstant:
                case FormulaNodeKind.MissingArgument:
                    item = new PrecedentItem(PrecedentKind.Group, Label(node), valueText: ConstantValue(node), canExpand: false,
                        argument: argument);
                    break;
                default:
                    // A group, an operator expression (an argument such as E2>0), a LAMBDA called in place, or another
                    // construct.
                    item = new PrecedentItem(node.IsLambdaCall ? PrecedentKind.Function : PrecedentKind.Group, Label(node),
                        valueText: Value(node, valueOf, ReferenceEquals(node, whole) ? cellValue : null),
                        canExpand: AnyShown(node.TraceChildren), argument: argument);
                    break;
            }

            rows.Add(new EvaluatePrecedent(node, item));
        }

        return rows;
    }

    // A group's or function's Value column. The whole formula's (cellValue given) is the cell's value, except that a
    // call that can return a reference is evaluated to find its range (if it fits). Anything else is evaluated unless
    // NotEvaluated says why not.
    private static string? Value(FormulaNode node, Func<FormulaNode, string?> valueOf, Func<string>? cellValue)
    {
        var reason = NotEvaluated(node);
        if (cellValue is null)
        {
            return reason ?? valueOf(node);
        }

        if (reason is null && node.ReturnsReference)
        {
            var value = valueOf(node);
            return value == TooLong ? cellValue() : value;
        }

        return cellValue();
    }

    // The node and every node below it, in source order (parents before children). A loop: a long operator chain is
    // thousands of levels deep.
    private static List<FormulaNode> Subtree(FormulaNode node)
    {
        var result = new List<FormulaNode>();
        var pending = new Stack<FormulaNode>();
        pending.Push(node);
        while (pending.Count > 0)
        {
            var current = pending.Pop();
            result.Add(current);
            for (var index = current.Children.Count - 1; index >= 0; index--)
            {
                pending.Push(current.Children[index]);
            }
        }

        return result;
    }

    private static string WithoutLocalPrefix(string name) =>
        name.StartsWith(LocalNamePrefix, StringComparison.OrdinalIgnoreCase) ? name.Substring(LocalNamePrefix.Length) : name;

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

/// <summary>
/// What <see cref="EvaluatePrecedents.Prepare"/> decided for a group or function row: the text to give
/// <c>Worksheet.Evaluate</c>, or the Value column that says why it is not evaluated.
/// </summary>
public sealed class EvaluateRequest
{
    internal EvaluateRequest(string? text, string? notEvaluated, bool isVolatile)
    {
        Text = text;
        NotEvaluated = notEvaluated;
        IsVolatile = isVolatile;
    }

    /// <summary>The text to evaluate in the formula's sheet, or null if the row is not evaluated.</summary>
    public string? Text { get; }

    /// <summary>The Value column when <see cref="Text"/> is null: why the row is not evaluated.</summary>
    public string? NotEvaluated { get; }

    /// <summary>True if the text calls NOW, TODAY or a RAND function (<see cref="ExcelFunctions.IsVolatile"/>).</summary>
    public bool IsVolatile { get; }

    /// <summary>
    /// The Value column for what Excel returned: <paramref name="value"/>, followed by
    /// <see cref="EvaluatePrecedents.VolatileNote"/> if the text is volatile.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="value"/> is null.</exception>
    public string Finish(string value) =>
        (value ?? throw new ArgumentNullException(nameof(value))) + (IsVolatile ? EvaluatePrecedents.VolatileNote : string.Empty);
}
