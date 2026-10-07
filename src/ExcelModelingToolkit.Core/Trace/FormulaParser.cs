using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelModelingToolkit.Core.Undo;
using Irony.Parsing;
using XLParser;

namespace ExcelModelingToolkit.Core.Trace;

/// <summary>
/// Parses an Excel formula, as <c>Range.Formula</c> returns it (en-US, A1 style), into its references
/// (<see cref="ParsedFormula.References"/>) and its structure (<see cref="ParsedFormula.Structure"/>).
/// </summary>
/// <remarks>
/// <para>
/// The grammar is XLParser's (https://github.com/spreadsheetlab/XLParser, MPL-2.0). On top of its parse tree this
/// class adds what Trace In needs: source spans, spill references (<c>A1#</c>), LET and LAMBDA local names (XLParser
/// reports them as defined names), sheet, workbook and path prefixes read from the text (correctly unescaping
/// <c>''</c> in quoted names), workbook-level external names (<c>Book.xlsx!Name</c>), and normalized addresses.
/// </para>
/// <para>
/// Thread-safe: XLParser keeps one parser per thread. The first parse on a thread builds the grammar (about
/// 0.1 s); later parses take well under a millisecond.
/// </para>
/// </remarks>
public static class FormulaParser
{
    private const string LocalNamePrefix = "_xlpm.";

    // Functions whose result can itself be a reference (docs/PLAN.md section 4.5).
    private static readonly HashSet<string> ReferenceFunctions =
        new HashSet<string>(StringComparer.Ordinal) { "INDEX", "OFFSET", "INDIRECT", "CHOOSE" };

    private static readonly string[] WorkbookExtensions =
        { ".xlsx", ".xlsm", ".xlsb", ".xls", ".xltx", ".xltm", ".xlt", ".xlam", ".xla", ".csv" };

    /// <summary>Parses <paramref name="formula"/>, which lives on <paramref name="context"/>'s sheet.</summary>
    /// <param name="formula">The formula text, starting with <c>=</c>.</param>
    /// <param name="context">The workbook and sheet holding the formula.</param>
    /// <returns>
    /// The references and structure; or, if the text is not a formula or the grammar rejects it, a result with
    /// <see cref="ParsedFormula.Error"/> set and no references. It never throws for bad input.
    /// </returns>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    public static ParsedFormula Parse(string formula, FormulaContext context)
    {
        if (formula is null)
        {
            throw new ArgumentNullException(nameof(formula));
        }

        if (context is null)
        {
            throw new ArgumentNullException(nameof(context));
        }

        if (!formula.StartsWith("=", StringComparison.Ordinal) || formula.Trim().Length < 2)
        {
            return ParsedFormula.Failed(formula, context, "Not a formula: it must start with '=' and have an expression.");
        }

        try
        {
            var root = ExcelFormulaParser.Parse(formula);
            var builder = new Builder(formula, context, root);
            var structure = builder.Build(root);
            return ParsedFormula.Parsed(formula, context, structure, builder.References, builder.DynamicReferences);
        }
        catch (Exception ex)
        {
            // ArgumentException is XLParser's documented failure. Anything else from the third-party grammar is
            // reported the same way: the caller falls back (Trace In can still use Range.DirectPrecedents).
            return ParsedFormula.Failed(formula, context, "The formula could not be parsed: " + FirstLine(ex.Message));
        }
    }

    private static string FirstLine(string text)
    {
        var end = text.IndexOfAny(new[] { '\r', '\n' });
        return end < 0 ? text : text.Substring(0, end);
    }

    /// <summary>Walks XLParser's tree once, building the structure and collecting references in source order.</summary>
    private sealed class Builder
    {
        private readonly string _formula;
        private readonly FormulaContext _context;
        private readonly Dictionary<ParseTreeNode, ParserReference> _parserReferences =
            new Dictionary<ParseTreeNode, ParserReference>();

        // LET / LAMBDA names visible at the current point of the walk (innermost last).
        private readonly List<string> _localNames = new List<string>();

        // Enclosing INDEX / OFFSET / INDIRECT / CHOOSE calls (innermost last).
        private readonly List<string> _referenceFunctions = new List<string>();

        // True while building a LET name or LAMBDA parameter declaration.
        private bool _declaring;

        public Builder(string formula, FormulaContext context, ParseTreeNode root)
        {
            _formula = formula;
            _context = context;
            foreach (var reference in root.GetParserReferences())
            {
                if (reference.ReferenceType != ReferenceType.UserDefinedFunction)
                {
                    _parserReferences[reference.ReferenceNode] = reference;
                }
            }
        }

        public List<FormulaReference> References { get; } = new List<FormulaReference>();

        public List<DynamicReference> DynamicReferences { get; } = new List<DynamicReference>();

        public FormulaNode Build(ParseTreeNode node)
        {
            if (_parserReferences.TryGetValue(node, out var parserReference))
            {
                return ReferenceLeaf(node, parserReference);
            }

            var children = node.ChildNodes;
            switch (node.Term.Name)
            {
                case GrammarNames.FormulaWithEq:
                    return Build(children[children.Count - 1]);

                case GrammarNames.Formula:
                case GrammarNames.Reference:
                    if (node.IsParentheses())
                    {
                        return WithChildren(Node(FormulaNodeKind.Group, node), Build(children[0]));
                    }

                    if (children.Count == 1)
                    {
                        return Build(children[0]);
                    }

                    if (children.Count == 2 && children[0].Is(GrammarNames.Prefix) && children[1].Is(GrammarNames.UDFunctionCall))
                    {
                        // An external user-defined function: Book.xlsx!MyUdf(...).
                        return Function(node, children[1], SourceText(children[0]));
                    }

                    break;

                case GrammarNames.FunctionCall:
                case GrammarNames.ReferenceFunctionCall:
                case GrammarNames.UDFunctionCall:
                    var built = Operation(node);
                    if (built is not null)
                    {
                        return built;
                    }

                    break;

                case GrammarNames.Argument:
                    return children.Count == 1 ? Build(children[0]) : Node(FormulaNodeKind.MissingArgument, node);

                case GrammarNames.EmptyArgument:
                    return Node(FormulaNodeKind.MissingArgument, node);

                case GrammarNames.Constant:
                    return Node(FormulaNodeKind.Constant, node);

                case GrammarNames.ConstantArray:
                    return Node(FormulaNodeKind.ArrayConstant, node);

                case GrammarNames.ReservedName:
                    // A built-in name such as _xlnm.Print_Area: a defined name Excel created.
                    var reserved = new FormulaReference(FormulaReferenceKind.Name, SourceText(node), Start(node))
                    {
                        Name = SourceText(node),
                        EnclosingReferenceFunction = CurrentReferenceFunction,
                    };
                    return Leaf(node, reserved);
            }

            var other = Node(FormulaNodeKind.Other, node);
            other.Children = children.Where(child => child.Term is NonTerminal).Select(Build).ToList();
            return other;
        }

        private string? CurrentReferenceFunction =>
            _referenceFunctions.Count == 0 ? null : _referenceFunctions[_referenceFunctions.Count - 1];

        private FormulaNode? Operation(ParseTreeNode node)
        {
            var children = node.ChildNodes;
            if (node.IsNamedFunction())
            {
                return Function(node, node, null);
            }

            if (node.IsUnion())
            {
                var union = Operator(node, ",", OperatorFixity.Infix);
                union.Children = children[0].ChildNodes.Select(Build).ToList();
                return union;
            }

            if (children.Count == 2 && children[1].Term.Name == "#")
            {
                // A spill reference, A1#: the reference with the # included.
                var inner = Build(children[0]);
                if (inner.Kind == FormulaNodeKind.Reference && inner.Reference is FormulaReference spilled)
                {
                    spilled.IsSpill = true;
                    spilled.Text = SourceText(node);
                    inner.Text = spilled.Text;
                    return inner;
                }

                return WithChildren(Operator(node, "#", OperatorFixity.Postfix), inner);
            }

            if (node.IsIntersection())
            {
                return WithChildren(Operator(node, " ", OperatorFixity.Infix), Build(children[0]), Build(children[2]));
            }

            if (node.IsBinaryOperation())
            {
                return WithChildren(Operator(node, SourceText(children[1]), OperatorFixity.Infix), Build(children[0]), Build(children[2]));
            }

            if (node.IsUnaryPrefixOperation())
            {
                return WithChildren(Operator(node, SourceText(children[0]), OperatorFixity.Prefix), Build(children[1]));
            }

            if (node.IsUnaryPostfixOperation())
            {
                return WithChildren(Operator(node, SourceText(children[1]), OperatorFixity.Postfix), Build(children[0]));
            }

            return null;
        }

        private FormulaNode Function(ParseTreeNode spanNode, ParseTreeNode call, string? prefix)
        {
            var name = (prefix ?? string.Empty) + FunctionSignatures.Normalize(SourceText(call.ChildNodes[0]).TrimEnd('('));
            var function = Node(FormulaNodeKind.Function, spanNode);
            function.FunctionName = name;
            function.ReturnsReference = ReferenceFunctions.Contains(name);
            if (function.ReturnsReference)
            {
                DynamicReferences.Add(new DynamicReference(name, function.Text, function.Start));
                _referenceFunctions.Add(name);
            }

            var arguments = call.ChildNodes.Count > 1 && call.ChildNodes[1].Is(GrammarNames.Arguments)
                ? call.ChildNodes[1].ChildNodes
                : new ParseTreeNodeList();
            var localNamesBefore = _localNames.Count;
            var built = new List<FormulaNode>(arguments.Count);
            for (var index = 0; index < arguments.Count; index++)
            {
                var declaration = IsDeclaration(name, index, arguments.Count);
                if (name == "LET" && index % 2 == 0 && index > 0)
                {
                    // LET(name1, value1, name2, value2, ..., calculation): name1 is visible from value2 on.
                    AddLocalName(built[index - 2]);
                }

                if (name == "LAMBDA" && index == arguments.Count - 1)
                {
                    // LAMBDA(param1, param2, ..., calculation): the parameters are visible in the calculation.
                    foreach (var parameter in built)
                    {
                        AddLocalName(parameter);
                    }
                }

                _declaring = declaration;
                var argument = Build(arguments[index]);
                _declaring = false;
                argument.ArgumentIndex = index;
                argument.ParameterName = FunctionSignatures.GetParameterName(name, index, arguments.Count);
                built.Add(argument);
            }

            _localNames.RemoveRange(localNamesBefore, _localNames.Count - localNamesBefore);
            if (function.ReturnsReference)
            {
                _referenceFunctions.RemoveAt(_referenceFunctions.Count - 1);
            }

            function.Children = built;
            return function;
        }

        private static bool IsDeclaration(string function, int index, int count) =>
            index < count - 1 && (function == "LAMBDA" || (function == "LET" && index % 2 == 0));

        private void AddLocalName(FormulaNode declaration)
        {
            if (declaration.Reference is { Kind: FormulaReferenceKind.LocalName } reference && reference.Name is not null)
            {
                _localNames.Add(reference.Name);
            }
        }

        private bool IsLocalName(string name)
        {
            var bare = StripLocalPrefix(name);
            return _localNames.Any(local => string.Equals(StripLocalPrefix(local), bare, StringComparison.OrdinalIgnoreCase));
        }

        private static string StripLocalPrefix(string name) =>
            name.StartsWith(LocalNamePrefix, StringComparison.OrdinalIgnoreCase) ? name.Substring(LocalNamePrefix.Length) : name;

        private FormulaNode ReferenceLeaf(ParseTreeNode node, ParserReference parsed)
        {
            var text = SourceText(node);
            var reference = new FormulaReference(Kind(parsed.ReferenceType), text, Start(node))
            {
                EnclosingReferenceFunction = CurrentReferenceFunction,
            };

            var prefixNode = FindPrefix(node);
            var prefix = prefixNode is null ? default : ParsePrefix(SourceText(prefixNode), reference.Kind);
            reference.WorkbookPath = prefix.Path;
            reference.WorkbookName = prefix.File;
            reference.Sheet = prefix.Sheet;
            reference.LastSheet = prefix.LastSheet;
            if (reference.WorkbookName is not null && reference.WorkbookPath is null &&
                string.Equals(reference.WorkbookName, _context.WorkbookName, StringComparison.OrdinalIgnoreCase))
            {
                reference.WorkbookName = null;
            }

            if (reference.WorkbookName is null && reference.LastSheet is null &&
                string.Equals(reference.Sheet, _context.SheetName, StringComparison.OrdinalIgnoreCase))
            {
                reference.Sheet = null;
            }

            switch (reference.Kind)
            {
                case FormulaReferenceKind.Name:
                    reference.Name = parsed.Name;
                    if (prefixNode is null && (_declaring || IsLocalName(parsed.Name)))
                    {
                        reference.Kind = FormulaReferenceKind.LocalName;
                    }

                    break;
                case FormulaReferenceKind.StructuredReference:
                    reference.Name = string.IsNullOrEmpty(parsed.Name) ? null : parsed.Name;
                    reference.TableColumns = parsed.TableColumns ?? new string[0];
                    reference.TableSpecifiers = parsed.TableSpecifiers ?? new string[0];
                    break;
                case FormulaReferenceKind.RefError:
                    break;
                default:
                    SetAddress(reference, parsed.MinLocation, parsed.MaxLocation);
                    break;
            }

            return Leaf(node, reference);
        }

        private FormulaNode Leaf(ParseTreeNode node, FormulaReference reference)
        {
            References.Add(reference);
            var leaf = Node(FormulaNodeKind.Reference, node);
            leaf.Reference = reference;
            return leaf;
        }

        private static FormulaReferenceKind Kind(ReferenceType type)
        {
            switch (type)
            {
                case ReferenceType.Cell:
                    return FormulaReferenceKind.Cell;
                case ReferenceType.CellRange:
                    return FormulaReferenceKind.Range;
                case ReferenceType.VerticalRange:
                    return FormulaReferenceKind.WholeColumn;
                case ReferenceType.HorizontalRange:
                    return FormulaReferenceKind.WholeRow;
                case ReferenceType.Table:
                    return FormulaReferenceKind.StructuredReference;
                case ReferenceType.RefError:
                    return FormulaReferenceKind.RefError;
                default:
                    return FormulaReferenceKind.Name;
            }
        }

        // Fills Address and Area from XLParser's corners (A1 / $A$1 for cells, A for columns, 1 for rows).
        private static void SetAddress(FormulaReference reference, string min, string max)
        {
            min = min.Replace("$", string.Empty).ToUpperInvariant();
            max = max.Replace("$", string.Empty).ToUpperInvariant();
            int row1 = 1, column1 = 1, row2 = CellRect.MaxRows, column2 = CellRect.MaxColumns;
            bool valid;
            switch (reference.Kind)
            {
                case FormulaReferenceKind.WholeColumn:
                    valid = TryColumn(min, out column1) && TryColumn(max, out column2);
                    break;
                case FormulaReferenceKind.WholeRow:
                    valid = TryRow(min, out row1) && TryRow(max, out row2);
                    break;
                default:
                    valid = TryCell(min, out row1, out column1) && TryCell(max, out row2, out column2);
                    break;
            }

            if (!valid)
            {
                // XLParser accepted something outside the sheet; keep the text, with no cells.
                reference.Address = reference.Kind == FormulaReferenceKind.Cell || min == max ? min : min + ":" + max;
                return;
            }

            var top = Math.Min(row1, row2);
            var left = Math.Min(column1, column2);
            var area = new CellRect(top, left, Math.Abs(row2 - row1) + 1, Math.Abs(column2 - column1) + 1);
            reference.Area = area;
            switch (reference.Kind)
            {
                case FormulaReferenceKind.WholeColumn:
                    reference.Address = CellRect.ColumnName(area.Column) + ":" + CellRect.ColumnName(area.LastColumn);
                    break;
                case FormulaReferenceKind.WholeRow:
                    reference.Address = area.Row.ToString(CultureInfo.InvariantCulture) + ":" +
                        area.LastRow.ToString(CultureInfo.InvariantCulture);
                    break;
                default:
                    var first = CellRect.ColumnName(area.Column) + area.Row.ToString(CultureInfo.InvariantCulture);
                    reference.Address = reference.Kind == FormulaReferenceKind.Cell
                        ? first
                        : first + ":" + CellRect.ColumnName(area.LastColumn) + area.LastRow.ToString(CultureInfo.InvariantCulture);
                    break;
            }
        }

        private static bool TryCell(string text, out int row, out int column)
        {
            var split = 0;
            while (split < text.Length && text[split] >= 'A' && text[split] <= 'Z')
            {
                split++;
            }

            row = 0;
            return TryColumn(text.Substring(0, split), out column) && TryRow(text.Substring(split), out row);
        }

        private static bool TryColumn(string letters, out int column)
        {
            column = 0;
            if (letters.Length == 0 || letters.Length > 3)
            {
                return false;
            }

            foreach (var letter in letters)
            {
                if (letter < 'A' || letter > 'Z')
                {
                    return false;
                }

                column = (column * 26) + (letter - 'A' + 1);
            }

            return column <= CellRect.MaxColumns;
        }

        private static bool TryRow(string digits, out int row) =>
            int.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out row) && row >= 1 && row <= CellRect.MaxRows;

        // The first Prefix node (Sheet1!, 'My Sheet'!, [Book.xlsx]Sheet1!, 'C:\dir\[Book.xlsx]Sheet1'!) in a reference.
        private static ParseTreeNode? FindPrefix(ParseTreeNode node)
        {
            if (node.Is(GrammarNames.Prefix))
            {
                return node;
            }

            foreach (var child in node.ChildNodes)
            {
                var found = FindPrefix(child);
                if (found is not null)
                {
                    return found;
                }
            }

            return null;
        }

        private FormulaNode Node(FormulaNodeKind kind, ParseTreeNode node) =>
            new FormulaNode(kind, SourceText(node), Start(node));

        private FormulaNode Operator(ParseTreeNode node, string op, OperatorFixity fixity)
        {
            var result = Node(FormulaNodeKind.Operator, node);
            result.Operator = op;
            result.Fixity = fixity;
            return result;
        }

        private static FormulaNode WithChildren(FormulaNode node, params FormulaNode[] children)
        {
            node.Children = children;
            return node;
        }

        private static int Start(ParseTreeNode node) => node.Span.Location.Position;

        private string SourceText(ParseTreeNode node) => _formula.Substring(node.Span.Location.Position, node.Span.Length);
    }

    /// <summary>
    /// Splits a reference prefix (the text up to and including <c>!</c>) into its folder, workbook file and
    /// sheet(s). Quoted prefixes are unquoted first (<c>''</c> is one apostrophe).
    /// </summary>
    internal static (string? Path, string? File, string? Sheet, string? LastSheet) ParsePrefix(string prefix, FormulaReferenceKind kind)
    {
        var text = prefix.EndsWith("!", StringComparison.Ordinal) ? prefix.Substring(0, prefix.Length - 1) : prefix;
        if (text.Length >= 2 && text[0] == '\'' && text[text.Length - 1] == '\'')
        {
            text = text.Substring(1, text.Length - 2).Replace("''", "'");
        }

        string? path = null;
        string? file = null;
        var sheetPart = text;
        var open = text.IndexOf('[');
        var close = open < 0 ? -1 : text.IndexOf(']', open);
        if (close > open)
        {
            // [Book.xlsx]Sheet1 or C:\dir\[Book.xlsx]Sheet1
            path = open > 0 ? text.Substring(0, open) : null;
            file = text.Substring(open + 1, close - open - 1);
            sheetPart = text.Substring(close + 1);
        }
        else if (text.IndexOfAny(new[] { '\\', '/' }) >= 0)
        {
            // 'C:\dir\Book.xlsx'!Name: a workbook-level name in a closed workbook.
            var split = text.LastIndexOfAny(new[] { '\\', '/' });
            path = text.Substring(0, split + 1);
            file = text.Substring(split + 1);
            sheetPart = string.Empty;
        }
        else if ((kind == FormulaReferenceKind.Name || kind == FormulaReferenceKind.StructuredReference) &&
            WorkbookExtensions.Any(extension => text.EndsWith(extension, StringComparison.OrdinalIgnoreCase)))
        {
            // Book.xlsx!Name or Book.xlsx!Table1[Col]: a workbook-level name or table in an open workbook. (Cell
            // references into another workbook always have a [Book] part.)
            file = text;
            sheetPart = string.Empty;
        }

        if (string.IsNullOrWhiteSpace(sheetPart))
        {
            return (path, file, null, null);
        }

        // Sheet names cannot contain ':', so a colon separates the sheets of a 3-D reference.
        var colon = sheetPart.IndexOf(':');
        return colon < 0
            ? (path, file, sheetPart, null)
            : (path, file, sheetPart.Substring(0, colon), sheetPart.Substring(colon + 1));
    }
}
