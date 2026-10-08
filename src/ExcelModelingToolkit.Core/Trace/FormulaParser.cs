using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Threading;
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
/// 0.1 s); later parses take well under a millisecond. Formulas longer than a few hundred characters parse on a
/// background thread with a large stack (the caller waits), since XLParser's recursion could overflow the caller's.
/// No formula, however long or deeply nested, can overflow the stack: the result is parsed, or a failure.
/// </para>
/// </remarks>
public static class FormulaParser
{
    private const string LocalNamePrefix = "_xlpm.";

    // Functions whose result can itself be a reference (docs/PLAN.md section 4.5).
    private static readonly HashSet<string> ReferenceFunctions =
        new HashSet<string>(StringComparer.Ordinal) { "INDEX", "OFFSET", "INDIRECT", "CHOOSE" };

    private static readonly Regex ExternalRefErrorPattern = new Regex(@"^\[([^\[\]]+)\]#REF!", RegexOptions.CultureInvariant);

    // XLParser walks its parse tree recursively, up to about 1 KB of stack per chained range or intersection
    // operator (A1:A2:A3..., A1 A2 A3...), and a stack overflow cannot be caught: it ends the Excel process. For the
    // worst formulas this short, XLParser needs about 2-14 KB of stack beyond the caller's, and the Builder's own
    // recursion is guarded by stack checks; a longer formula parses on LargeStackThread. Those checks are cautious:
    // on .NET Framework they fail once half the thread's stack is used, which Excel's main thread can reach, so a
    // parse that fails for want of stack is retried on LargeStackThread.
    private const int InlineLength = 256;

    // Excel's limit is 8,192 characters, more once a closed workbook's path is written out. LargeStackThread's
    // stack holds a formula of this length several times over.
    private const int MaxLength = 16384;

    // XLParser's time grows with the square of a run of unary operators (=------...1): on .NET Framework 1,000 take
    // about 13 ms, 4,000 150 ms and 16,382 (all a formula of MaxLength holds) 4-17 s, while Excel's main thread waits.
    // Real formulas write two or three in a row at most (--A1 to coerce, -A1%), and no string constant (255
    // characters at most), sheet name (31) or table column name (255) can hold a run this long, so a formula with a
    // longer run is rejected unparsed. Then no formula of MaxLength takes more than about 0.25 s.
    private const int MaxOperatorRun = 1000;

    // The files Excel opens as workbooks, for Book.xlsx!Name. (An unsaved workbook has no extension, so Book2!Rate
    // reads as a sheet-scoped name; the provider resolves that against the open workbooks.) A quoted prefix is
    // ambiguous: 'My Book.xlsx'!Rate is a workbook-level name in a workbook whose file name has a space, but
    // 'Import.csv'!Rate can just as well be a name scoped to a sheet called Import.csv. So these extensions mark a
    // workbook however the prefix is written...
    private static readonly string[] WorkbookExtensions =
    {
        ".xlsx", ".xlsm", ".xlsb", ".xls", ".xltx", ".xltm", ".xlt", ".xlam", ".xla", ".ods",
    };

    // ...and these, formats that rarely hold a name a formula points to (a text file cannot save one), only when
    // unquoted. Either way it is a guess the provider checks against the open workbooks and sheets.
    private static readonly string[] TextFileExtensions = { ".csv", ".txt", ".prn", ".xml", ".slk", ".dif", ".xlw" };

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

        if (formula.Length > MaxLength)
        {
            return ParsedFormula.Failed(formula, context, string.Format(CultureInfo.InvariantCulture,
                "The formula could not be parsed: it is longer than {0:N0} characters.", MaxLength));
        }

        if (HasLongOperatorRun(formula))
        {
            return ParsedFormula.Failed(formula, context, string.Format(CultureInfo.InvariantCulture,
                "The formula could not be parsed: it has more than {0:N0} operators in a row.", MaxOperatorRun));
        }

        if (formula.Length > InlineLength)
        {
            return LargeStackThread.Shared.Parse(formula, context);
        }

        var parsed = ParseOnCurrentThread(formula, context, out var outOfStack);
        return outOfStack ? LargeStackThread.Shared.Parse(formula, context) : parsed;
    }

    // Parse's work, on the calling thread. Internal so tests can run it on a small stack.
    internal static ParsedFormula ParseOnCurrentThread(string formula, FormulaContext context) =>
        ParseOnCurrentThread(formula, context, out _);

    private static ParsedFormula ParseOnCurrentThread(string formula, FormulaContext context, out bool outOfStack)
    {
        outOfStack = false;
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
            // reported the same way: the caller falls back (Trace In can still use Range.DirectPrecedents). So is
            // InsufficientExecutionStackException from the Builder's stack checks, which stop a deep recursion before
            // it overflows.
            outOfStack = ex is InsufficientExecutionStackException;
            return ParsedFormula.Failed(formula, context, "The formula could not be parsed: " + FirstLine(ex.Message));
        }
    }

    private static bool HasLongOperatorRun(string formula)
    {
        var run = 0;
        foreach (var character in formula)
        {
            switch (character)
            {
                case '-':
                case '+':
                case '%':
                    if (++run > MaxOperatorRun)
                    {
                        return true;
                    }

                    break;
                case ' ':
                case '\r':
                case '\n':
                    // Spaces and line breaks may separate the operators of a run.
                    break;
                default:
                    run = 0;
                    break;
            }
        }

        return false;
    }

    private static string FirstLine(string text)
    {
        var end = text.IndexOfAny(new[] { '\r', '\n' });
        return end < 0 ? text : text.Substring(0, end);
    }

    /// <summary>
    /// A background thread with a large stack that parses long formulas, so no formula can overflow the caller's
    /// stack. <see cref="Shared"/> lives as long as the process, so XLParser builds its grammar for it once; tests make
    /// their own.
    /// </summary>
    internal sealed class LargeStackThread : IDisposable
    {
        private const int DefaultStackSize = 16 * 1024 * 1024;

        private const string TooLong = "the parser took too long.";

        private const string Busy = "the parser is still busy with a formula that took too long.";

        // A long formula parses in well under a second. If the grammar ever hangs on one, the caller (Excel's main
        // thread) gives up once the thread has made no progress for this long and falls back, rather than freezing
        // Excel.
        private static readonly TimeSpan Timeout = TimeSpan.FromSeconds(5);

        private readonly int _stackSize;
        private readonly object _startLock = new object();

        // Null until the first request starts the thread.
        private BlockingCollection<Request>? _requests;

        // Stopwatch timestamp of the thread's last progress: when it last started or finished a request.
        private long _progress = Stopwatch.GetTimestamp();

        // The request the thread is running, or null.
        private volatile Request? _running;

        internal LargeStackThread(int stackSize)
        {
            _stackSize = stackSize;
        }

        /// <summary>The thread <see cref="FormulaParser.Parse"/> uses. It starts on the first long formula.</summary>
        public static LargeStackThread Shared { get; } = new LargeStackThread(DefaultStackSize);

        public ParsedFormula Parse(string formula, FormulaContext context) =>
            Run(() => ParseOnCurrentThread(formula, context),
                reason => ParsedFormula.Failed(formula, context, "The formula could not be parsed: " + reason), Timeout);

        /// <summary>
        /// Runs <paramref name="work"/> on the thread and waits for its result. Never throws: returns
        /// <paramref name="fail"/>'s result for a reason instead if the thread cannot start, if <paramref name="work"/>
        /// throws or returns null, or once the thread has made no progress for <paramref name="timeout"/> (not on this
        /// request, and not on requests queued ahead of it). A request that times out is abandoned: the thread skips
        /// it if it has not started it, and drops its late result if it has; while the thread is still running one,
        /// every request fails at once rather than waiting behind it.
        /// </summary>
        internal ParsedFormula Run(Func<ParsedFormula> work, Func<string, ParsedFormula> fail, TimeSpan timeout)
        {
            if (_running is { Abandoned: true })
            {
                // Waiting would only time out again, spinning Excel's main thread for another timeout.
                return fail(Busy);
            }

            var requests = Requests(out var startError);
            if (requests is null)
            {
                return fail("the parser could not start: " + startError);
            }

            var request = new Request(work);
            try
            {
                requests.Add(request);
            }
            catch (InvalidOperationException)
            {
                return fail("the parser has stopped.");
            }

            // Spin rather than block: a blocking wait on Excel's (STA) main thread pumps COM messages, which can
            // re-enter the add-in.
            var queued = Stopwatch.GetTimestamp();
            var limit = (long)(timeout.TotalSeconds * Stopwatch.Frequency);
            while (!request.IsDone)
            {
                var running = _running;
                if (running is not null && running != request && running.Abandoned)
                {
                    // Another caller gave up on the request ahead of this one.
                    request.Abandoned = true;
                    return fail(Busy);
                }

                if (Stopwatch.GetTimestamp() - Math.Max(queued, Interlocked.Read(ref _progress)) >= limit)
                {
                    request.Abandoned = true;
                    return fail(TooLong);
                }

                Thread.Yield();
            }

            return request.Result ?? fail(request.Error ?? "the parser returned no result.");
        }

        /// <summary>Stops the thread once it has run the requests already queued.</summary>
        public void Dispose()
        {
            lock (_startLock)
            {
                _requests?.CompleteAdding();
            }
        }

        // Starts the thread on first use, not in a static initializer: if that fails (in 32-bit Excel, 16 MB of
        // address space for the stack may not be free), every later call would throw TypeInitializationException. A
        // failure is not remembered, so the next long formula tries again, once memory may have been freed.
        private BlockingCollection<Request>? Requests(out string? error)
        {
            lock (_startLock)
            {
                error = null;
                if (_requests is null)
                {
                    try
                    {
                        _requests = Start();
                    }
                    catch (Exception ex)
                    {
                        error = FirstLine(ex.Message);
                    }
                }

                return _requests;
            }
        }

        private BlockingCollection<Request> Start()
        {
            var requests = new BlockingCollection<Request>();
            var thread = new Thread(() => Serve(requests), _stackSize)
            {
                IsBackground = true,
                Name = "Formula parser",
            };
            thread.Start();
            return requests;
        }

        private void Serve(BlockingCollection<Request> requests)
        {
            foreach (var request in requests.GetConsumingEnumerable())
            {
                if (request.Abandoned)
                {
                    continue;
                }

                _running = request;
                Interlocked.Exchange(ref _progress, Stopwatch.GetTimestamp());
                try
                {
                    request.Result = request.Work();
                }
                catch (Exception ex)
                {
                    // An exception escaping a background thread ends the process, and the process is Excel.
                    request.Error = "the parser failed: " + FirstLine(ex.Message);
                }

                request.IsDone = true;
                _running = null;
                Interlocked.Exchange(ref _progress, Stopwatch.GetTimestamp());
            }
        }

        private sealed class Request
        {
            private volatile bool _abandoned;
            private volatile bool _done;

            public Request(Func<ParsedFormula> work)
            {
                Work = work;
            }

            public Func<ParsedFormula> Work { get; }

            // Result and Error are written before IsDone and read after it; its volatile write and read order them.
            public ParsedFormula? Result { get; set; }

            public string? Error { get; set; }

            public bool IsDone
            {
                get => _done;
                set => _done = value;
            }

            public bool Abandoned
            {
                get => _abandoned;
                set => _abandoned = value;
            }
        }
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
            // The recursion follows the formula's nesting. Operator chains, the nesting a formula can repeat
            // thousands of times, are built in a loop (InfixChain); for anything else nested too deeply this check
            // throws (and Parse reports a failure) before the stack overflows.
            RuntimeHelpers.EnsureSufficientExecutionStack();
            if (_parserReferences.TryGetValue(node, out var parserReference))
            {
                return ReferenceLeaf(node, parserReference);
            }

            if (IsExternalRefError(node))
            {
                return ExternalRefError(node);
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
                    var reserved = new FormulaReference(FormulaReferenceKind.Name, _formula, Start(node), Length(node))
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
                    spilled.Length = Length(node);
                    inner.Length = spilled.Length;
                    return inner;
                }

                return WithChildren(Operator(node, "#", OperatorFixity.Postfix), inner);
            }

            if (node.IsIntersection() || node.IsBinaryOperation())
            {
                return InfixChain(node);
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

        // A1+A2+...+An parses left-nested: ((A1+A2)+...)+An, n levels deep, and Excel's 8,192 characters allow
        // thousands of terms. So the left operands are followed in a loop, and the operator nodes built bottom-up
        // (the tree keeps the same shape, and references stay in source order).
        private FormulaNode InfixChain(ParseTreeNode node)
        {
            var chain = new List<ParseTreeNode> { node };
            var first = Unwrap(node.ChildNodes[0]);
            while (IsInfix(first))
            {
                chain.Add(first);
                first = Unwrap(first.ChildNodes[0]);
            }

            var result = Build(first);
            for (var index = chain.Count - 1; index >= 0; index--)
            {
                var operation = chain[index];
                var children = operation.ChildNodes;
                var op = operation.IsIntersection() ? " " : SourceText(children[1]);
                var right = Build(children[2]);
                result = (op == ":" ? BoundingRange(operation, result, right) : null) ??
                    WithChildren(Operator(operation, op, OperatorFixity.Infix), result, right);
            }

            return result;
        }

        // Steps through the single-child Formula / Reference wrappers Build passes straight through.
        private ParseTreeNode Unwrap(ParseTreeNode node)
        {
            while ((node.Is(GrammarNames.Formula) || node.Is(GrammarNames.Reference)) && node.ChildNodes.Count == 1 &&
                !node.IsParentheses() && !_parserReferences.ContainsKey(node))
            {
                node = node.ChildNodes[0];
            }

            return node;
        }

        // True for a node Operation would build as a binary operator or intersection.
        private bool IsInfix(ParseTreeNode node) =>
            (node.Is(GrammarNames.FunctionCall) || node.Is(GrammarNames.ReferenceFunctionCall)) &&
            !_parserReferences.ContainsKey(node) && !IsExternalRefError(node) && !node.IsNamedFunction() &&
            !node.IsUnion() && !(node.ChildNodes.Count == 2 && node.ChildNodes[1].Term.Name == "#") &&
            (node.IsIntersection() || node.IsBinaryOperation());

        // A1:B2:C3 is the range operator applied to the reference A1:B2 and the cell C3; on one sheet its result is
        // the bounding rectangle, A1:C3, which replaces both as one reference. Null when the operands are not two
        // cell or range references on the same sheet.
        private FormulaNode? BoundingRange(ParseTreeNode node, FormulaNode left, FormulaNode right)
        {
            var count = References.Count;
            if (left.Reference is not FormulaReference first || right.Reference is not FormulaReference second ||
                first.Area is not CellRect a || second.Area is not CellRect b || !IsCellOrRange(first) ||
                !IsCellOrRange(second) || first.IsSpill || second.IsSpill || !SameText(first.Sheet, second.Sheet) ||
                !SameText(first.LastSheet, second.LastSheet) || !SameText(first.WorkbookName, second.WorkbookName) ||
                !SameText(first.WorkbookPath, second.WorkbookPath) || count < 2 ||
                !ReferenceEquals(References[count - 2], first) || !ReferenceEquals(References[count - 1], second))
            {
                return null;
            }

            References.RemoveRange(count - 2, 2);
            var reference = new FormulaReference(FormulaReferenceKind.Range, _formula, Start(node), Length(node))
            {
                WorkbookPath = first.WorkbookPath,
                WorkbookName = first.WorkbookName,
                Sheet = first.Sheet,
                LastSheet = first.LastSheet,
                EnclosingReferenceFunction = first.EnclosingReferenceFunction,
            };
            SetAddress(reference, Corner(Math.Min(a.Row, b.Row), Math.Min(a.Column, b.Column)),
                Corner(Math.Max(a.LastRow, b.LastRow), Math.Max(a.LastColumn, b.LastColumn)));
            return Leaf(node, reference);
        }

        private static bool IsCellOrRange(FormulaReference reference) =>
            reference.Kind == FormulaReferenceKind.Cell || reference.Kind == FormulaReferenceKind.Range;

        private static bool SameText(string? first, string? second) =>
            string.Equals(first, second, StringComparison.OrdinalIgnoreCase);

        private static string Corner(int row, int column) =>
            CellRect.ColumnName(column) + row.ToString(CultureInfo.InvariantCulture);

        // XLParser reads [Book.xlsx]#REF!A1 (a reference into a deleted sheet of another workbook) as the table column
        // [Book.xlsx], spilled, intersected with sheet REF's A1. No real formula has "]#REF!" there otherwise
        // (an intersection needs a space). The pattern is matched in place: every link of a long intersection chain
        // is checked, and copying each one's text out would be quadratic.
        private bool IsExternalRefError(ParseTreeNode node) =>
            node.Is(GrammarNames.ReferenceFunctionCall) && node.IsIntersection() &&
            ExternalRefErrorPattern.Match(_formula, Start(node), Length(node)).Success &&
            SourceText(node.ChildNodes[2]).StartsWith("REF!", StringComparison.Ordinal);

        private FormulaNode ExternalRefError(ParseTreeNode node)
        {
            var workbook = ExternalRefErrorPattern.Match(SourceText(node)).Groups[1].Value;
            var reference = new FormulaReference(FormulaReferenceKind.RefError, _formula, Start(node), Length(node))
            {
                WorkbookName = string.Equals(workbook, _context.WorkbookName, StringComparison.OrdinalIgnoreCase) ? null : workbook,
                EnclosingReferenceFunction = CurrentReferenceFunction,
            };
            return Leaf(node, reference);
        }

        private FormulaNode Function(ParseTreeNode spanNode, ParseTreeNode call, string? prefix)
        {
            var name = (prefix ?? string.Empty) + FunctionSignatures.Normalize(SourceText(call.ChildNodes[0]).TrimEnd('('));
            var function = Node(FormulaNodeKind.Function, spanNode);
            function.FunctionName = name;
            function.ReturnsReference = ReferenceFunctions.Contains(name);
            if (function.ReturnsReference)
            {
                DynamicReferences.Add(new DynamicReference(name, _formula, function.Start, function.Length));
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
            return AnchorArray(function) ?? function;
        }

        // _xlfn.ANCHORARRAY(A1) is how Range.Formula writes the spill reference A1# on some versions: the reference,
        // spilled, spanning the whole call. Null if the call is anything else.
        private static FormulaNode? AnchorArray(FormulaNode function)
        {
            if (function.FunctionName != "ANCHORARRAY" || function.Children.Count != 1 ||
                function.Children[0].Kind != FormulaNodeKind.Reference ||
                function.Children[0].Reference is not FormulaReference anchor || anchor.IsSpill ||
                (anchor.Kind != FormulaReferenceKind.Cell && anchor.Kind != FormulaReferenceKind.Range))
            {
                return null;
            }

            var node = function.Children[0];
            anchor.IsSpill = true;
            anchor.Start = function.Start;
            anchor.Length = function.Length;
            node.Start = function.Start;
            node.Length = function.Length;
            return node;
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
            var reference = new FormulaReference(Kind(parsed.ReferenceType), _formula, Start(node), Length(node))
            {
                EnclosingReferenceFunction = CurrentReferenceFunction,
            };

            var prefixNode = FindPrefix(node);
            if (prefixNode is not null && prefixNode.ChildNodes.Any(child => child.Is(GrammarNames.TokenRefError)))
            {
                // #REF!A1: Excel's text for a reference into a deleted sheet. The address no longer means anything.
                reference.Kind = FormulaReferenceKind.RefError;
                return Leaf(node, reference);
            }

            var prefix = prefixNode is null ? default : ParsePrefix(SourceText(prefixNode), reference.Kind);
            reference.WorkbookPath = prefix.Path;
            reference.WorkbookName = prefix.File;
            reference.Sheet = prefix.Sheet;
            reference.LastSheet = prefix.LastSheet;
            reference.IsWorkbookLevelQualified = prefix.File is not null && prefix.Sheet is null;
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
            RuntimeHelpers.EnsureSufficientExecutionStack();
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
            new FormulaNode(kind, _formula, Start(node), Length(node));

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

        private static int Length(ParseTreeNode node) => node.Span.Length;

        private string SourceText(ParseTreeNode node) => _formula.Substring(node.Span.Location.Position, node.Span.Length);
    }

    /// <summary>
    /// Splits a reference prefix (the text up to and including <c>!</c>) into its folder, workbook file and
    /// sheet(s). Quoted prefixes are unquoted first (<c>''</c> is one apostrophe).
    /// </summary>
    internal static (string? Path, string? File, string? Sheet, string? LastSheet) ParsePrefix(string prefix, FormulaReferenceKind kind)
    {
        var text = prefix.EndsWith("!", StringComparison.Ordinal) ? prefix.Substring(0, prefix.Length - 1) : prefix;
        var quoted = text.Length >= 2 && text[0] == '\'' && text[text.Length - 1] == '\'';
        if (quoted)
        {
            text = text.Substring(1, text.Length - 2).Replace("''", "'");
        }

        string? path = null;
        string? file = null;
        var sheetPart = text;

        // Sheet names cannot contain [ ] \ /, and file names cannot contain \ /: so the workbook's closing bracket is
        // the last ']', if no separator follows it, and its opening bracket follows the last separator before it. The
        // folders may have brackets of their own: C:\Deals [2024]\[Book.xlsx]Sheet1.
        var close = text.LastIndexOf(']');
        var open = close < 0 || text.IndexOfAny(new[] { '\\', '/' }, close) >= 0
            ? -1
            : text.LastIndexOfAny(new[] { '\\', '/' }, close) + 1;
        if (open >= 0 && text[open] == '[')
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
            (HasExtension(text, WorkbookExtensions) || (!quoted && HasExtension(text, TextFileExtensions))))
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

    private static bool HasExtension(string file, string[] extensions) =>
        extensions.Any(extension => file.EndsWith(extension, StringComparison.OrdinalIgnoreCase));
}
