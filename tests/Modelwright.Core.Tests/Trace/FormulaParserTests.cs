using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using Modelwright.Core.Trace;
using Modelwright.Core.Undo;
using Xunit;

namespace Modelwright.Core.Tests.Trace;

public class FormulaParserTests
{
    private static readonly FormulaContext Context = new FormulaContext("Model.xlsx", "Calc");

    private static ParsedFormula Parse(string formula)
    {
        var parsed = FormulaParser.Parse(formula, Context);
        Assert.True(parsed.IsParsed, parsed.Error);
        return parsed;
    }

    [Fact]
    public void Spans_cover_each_reference_as_written_including_prefixes()
    {
        const string formula = "=SUM('My Sheet'!$A$1:B$5)+[Book.xlsx]Data!C3*Rate";
        var references = Parse(formula).References;

        Assert.Equal(new[] { "'My Sheet'!$A$1:B$5", "[Book.xlsx]Data!C3", "Rate" }, references.Select(r => r.Text));
        Assert.Equal(new[] { 5, 26, 45 }, references.Select(r => r.Start));
        Assert.Equal(new[] { 19, 18, 4 }, references.Select(r => r.Length));
    }

    [Fact]
    public void Cell_and_range_targets_are_normalized_with_areas()
    {
        // Range.Formula always writes column letters in upper case (XLParser does not read lower-case references).
        var references = Parse("=$B$3+Sheet2!D10:B5+$C:$E+7:9").References;

        Assert.Equal("B3", references[0].Address);
        Assert.Equal(new CellRect(3, 2, 1, 1), references[0].Area);
        Assert.Equal("B5:D10", references[1].Address);
        Assert.Equal("Sheet2", references[1].Sheet);
        Assert.Equal(new CellRect(5, 2, 6, 3), references[1].Area);
        Assert.Equal("C:E", references[2].Address);
        Assert.Equal(new CellRect(1, 3, CellRect.MaxRows, 3), references[2].Area);
        Assert.Equal("7:9", references[3].Address);
        Assert.Equal(new CellRect(7, 1, 3, CellRect.MaxColumns), references[3].Area);
    }

    [Fact]
    public void Names_tables_and_errors_have_no_address()
    {
        var references = Parse("=Rate+Sales[Amount]+#REF!").References;

        Assert.All(references, reference =>
        {
            Assert.Null(reference.Address);
            Assert.Null(reference.Area);
        });
        Assert.Equal(new[] { FormulaReferenceKind.Name, FormulaReferenceKind.StructuredReference, FormulaReferenceKind.RefError },
            references.Select(r => r.Kind));
    }

    [Fact]
    public void External_reference_with_path_reports_workbook_path_and_sheet()
    {
        var reference = Parse("='C:\\Deals\\Project X\\[LBO v3.xlsm]Returns Summary'!$F$12").References.Single();

        Assert.True(reference.IsExternal);
        Assert.Equal("C:\\Deals\\Project X\\", reference.WorkbookPath);
        Assert.Equal("LBO v3.xlsm", reference.WorkbookName);
        Assert.Equal("Returns Summary", reference.Sheet);
        Assert.Equal("F12", reference.Address);
    }

    [Fact]
    public void External_reference_to_an_open_workbook_has_no_path()
    {
        var reference = Parse("=[Comps.xlsx]Output!B7").References.Single();

        Assert.Equal("Comps.xlsx", reference.WorkbookName);
        Assert.Null(reference.WorkbookPath);
        Assert.Equal("Output", reference.Sheet);
    }

    [Fact]
    public void References_to_the_formulas_own_workbook_and_sheet_come_back_unqualified()
    {
        var references = Parse("=[MODEL.XLSX]calc!A1+CALC!A2+[Model.xlsx]Other!A3").References;

        Assert.All(references, reference => Assert.False(reference.IsExternal));
        Assert.Null(references[0].Sheet);
        Assert.Null(references[1].Sheet);
        Assert.Equal("Other", references[2].Sheet);
    }

    [Fact]
    public void Three_d_reference_reports_first_and_last_sheet()
    {
        var reference = Parse("=SUM('Jan 2024:Dec 2024'!B5)").References.Single();

        Assert.True(reference.Is3D);
        Assert.Equal("Jan 2024", reference.Sheet);
        Assert.Equal("Dec 2024", reference.LastSheet);
    }

    [Fact]
    public void Three_d_reference_through_the_formulas_own_sheet_keeps_its_sheets()
    {
        var reference = Parse("=SUM(Calc:Summary!B5)").References.Single();

        Assert.Equal("Calc", reference.Sheet);
        Assert.Equal("Summary", reference.LastSheet);
    }

    [Fact]
    public void Spill_reference_includes_the_hash_in_its_span()
    {
        const string formula = "=SUM(Data!B2#)*2";
        var reference = Parse(formula).References.Single();

        Assert.True(reference.IsSpill);
        Assert.Equal("Data!B2#", reference.Text);
        Assert.Equal(5, reference.Start);
        Assert.Equal("B2", reference.Address);
    }

    [Fact]
    public void Anchorarray_is_a_spill_reference_spanning_the_call()
    {
        // Range.Formula's form of =SUM(Data!B2#)*2 on versions that write spill references as a function.
        const string formula = "=SUM(_xlfn.ANCHORARRAY(Data!B2))*2";
        var parsed = Parse(formula);
        var reference = parsed.References.Single();

        Assert.True(reference.IsSpill);
        Assert.Equal("_xlfn.ANCHORARRAY(Data!B2)", reference.Text);
        Assert.Equal(5, reference.Start);
        Assert.Equal("B2", reference.Address);
        Assert.Equal("Data", reference.Sheet);
        var sum = parsed.Structure!.Children[0];
        Assert.Equal(FormulaNodeKind.Reference, sum.Children[0].Kind);
        Assert.Same(reference, sum.Children[0].Reference);
        Assert.Equal(0, sum.Children[0].ArgumentIndex);
    }

    [Fact]
    public void Anchorarray_of_something_other_than_a_cell_stays_a_function()
    {
        var parsed = Parse("=ROWS(_xlfn.ANCHORARRAY(INDEX(A1:A9,2)))");

        Assert.False(parsed.References.Single().IsSpill);
        Assert.Equal("ANCHORARRAY", parsed.Structure!.Children[0].FunctionName);
    }

    [Fact]
    public void A_name_qualified_with_the_formulas_own_workbook_is_marked_workbook_level()
    {
        var own = Parse("=Model.xlsx!Rate+Rate+[Model.xlsx]Calc!Tax").References;

        Assert.Null(own[0].WorkbookName);
        Assert.True(own[0].IsWorkbookLevelQualified);
        Assert.False(own[1].IsWorkbookLevelQualified);
        Assert.False(own[2].IsWorkbookLevelQualified);
        Assert.True(Parse("=Book.xlsx!Rate").References.Single().IsWorkbookLevelQualified);
    }

    [Fact]
    public void Unqualified_structured_reference_needs_table_context()
    {
        var references = Parse("=[@Qty]*Sales[@Price]").References;

        Assert.True(references[0].NeedsTableContext);
        Assert.Null(references[0].Name);
        Assert.False(references[1].NeedsTableContext);
        Assert.Equal("Sales", references[1].Name);
        Assert.Equal(new[] { "Price" }, references[1].TableColumns);
    }

    [Fact]
    public void Built_in_reserved_name_is_a_name()
    {
        var reference = Parse("=ROWS(_xlnm.Print_Area)").References.Single();

        Assert.Equal(FormulaReferenceKind.Name, reference.Kind);
        Assert.Equal("_xlnm.Print_Area", reference.Name);
        Assert.Equal(6, reference.Start);
    }

    [Fact]
    public void Local_names_are_not_defined_names()
    {
        var references = Parse("=LET(rate,Rate*2,rate+1)").References;

        Assert.Equal(new[] { FormulaReferenceKind.LocalName, FormulaReferenceKind.Name, FormulaReferenceKind.LocalName },
            references.Select(r => r.Kind));
    }

    [Fact]
    public void Local_name_scope_ends_with_its_let()
    {
        var references = Parse("=LET(a,1,a)+a").References;

        Assert.Equal(FormulaReferenceKind.Name, references.Last().Kind);
    }

    [Fact]
    public void Sheet_qualified_name_is_never_local()
    {
        var references = Parse("=LET(x,1,Inputs!x+x)").References;

        Assert.Equal(new[] { FormulaReferenceKind.LocalName, FormulaReferenceKind.Name, FormulaReferenceKind.LocalName },
            references.Select(r => r.Kind));
        Assert.Equal("Inputs", references[1].Sheet);
    }

    [Fact]
    public void Dynamic_reference_functions_are_listed_and_flag_their_references()
    {
        var parsed = Parse("=SUM(OFFSET(A1,0,0,3,1))+INDIRECT(\"B\"&C1)+INDEX(D1:F9,2,CHOOSE(2,1,G1))+H1");

        Assert.True(parsed.HasDynamicReferences);
        Assert.Equal(new[] { "OFFSET", "INDIRECT", "INDEX", "CHOOSE" }, parsed.DynamicReferences.Select(d => d.FunctionName));
        Assert.Equal("OFFSET(A1,0,0,3,1)", parsed.DynamicReferences[0].Text);
        Assert.Equal(5, parsed.DynamicReferences[0].Start);
        Assert.Equal(new[] { "OFFSET", "INDIRECT", "INDEX", "CHOOSE", null },
            parsed.References.Select(r => r.EnclosingReferenceFunction));
    }

    [Fact]
    public void Formula_without_dynamic_functions_has_none()
    {
        var parsed = Parse("=SUM(A1:A3)");

        Assert.False(parsed.HasDynamicReferences);
        Assert.Empty(parsed.DynamicReferences);
    }

    [Theory]
    [InlineData("")]
    [InlineData("=")]
    [InlineData("= ")]
    [InlineData("100")]
    [InlineData("SUM(A1)")]
    public void Text_that_is_not_a_formula_is_reported_not_thrown(string text)
    {
        var parsed = FormulaParser.Parse(text, Context);

        Assert.False(parsed.IsParsed);
        Assert.NotNull(parsed.Error);
        Assert.Empty(parsed.References);
        Assert.Null(parsed.Structure);
    }

    [Theory]
    [InlineData("=SUM(A1")]
    [InlineData("=A1+")]
    [InlineData("=LAMBDA(x,x+1)(5)")]
    public void Formula_the_grammar_rejects_is_reported_not_thrown(string formula)
    {
        // XLParser 1.7.5 cannot parse an immediately invoked LAMBDA; the add-in falls back to Excel's own
        // precedents for such cells.
        var parsed = FormulaParser.Parse(formula, Context);

        Assert.False(parsed.IsParsed);
        Assert.StartsWith("The formula could not be parsed", parsed.Error);
        Assert.Empty(parsed.References);
    }

    [Fact]
    public void Arguments_are_checked()
    {
        Assert.Throws<ArgumentNullException>(() => FormulaParser.Parse(null!, Context));
        Assert.Throws<ArgumentNullException>(() => FormulaParser.Parse("=A1", null!));
        Assert.Throws<ArgumentException>(() => new FormulaContext("", "Calc"));
        Assert.Throws<ArgumentException>(() => new FormulaContext("Model.xlsx", ""));
    }

    [Theory]
    [InlineData("'Sheet1'!", null, null, "Sheet1", null)]
    [InlineData("'It''s'!", null, null, "It's", null)]
    [InlineData("Jan:Dec!", null, null, "Jan", "Dec")]
    [InlineData("[Book.xlsx]Data!", null, "Book.xlsx", "Data", null)]
    [InlineData("'C:\\a b\\[It''s.xlsx]S 1'!", "C:\\a b\\", "It's.xlsx", "S 1", null)]
    [InlineData("'C:\\dir\\Book.xlsx'!", "C:\\dir\\", "Book.xlsx", null, null)]
    [InlineData("'https://x.com/a/[B.xlsx]S'!", "https://x.com/a/", "B.xlsx", "S", null)]
    [InlineData("'C:\\Deals [2024]\\[Book.xlsx]Sheet1'!", "C:\\Deals [2024]\\", "Book.xlsx", "Sheet1", null)]
    [InlineData("'C:\\Deals [2024]\\Book.xlsx'!", "C:\\Deals [2024]\\", "Book.xlsx", null, null)]
    [InlineData("'C:\\Deals [2024]\\[Book.xlsx]Jan:Dec'!", "C:\\Deals [2024]\\", "Book.xlsx", "Jan", "Dec")]
    [InlineData("'\\\\srv\\share [x]\\[B.xlsx]S'!", "\\\\srv\\share [x]\\", "B.xlsx", "S", null)]
    [InlineData("'https://x.com/a [1]/[B.xlsx]S'!", "https://x.com/a [1]/", "B.xlsx", "S", null)]
    [InlineData("'https://x.com/a [1]/B.xlsx'!", "https://x.com/a [1]/", "B.xlsx", null, null)]
    [InlineData("'C:\\d\\[Book [v2].xlsx]S'!", "C:\\d\\", "Book [v2].xlsx", "S", null)]
    public void Prefixes_split_into_path_workbook_and_sheets(string prefix, string? path, string? file, string? sheet, string? lastSheet)
    {
        var parts = FormulaParser.ParsePrefix(prefix, FormulaReferenceKind.Cell);

        Assert.Equal((path, file, sheet, lastSheet), parts);
    }

    [Fact]
    public void Workbook_like_prefix_is_a_workbook_only_for_names_and_tables()
    {
        Assert.Equal((null, "Data.ods", null, null), FormulaParser.ParsePrefix("Data.ods!", FormulaReferenceKind.Name));
        Assert.Equal((null, "Book.xlsx", null, null), FormulaParser.ParsePrefix("Book.xlsx!", FormulaReferenceKind.Name));
        Assert.Equal((null, "Book.xlsx", null, null), FormulaParser.ParsePrefix("Book.xlsx!", FormulaReferenceKind.StructuredReference));
        Assert.Equal((null, null, "Book.xlsx", null), FormulaParser.ParsePrefix("Book.xlsx!", FormulaReferenceKind.Cell));
    }

    [Fact]
    public void Quoted_prefix_with_a_text_file_extension_is_a_sheet_but_with_a_workbook_extension_a_workbook()
    {
        Assert.Equal((null, "My Book.xlsx", null, null), FormulaParser.ParsePrefix("'My Book.xlsx'!", FormulaReferenceKind.Name));
        Assert.Equal((null, "Data.ods", null, null), FormulaParser.ParsePrefix("'Data.ods'!", FormulaReferenceKind.StructuredReference));
        Assert.Equal((null, "Import.csv", null, null), FormulaParser.ParsePrefix("Import.csv!", FormulaReferenceKind.Name));
        Assert.Equal((null, null, "Import.csv", null), FormulaParser.ParsePrefix("'Import.csv'!", FormulaReferenceKind.Name));
        Assert.Equal((null, null, "Notes.txt", null), FormulaParser.ParsePrefix("'Notes.txt'!", FormulaReferenceKind.Name));
        Assert.Equal(("C:\\dir\\", "Import.csv", null, null), FormulaParser.ParsePrefix("'C:\\dir\\Import.csv'!", FormulaReferenceKind.Name));

        var reference = Parse("='Import.csv'!Rate").References.Single();
        Assert.Equal(FormulaReferenceKind.Name, reference.Kind);
        Assert.Equal("Import.csv", reference.Sheet);
        Assert.Null(reference.WorkbookName);
    }

    [Fact]
    public void Reference_into_a_deleted_sheet_is_a_ref_error()
    {
        var references = Parse("=#REF!A1+#REF!B1:B5+[Budget.xlsx]#REF!C3+'#REF'!D4").References;

        Assert.Equal(
            new[] { FormulaReferenceKind.RefError, FormulaReferenceKind.RefError, FormulaReferenceKind.RefError, FormulaReferenceKind.Cell },
            references.Select(r => r.Kind));
        Assert.Equal(new[] { "#REF!A1", "#REF!B1:B5", "[Budget.xlsx]#REF!C3", "'#REF'!D4" }, references.Select(r => r.Text));
        Assert.All(references.Take(3), reference =>
        {
            Assert.Null(reference.Sheet);
            Assert.Null(reference.Address);
            Assert.Null(reference.Area);
        });
        Assert.Null(references[0].WorkbookName);
        Assert.Equal("Budget.xlsx", references[2].WorkbookName);

        // A quoted '#REF' is a sheet that really has that name.
        Assert.Equal("#REF", references[3].Sheet);
    }

    // A formula may use all of Excel's 8,192 characters, and XLParser recurses once per chained range or
    // intersection operator. Each case runs on a thread with a 256 KB stack (Excel's main thread has 1 MB, some of it
    // in use): the formula is parsed with every reference, or fails, but never overflows the stack, which would end
    // the test process (and Excel).
    [Theory]
    [InlineData("+")]
    [InlineData("&")]
    [InlineData(" ")]
    [InlineData(":")]
    public void Long_operator_chain_parses_on_a_small_stack(string op)
    {
        var formula = "=" + string.Join(op, Enumerable.Repeat("A1", 2700));

        var parsed = OnSmallStack(() => FormulaParser.Parse(formula, Context));

        Assert.True(parsed.IsParsed, parsed.Error);
        var expected = op == ":" ? 1 : 2700;
        Assert.Equal(expected, parsed.References.Count);
        Assert.Equal(expected, OnSmallStack(() => parsed.TopLevelNodes.Count));
        Assert.Equal(op == ":" ? 0 : expected, OnSmallStack(() => parsed.Structure!.TraceChildren.Count));
        Assert.Equal(parsed.References.Select(r => r.Start).OrderBy(start => start), parsed.References.Select(r => r.Start));
    }

    [Fact]
    public void Long_mixed_operator_chain_parses_on_a_small_stack()
    {
        var ops = new[] { "+", "-", "*", "/", "^", "&", "=", "<>", "<=", ">", "+", "&" };
        var formula = "=B1" + string.Concat(Enumerable.Range(0, 2000).Select(i => ops[i % ops.Length] + "B" + ((i % 9) + 1)));
        Assert.True(formula.Length <= 8192, $"The formula has {formula.Length} characters.");

        var parsed = OnSmallStack(() => FormulaParser.Parse(formula, Context));

        Assert.True(parsed.IsParsed, parsed.Error);
        Assert.Equal(2001, parsed.References.Count);
        Assert.Equal(2001, OnSmallStack(() => parsed.TopLevelNodes.Count));
        Assert.All(parsed.References, reference => Assert.Equal(reference.Text, formula.Substring(reference.Start, reference.Length)));
    }

    [Fact]
    public void Long_range_operator_chain_is_one_bounding_range()
    {
        var formula = "=SUM(" + string.Join(":", Enumerable.Range(1, 1500).Select(row => "A" + row)) + ")";

        var parsed = OnSmallStack(() => FormulaParser.Parse(formula, Context));

        Assert.True(parsed.IsParsed, parsed.Error);
        Assert.Equal("A1:A1500", parsed.References.Single().Address);
        Assert.Equal(FormulaNodeKind.Reference, parsed.TopLevelNodes.Single().TraceChildren.Single().Kind);
    }

    [Theory]
    [InlineData(" ")]
    [InlineData(":")]
    public void Reference_chain_short_enough_to_parse_on_the_callers_thread_fits_a_small_stack(string op)
    {
        // The longest formula Parse runs on the calling thread (256 characters).
        var formula = "=" + string.Join(op, Enumerable.Repeat("A1", 85));

        var parsed = OnSmallStack(() => FormulaParser.ParseOnCurrentThread(formula, Context));

        Assert.True(parsed.IsParsed, parsed.Error);
    }

    [Fact]
    public void Deep_parentheses_parse_on_a_small_stack()
    {
        var formula = "=" + new string('(', 4000) + "A1" + new string(')', 4000);

        var parsed = OnSmallStack(() => FormulaParser.Parse(formula, Context));

        Assert.True(parsed.IsParsed, parsed.Error);
        Assert.Equal("A1", parsed.References.Single().Text);
        Assert.Equal(FormulaNodeKind.Group, parsed.TopLevelNodes.Single().Kind);
    }

    [Theory]
    [InlineData(64)]
    [InlineData(1300)]
    public void Deeply_nested_functions_parse_on_a_small_stack(int depth)
    {
        // Excel allows 64 levels; 1,300 fit in 8,192 characters (XLParser alone overflows 256 KB at about 600).
        var formula = "=" + string.Concat(Enumerable.Repeat("IF(A1,", depth)) + "B1" + new string(')', depth);

        var parsed = OnSmallStack(() => FormulaParser.Parse(formula, Context));

        Assert.True(parsed.IsParsed, parsed.Error);
        Assert.Equal(depth + 1, parsed.References.Count);
        Assert.Equal("IF", parsed.TopLevelNodes.Single().FunctionName);
    }

    [Fact]
    public void Recursion_too_deep_for_the_stack_fails_instead_of_overflowing()
    {
        // 1,000 nested negations, the most Parse accepts in a row: XLParser reads them, and building the structure
        // runs out of a 256 KB stack.
        var formula = "=" + new string('-', 1000) + "A1";

        var parsed = OnSmallStack(() => FormulaParser.ParseOnCurrentThread(formula, Context));

        AssertFailed(parsed);
        Assert.True(OnSmallStack(() => FormulaParser.Parse(formula, Context)).IsParsed);
    }

    [Fact]
    public void Formula_longer_than_the_limit_fails()
    {
        var parsed = FormulaParser.Parse("=" + string.Join("+", Enumerable.Repeat("A1", 5462)), Context);

        AssertFailed(parsed);
    }

    [Fact]
    public void Formula_with_more_than_a_thousand_operators_in_a_row_fails_at_once()
    {
        // 16,382 negations, all a formula of the maximum length holds: XLParser alone took 4-13 s on .NET Framework.
        var watch = Stopwatch.StartNew();
        var parsed = FormulaParser.Parse("=" + new string('-', 16382) + "1", Context);
        watch.Stop();

        AssertFailed(parsed);
        Assert.Contains("more than 1,000 operators in a row", parsed.Error);
        Assert.True(watch.ElapsedMilliseconds < 1000, $"The rejection took {watch.ElapsedMilliseconds} ms.");
        AssertFailed(FormulaParser.Parse("=A1" + new string('%', 1001), Context));
        AssertFailed(FormulaParser.Parse("=" + string.Concat(Enumerable.Repeat("- ", 1001)) + "1", Context));
        Assert.True(FormulaParser.Parse("=" + new string('-', 1000) + "1", Context).IsParsed);
        Assert.True(FormulaParser.Parse("=" + string.Join("*", Enumerable.Repeat(new string('-', 1000) + "1", 16)), Context).IsParsed);
    }

    [Fact]
    public async Task Request_that_times_out_while_queued_is_skipped_by_the_parse_thread()
    {
        using var thread = new FormulaParser.LargeStackThread(1024 * 1024);
        using var started = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var first = Task.Run(() => thread.Run(
            () =>
            {
                started.Set();
                release.Wait();
                return Parsed;
            },
            Fail,
            TimeSpan.FromSeconds(30)));
        Assert.True(started.Wait(TimeSpan.FromSeconds(10)));
        var ran = 0;

        var queued = thread.Run(
            () =>
            {
                Interlocked.Increment(ref ran);
                return Parsed;
            },
            Fail,
            TimeSpan.FromMilliseconds(200));
        release.Set();

        Assert.Contains("took too long", queued.Error);
        Assert.Same(Parsed, await first);
        Assert.Same(Parsed, thread.Run(() => Parsed, Fail, TimeSpan.FromSeconds(10)));
        Assert.Equal(0, Volatile.Read(ref ran));
    }

    [Fact]
    public void While_a_timed_out_request_still_runs_new_requests_fail_at_once_then_parsing_works_again()
    {
        using var thread = new FormulaParser.LargeStackThread(1024 * 1024);
        using var started = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        Assert.Same(Parsed, thread.Run(() => Parsed, Fail, TimeSpan.FromSeconds(10)));

        var timedOut = thread.Run(
            () =>
            {
                started.Set();
                release.Wait();
                return Parsed;
            },
            Fail,
            TimeSpan.FromSeconds(1));
        Assert.True(started.Wait(TimeSpan.FromSeconds(10)), "The request was not running when it timed out.");
        var watch = Stopwatch.StartNew();
        var busy = thread.Run(() => Parsed, Fail, TimeSpan.FromSeconds(30));
        watch.Stop();
        release.Set();

        Assert.Contains("took too long", timedOut.Error);
        Assert.Contains("still busy", busy.Error);
        Assert.True(watch.ElapsedMilliseconds < 1000, $"The busy request waited {watch.ElapsedMilliseconds} ms.");
        Assert.True(SpinWait.SpinUntil(() => thread.Run(() => Parsed, Fail, TimeSpan.FromSeconds(10)).IsParsed, TimeSpan.FromSeconds(10)));
    }

    [Fact]
    public void Parse_thread_reports_an_exception_from_its_work_as_a_failure_and_keeps_running()
    {
        using var thread = new FormulaParser.LargeStackThread(1024 * 1024);

        var failed = thread.Run(() => throw new InvalidOperationException("Boom"), Fail, TimeSpan.FromSeconds(10));

        Assert.Equal("the parser failed: Boom", failed.Error);
        Assert.Same(Parsed, thread.Run(() => Parsed, Fail, TimeSpan.FromSeconds(10)));
    }

    [Fact]
    public void Parse_thread_that_cannot_start_fails_each_request_without_throwing()
    {
        // A negative stack size makes the thread's constructor throw, as an out-of-memory start would.
        using var thread = new FormulaParser.LargeStackThread(-1);

        var first = thread.Run(() => Parsed, Fail, TimeSpan.FromSeconds(10));
        var second = thread.Run(() => Parsed, Fail, TimeSpan.FromSeconds(10));

        Assert.StartsWith("the parser could not start: ", first.Error);
        Assert.StartsWith("the parser could not start: ", second.Error);
    }

    [Fact]
    public void Short_formula_parses_on_the_large_stack_thread_when_the_callers_stack_is_nearly_used()
    {
        const string formula = "=SUM(A1,B1)";

        var (inline, parsed) = OnSmallStack(() =>
        {
            // The first parse on a thread builds XLParser's grammar; do that before using the stack up.
            FormulaParser.Parse("=1", Context);
            return AtStackLimit(() => (FormulaParser.ParseOnCurrentThread(formula, Context), FormulaParser.Parse(formula, Context)));
        });

        AssertFailed(inline);
        Assert.True(parsed.IsParsed, parsed.Error);
        Assert.Equal(new[] { "A1", "B1" }, parsed.References.Select(reference => reference.Text));
    }

    private static readonly ParsedFormula Parsed = FormulaParser.Parse("=1", Context);

    private static ParsedFormula Fail(string reason) => ParsedFormula.Failed("=1", Context, reason);

    // Recurses until the runtime's stack check fails, then runs work there.
    [MethodImpl(MethodImplOptions.NoInlining)]
    private static T AtStackLimit<T>(Func<T> work)
    {
        bool atLimit;
        try
        {
            RuntimeHelpers.EnsureSufficientExecutionStack();
            atLimit = false;
        }
        catch (InsufficientExecutionStackException)
        {
            atLimit = true;
        }

        if (atLimit)
        {
            return work();
        }

        var result = AtStackLimit(work);
        GC.KeepAlive(work); // Not a tail call, which could reuse the frame.
        return result;
    }

    private static void AssertFailed(ParsedFormula parsed)
    {
        Assert.False(parsed.IsParsed);
        Assert.StartsWith("The formula could not be parsed", parsed.Error);
        Assert.Empty(parsed.References);
        Assert.Empty(parsed.TopLevelNodes);
    }

    private static T OnSmallStack<T>(Func<T> work)
    {
        var result = default(T)!;
        Exception? error = null;
        var thread = new Thread(
            () =>
            {
                try
                {
                    result = work();
                }
                catch (Exception ex)
                {
                    error = ex;
                }
            },
            256 * 1024);
        thread.Start();
        thread.Join();
        if (error is not null)
        {
            throw new InvalidOperationException("The work threw on the small-stack thread.", error);
        }

        return result;
    }
}
