using System;
using System.Linq;
using ExcelModelingToolkit.Core.Trace;
using ExcelModelingToolkit.Core.Undo;
using Xunit;

namespace ExcelModelingToolkit.Core.Tests.Trace;

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
    public void Prefixes_split_into_path_workbook_and_sheets(string prefix, string? path, string? file, string? sheet, string? lastSheet)
    {
        var parts = FormulaParser.ParsePrefix(prefix, FormulaReferenceKind.Cell);

        Assert.Equal((path, file, sheet, lastSheet), parts);
    }

    [Fact]
    public void Workbook_like_prefix_is_a_workbook_only_for_names_and_tables()
    {
        Assert.Equal((null, "Book.xlsx", null, null), FormulaParser.ParsePrefix("Book.xlsx!", FormulaReferenceKind.Name));
        Assert.Equal((null, "Book.xlsx", null, null), FormulaParser.ParsePrefix("Book.xlsx!", FormulaReferenceKind.StructuredReference));
        Assert.Equal((null, null, "Book.xlsx", null), FormulaParser.ParsePrefix("Book.xlsx!", FormulaReferenceKind.Cell));
    }

    [Fact]
    public void Parsing_is_fast_after_the_first_call()
    {
        // docs/PLAN.md Phase 4: a formula with 20 references opens in 300 ms, Excel reads included.
        var formula = "=" + string.Join("+", Enumerable.Range(1, 20).Select(i => "'Sheet " + i + "'!$B$" + i));
        Parse(formula);

        var watch = System.Diagnostics.Stopwatch.StartNew();
        for (var i = 0; i < 100; i++)
        {
            Assert.Equal(20, FormulaParser.Parse(formula, Context).References.Count);
        }

        watch.Stop();
        Assert.True(watch.ElapsedMilliseconds < 1000, $"100 parses took {watch.ElapsedMilliseconds} ms.");
    }
}
