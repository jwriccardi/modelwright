using System;
using System.Linq;
using ExcelModelingToolkit.Core.Trace;
using Xunit;

namespace ExcelModelingToolkit.Core.Tests.Trace;

public class NameLookupTests
{
    private static readonly FormulaContext Context = new FormulaContext("Model.xlsx", "Calc");

    private static string Candidates(string formula)
    {
        var reference = FormulaParser.Parse(formula, Context).References.Single();
        return string.Join("; ", NameLookup.Candidates(reference, Context).Select(c => c + (c.MayOpenWorkbook ? " (may open)" : string.Empty)));
    }

    [Fact]
    public void An_unqualified_name_is_looked_for_in_the_sheet_scope_then_the_workbook()
    {
        Assert.Equal("Calc!Rate; Rate", Candidates("=Rate*2"));
    }

    [Fact]
    public void A_sheet_qualified_name_is_that_sheets_then_an_open_workbook_of_that_name()
    {
        // Book2!Rate: a sheet called Book2, or a workbook-level name in the unsaved Book2.
        Assert.Equal("Book2!Rate; [Book2]Rate", Candidates("=Book2!Rate"));
    }

    [Fact]
    public void A_workbook_qualified_name_is_that_workbooks_and_may_open_it()
    {
        Assert.Equal("[Book.xlsx]Rate (may open)", Candidates("=Book.xlsx!Rate"));
    }

    [Fact]
    public void A_name_qualified_with_the_formulas_own_workbook_is_only_the_workbook_level_one()
    {
        // Calc may have its own Rate; Model.xlsx!Rate means the workbook's.
        Assert.Equal("Rate", Candidates("=Model.xlsx!Rate"));
    }

    [Fact]
    public void An_external_sheet_scoped_name_is_that_sheets()
    {
        Assert.Equal("[Book.xlsx]Inputs!Rate (may open)", Candidates("=[Book.xlsx]Inputs!Rate"));
    }

    [Fact]
    public void The_prefix_of_a_built_in_name_is_dropped()
    {
        var reference = FormulaParser.Parse("=ROWS(_xlnm.Print_Area)", Context).References.Single();

        var candidates = NameLookup.Candidates(reference, Context);

        Assert.All(candidates, c => Assert.Equal("Print_Area", c.Name));
        Assert.Equal("Calc", candidates[0].SheetName);
        Assert.Null(candidates[0].WorkbookName);
    }

    [Fact]
    public void Only_names_are_accepted()
    {
        var cell = FormulaParser.Parse("=A1", Context).References.Single();

        Assert.Throws<ArgumentException>(() => NameLookup.Candidates(cell, Context));
        Assert.Throws<ArgumentNullException>(() => NameLookup.Candidates(null!, Context));
        Assert.Throws<ArgumentNullException>(() => NameLookup.Candidates(cell, null!));
    }
}
