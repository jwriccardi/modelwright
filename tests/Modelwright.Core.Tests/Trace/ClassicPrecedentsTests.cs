using System;
using System.Linq;
using Modelwright.Core.Trace;
using Xunit;

namespace Modelwright.Core.Tests.Trace;

public class ClassicPrecedentsTests
{
    private static readonly FormulaContext Context = new FormulaContext("Model.xlsx", "Calc");

    private static string Rows(string formula) =>
        string.Join("; ", ClassicPrecedents.Of(FormulaParser.Parse(formula, Context))
            .Select(r => (r.Dynamic is null ? string.Empty : "dyn ") + r.Text));

    [Fact]
    public void References_come_in_the_order_written_with_duplicates()
    {
        Assert.Equal("A1; Sheet2!B5; A1; Rate", Rows("=A1+Sheet2!B5*A1-Rate"));
    }

    [Fact]
    public void A_dynamic_reference_comes_before_the_references_inside_it()
    {
        Assert.Equal("B1; dyn INDEX(A1:C9,2,3); A1:C9; D1", Rows("=B1+INDEX(A1:C9,2,3)+D1"));
        Assert.Equal("dyn INDIRECT(\"Sheet2!A1\")", Rows("=INDIRECT(\"Sheet2!A1\")"));
    }

    [Fact]
    public void Nested_dynamic_references_come_outer_first()
    {
        Assert.Equal("dyn INDEX(OFFSET(A1,1,1,3,3),1,1); dyn OFFSET(A1,1,1,3,3); A1", Rows("=INDEX(OFFSET(A1,1,1,3,3),1,1)"));
    }

    [Fact]
    public void Local_names_are_left_out()
    {
        Assert.Equal("Rate; A1", Rows("=LET(x,Rate,x*A1)"));
    }

    [Fact]
    public void A_constant_or_an_unparsed_formula_has_no_rows()
    {
        Assert.Empty(ClassicPrecedents.Of(FormulaParser.Parse("=1+2", Context)));
        Assert.Empty(ClassicPrecedents.Of(FormulaParser.Parse("=SUM(", Context)));
        Assert.Throws<ArgumentNullException>(() => ClassicPrecedents.Of(null!));
    }

    [Fact]
    public void Each_row_has_one_part()
    {
        var rows = ClassicPrecedents.Of(FormulaParser.Parse("=OFFSET(A1,1,0)", Context));

        Assert.NotNull(rows[0].Dynamic);
        Assert.Null(rows[0].Reference);
        Assert.NotNull(rows[1].Reference);
        Assert.Null(rows[1].Dynamic);
        Assert.Equal("A1", rows[1].ToString());
    }
}
