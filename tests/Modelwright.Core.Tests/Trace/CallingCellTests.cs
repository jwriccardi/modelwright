using System;
using System.Linq;
using Modelwright.Core.Trace;
using Xunit;

namespace Modelwright.Core.Tests.Trace;

public class CallingCellTests
{
    private static readonly FormulaContext Context = new FormulaContext("Model.xlsx", "Calc");

    [Theory]
    [InlineData("OFFSET($A$1,ROW()-1,0)", "OFFSET($A$1,14-1,0)")]
    [InlineData("INDEX(A:A,row())+column()", "INDEX(A:A,14)+3")]
    [InlineData("OFFSET(A1,ROW( )-1,COLUMN(  ))", "OFFSET(A1,14-1,3)")]
    [InlineData("=ROW()", "=14")]
    public void Argument_less_row_and_column_become_the_formula_cells(string expression, string expected)
    {
        Assert.Equal(expected, CallingCell.SubstituteRowAndColumn(expression, 14, 3));
    }

    [Theory]
    [InlineData("INDEX(A1:A9,ROWS(A1:A3))")]
    [InlineData("OFFSET(A1,COLUMNS(B1:C1),0)")]
    [InlineData("OFFSET(A1,ROW(B5),COLUMN(C1))")]
    [InlineData("INDIRECT(\"ROW()\")")]
    [InlineData("SUM('ROW()'!A1)")]
    [InlineData("SUM(Table1[ROW()])")]
    [InlineData("MYROW()+COLUMNX()+A1ROW()")]
    [InlineData("INDIRECT(\"a\"\"ROW()\")")]
    public void Other_calls_strings_sheet_names_and_table_columns_are_left_alone(string expression)
    {
        Assert.Equal(expression, CallingCell.SubstituteRowAndColumn(expression, 14, 3));
    }

    [Theory]
    [InlineData("INDIRECT(\"RC[-1]\",FALSE)", true)]
    [InlineData("INDIRECT(B1,0)", true)]
    [InlineData("INDIRECT(\"R2C3\",FALSE)", false)]
    [InlineData("INDIRECT(\"Inputs!R1C1:R5C2\",FALSE)", false)]
    [InlineData("INDIRECT(\"B3\")", false)]
    [InlineData("INDIRECT(\"B3\",TRUE)", false)]
    [InlineData("INDIRECT(\"B3\",1)", false)]
    [InlineData("INDIRECT(\"B3\",)", false)]
    [InlineData("OFFSET(INDIRECT(\"R[1]C\",FALSE),1,0)", true)]
    [InlineData("SUM(A1)", false)]
    public void R1C1_indirect_with_possibly_relative_text_is_detected(string expression, bool expected)
    {
        Assert.Equal(expected, CallingCell.HasRelativeR1C1Indirect(expression));
    }

    [Theory]
    [InlineData("=$A$1", false)]
    [InlineData("=Inputs!$A$1:$B$5", false)]
    [InlineData("='My Sheet'!$C:$D", false)]
    [InlineData("=$3:$5", false)]
    [InlineData("=A1", true)]
    [InlineData("=$A1", true)]
    [InlineData("=A$1", true)]
    [InlineData("=Inputs!$A$1:B5", true)]
    [InlineData("=C:$D", true)]
    [InlineData("=SUM(Rate)", false)]
    public void Relative_references_are_those_with_a_row_or_column_not_fixed(string formula, bool expected)
    {
        var reference = FormulaParser.Parse(formula, Context).References.Single();

        Assert.Equal(expected, CallingCell.IsRelative(reference));
    }

    [Fact]
    public void Arguments_are_checked()
    {
        Assert.Throws<ArgumentNullException>(() => CallingCell.SubstituteRowAndColumn(null!, 1, 1));
        Assert.Throws<ArgumentNullException>(() => CallingCell.HasRelativeR1C1Indirect(null!));
        Assert.Throws<ArgumentNullException>(() => CallingCell.IsRelative(null!));
    }
}
