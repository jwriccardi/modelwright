using System;
using System.Globalization;
using ExcelModelingToolkit.Core.Trace;
using Xunit;

namespace ExcelModelingToolkit.Core.Tests.Trace;

public class TraceValueTextTests
{
    private static readonly CultureInfo Us = CultureInfo.GetCultureInfo("en-US");
    private static readonly CultureInfo German = CultureInfo.GetCultureInfo("de-DE");

    [Theory]
    [InlineData(-2146826281, "#DIV/0!")]
    [InlineData(-2146826246, "#N/A")]
    [InlineData(-2146826259, "#NAME?")]
    [InlineData(-2146826288, "#NULL!")]
    [InlineData(-2146826252, "#NUM!")]
    [InlineData(-2146826265, "#REF!")]
    [InlineData(-2146826273, "#VALUE!")]
    [InlineData(-2146826243, "#SPILL!")]
    [InlineData(-2146826238, "#CALC!")]
    [InlineData(2042, "#N/A")]
    [InlineData(2045, "#SPILL!")]
    public void Excel_errors_read_as_excel_writes_them(int code, string expected)
    {
        Assert.Equal(expected, TraceValueText.ErrorText(code));
        Assert.Equal(expected, TraceValueText.FromValue(code, Us));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(2001)]
    [InlineData(-1)]
    public void Other_integers_are_not_errors(int code)
    {
        Assert.Null(TraceValueText.ErrorText(code));
        Assert.Equal(code.ToString(Us), TraceValueText.FromValue(code, Us));
    }

    [Fact]
    public void Values_are_formatted_like_general()
    {
        Assert.Equal(string.Empty, TraceValueText.FromValue(null, Us));
        Assert.Equal("1234.5", TraceValueText.FromValue(1234.5, Us));
        Assert.Equal("1234,5", TraceValueText.FromValue(1234.5, German));
        Assert.Equal("0.3", TraceValueText.FromValue(0.1 + 0.2, Us));
        Assert.Equal("TRUE", TraceValueText.FromValue(true, Us));
        Assert.Equal("FALSE", TraceValueText.FromValue(false, Us));
        Assert.Equal("Revenue", TraceValueText.FromValue("Revenue", Us));
        Assert.Equal("12", TraceValueText.FromValue(12m, Us));
    }

    [Fact]
    public void Text_is_one_line_and_cut_at_the_limit()
    {
        Assert.Equal("a↵b↵c d", TraceValueText.FromValue("a\r\nb\nc\td", Us));

        var cut = TraceValueText.FromValue(new string('x', 300), Us);
        Assert.Equal(TraceValueText.MaxLength, cut.Length);
        Assert.EndsWith("x…", cut);
        Assert.Equal(new string('x', 255), TraceValueText.FromValue(new string('x', 255), Us));
    }

    [Fact]
    public void A_surrogate_pair_is_never_split()
    {
        var text = new string('x', 253) + "\U0001F600" + "tail";

        var cut = TraceValueText.FromValue(text, Us);

        Assert.Equal(new string('x', 253) + "…", cut);
    }

    [Fact]
    public void Arrays_show_their_first_element_and_size()
    {
        var array = Array.CreateInstance(typeof(object), new[] { 3, 2 }, new[] { 1, 1 });
        array.SetValue(5.0, 1, 1);

        Assert.Equal("5 {3x2}", TraceValueText.FromValue(array, Us));
        Assert.Equal("{2}", TraceValueText.FromValue(new object?[] { null, 1.0 }, Us));
        Assert.Equal(string.Empty, TraceValueText.FromValue(new object[0], Us));
    }

    [Theory]
    [InlineData("5.0%", 0.05, "5.0%")]
    [InlineData("####", 123456.0, "123456")]
    [InlineData("", 7.0, "7")]
    [InlineData(null, 7.0, "7")]
    [InlineData("", null, "")]
    [InlineData("$1,234", 1234.0, "$1,234")]
    public void Display_prefers_the_formatted_text(string? text, object? value, string expected)
    {
        Assert.Equal(expected, TraceValueText.Display(text, value, Us));
    }

    [Fact]
    public void Ranges_show_the_first_cell_and_the_count()
    {
        Assert.Equal("5", TraceValueText.ForRange("5", 1, Us));
        Assert.Equal("5 (60 cells)", TraceValueText.ForRange("5", 60, Us));
        Assert.Equal("5 (1,048,576 cells)", TraceValueText.ForRange("5", 1048576, Us));
        Assert.Equal("5 (1.048.576 cells)", TraceValueText.ForRange("5", 1048576, German));
        Assert.Equal("(3 cells)", TraceValueText.ForRange(string.Empty, 3, Us));
        Assert.Throws<ArgumentNullException>(() => TraceValueText.ForRange(null!, 3, Us));
    }

    [Theory]
    [InlineData(false, false, false, false, null)]
    [InlineData(true, true, true, true, "hidden workbook")]
    [InlineData(false, true, false, false, "hidden sheet")]
    [InlineData(false, false, true, false, "hidden rows")]
    [InlineData(false, false, false, true, "hidden columns")]
    [InlineData(false, false, true, true, "hidden rows and columns")]
    public void Hidden_badges_name_the_first_reason(bool workbook, bool sheet, bool rows, bool columns, string? expected)
    {
        Assert.Equal(expected, TraceValueText.HiddenNote(workbook, sheet, rows, columns));
    }

    [Fact]
    public void Some_hidden_rows_or_columns_count_as_hidden()
    {
        Assert.Equal("hidden rows", TraceValueText.HiddenNote(false, false, null, false));
        Assert.Equal("hidden columns", TraceValueText.HiddenNote(false, false, false, null));
    }

    [Fact]
    public void A_hidden_note_marks_the_item()
    {
        var hidden = new PrecedentItem(PrecedentKind.Cell, "Secret!A1", "Model.xlsx", "Secret", "A1", hiddenNote: "hidden sheet");
        var visible = new PrecedentItem(PrecedentKind.Cell, "Calc!A1", "Model.xlsx", "Calc", "A1");

        Assert.True(hidden.IsHidden);
        Assert.Equal("hidden sheet", hidden.HiddenNote);
        Assert.False(visible.IsHidden);
        Assert.Null(visible.HiddenNote);
    }
}
