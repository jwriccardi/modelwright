using System;
using System.Linq;
using ExcelModelingToolkit.Core.Trace;
using Xunit;

namespace ExcelModelingToolkit.Core.Tests.Trace;

public class FormulaColoringTests
{
    private static readonly FormulaContext Context = new FormulaContext("Model.xlsx", "Calc");

    private static string Describe(string formula) =>
        string.Join(" | ", FormulaColoring.Segment(FormulaParser.Parse(formula, Context))
            .Select(s => s.IsReference ? s.Text + "#" + s.ColorIndex : s.Text));

    [Fact]
    public void References_get_palette_colors_in_order_and_the_rest_is_plain()
    {
        Assert.Equal("= | A1#0 | + | Sheet2!B5#1 | * | Rate#2", Describe("=A1+Sheet2!B5*Rate"));
    }

    [Fact]
    public void A_reference_written_again_keeps_its_color_ignoring_case_and_dollars()
    {
        Assert.Equal("= | Data!A1#0 | + | B1#1 | + | $A$1#2 | + | DATA!$A$1#0", Describe("=Data!A1+B1+$A$1+DATA!$A$1"));
    }

    [Fact]
    public void Segments_cover_the_whole_formula_in_order()
    {
        const string formula = "=SUM(Data!A1:A10)/COUNT([Book.xlsx]Data!C3,Sales[Amount])";

        var segments = FormulaColoring.Segment(FormulaParser.Parse(formula, Context));

        Assert.Equal(formula, string.Concat(segments.Select(s => s.Text)));
        var position = 0;
        foreach (var segment in segments)
        {
            Assert.Equal(position, segment.Start);
            position += segment.Text.Length;
        }

        Assert.Equal(new[] { "Data!A1:A10", "[Book.xlsx]Data!C3", "Sales[Amount]" }, segments.Where(s => s.IsReference).Select(s => s.Text));
    }

    [Fact]
    public void Colors_wrap_round_the_palette()
    {
        var formula = "=" + string.Join("+", Enumerable.Range(1, FormulaColoring.Palette.Count + 1).Select(i => "A" + i));

        var colors = FormulaColoring.Segment(FormulaParser.Parse(formula, Context)).Where(s => s.IsReference).Select(s => s.ColorIndex).ToArray();

        Assert.Equal(Enumerable.Range(0, FormulaColoring.Palette.Count).Concat(new[] { 0 }), colors);
    }

    [Fact]
    public void Local_names_and_ref_errors_are_plain()
    {
        Assert.Equal("=LET(x,1,x+ | A1#0 | )+#REF!", Describe("=LET(x,1,x+A1)+#REF!"));
    }

    [Fact]
    public void An_unparsed_formula_is_one_plain_segment()
    {
        var segment = Assert.Single(FormulaColoring.Segment(FormulaParser.Parse("=SUM(", Context)));

        Assert.Equal("=SUM(", segment.Text);
        Assert.False(segment.IsReference);
        Assert.Equal(-1, segment.ColorIndex);
        Assert.Equal("=SUM(", segment.ToString());
    }

    [Fact]
    public void The_palette_is_eight_readable_colors()
    {
        Assert.Equal(8, FormulaColoring.Palette.Count);
        Assert.Equal(8, FormulaColoring.Palette.Distinct().Count());
        Assert.All(FormulaColoring.Palette, color => Assert.InRange(color, 0, 0xFFFFFF));
        Assert.All(FormulaColoring.Palette, color =>
        {
            // Dark enough for text on white: relative luminance well below the background's.
            var r = (color >> 16) & 0xFF;
            var g = (color >> 8) & 0xFF;
            var b = color & 0xFF;
            Assert.True(0.2126 * r + 0.7152 * g + 0.0722 * b < 140, color.ToString("X6"));
        });
    }

    [Fact]
    public void Arguments_are_checked()
    {
        Assert.Throws<ArgumentNullException>(() => FormulaColoring.Segment(null!));
        Assert.Throws<ArgumentNullException>(() => new FormulaSegment(null!, 0, -1));
    }
}
