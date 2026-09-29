using System;
using ExcelModelingToolkit.Core.Formatting;
using Xunit;

namespace ExcelModelingToolkit.Core.Tests.Formatting;

public class OleColorTests
{
    /// <summary>
    /// The owner's Macabacus colors (docs/research/06) with their OLE values, computed as R + G*256 + B*65536.
    /// </summary>
    public static TheoryData<int, int, int, int, string> MacabacusPalette => new TheoryData<int, int, int, int, string>
    {
        // Font color cycle
        { 0, 0, 255, 16711680, "#0000FF" },
        { 0, 128, 0, 32768, "#008000" },
        { 128, 0, 128, 8388736, "#800080" },
        { 255, 0, 0, 255, "#FF0000" },
        { 255, 255, 255, 16777215, "#FFFFFF" },
        { 0, 0, 0, 0, "#000000" },

        // Fill color cycle (followed by "No fill")
        { 201, 218, 248, 16308937, "#C9DAF8" },
        { 210, 242, 255, 16773842, "#D2F2FF" },
        { 244, 204, 204, 13421812, "#F4CCCC" },
        { 252, 229, 205, 13493756, "#FCE5CD" },
        { 28, 69, 135, 8865052, "#1C4587" },

        // Border colors, recolor orange, default shading, chart series accents
        { 128, 128, 128, 8421504, "#808080" },
        { 204, 0, 0, 204, "#CC0000" },
        { 255, 102, 0, 26367, "#FF6600" },
        { 220, 220, 220, 14474460, "#DCDCDC" },
        { 91, 155, 213, 13998939, "#5B9BD5" },
        { 237, 125, 49, 3243501, "#ED7D31" },
    };

    [Theory]
    [MemberData(nameof(MacabacusPalette))]
    public void FromRgb_produces_excel_ole_value(int r, int g, int b, int ole, string hex)
    {
        var color = OleColor.FromRgb(r, g, b);

        Assert.Equal(ole, color.OleValue);
        Assert.Equal(r + (g * 256) + (b * 65536), color.OleValue);
        Assert.Equal(hex, color.ToString());
        Assert.False(color.IsNoFill);
    }

    [Theory]
    [MemberData(nameof(MacabacusPalette))]
    public void FromOle_recovers_components(int r, int g, int b, int ole, string hex)
    {
        var color = OleColor.FromOle(ole);

        Assert.Equal(r, color.R);
        Assert.Equal(g, color.G);
        Assert.Equal(b, color.B);
        Assert.Equal(OleColor.FromRgb(r, g, b), color);
        Assert.Equal(hex, color.ToString());
    }

    [Theory]
    [MemberData(nameof(MacabacusPalette))]
    public void Parse_round_trips_through_all_text_forms(int r, int g, int b, int ole, string hex)
    {
        var expected = OleColor.FromOle(ole);

        Assert.Equal(expected, OleColor.Parse(hex));
        Assert.Equal(expected, OleColor.Parse(hex.ToLowerInvariant()));
        Assert.Equal(expected, OleColor.Parse($"rgb({r},{g},{b})"));
        Assert.Equal(expected, OleColor.Parse(expected.ToString()));
    }

    [Theory]
    [InlineData(255, 0, 0, 255)] // vbRed
    [InlineData(0, 255, 0, 65280)] // vbGreen
    [InlineData(0, 0, 255, 16711680)] // vbBlue
    [InlineData(255, 255, 0, 65535)] // vbYellow
    [InlineData(255, 0, 255, 16711935)] // vbMagenta
    [InlineData(0, 255, 255, 16776960)] // vbCyan
    public void Matches_vba_color_constants(int r, int g, int b, int vbConstant)
    {
        Assert.Equal(vbConstant, OleColor.FromRgb(r, g, b).OleValue);
    }

    [Fact]
    public void Navy_is_bgr_ordered()
    {
        var navy = OleColor.FromRgb(28, 69, 135);

        Assert.Equal(0x87451C, navy.OleValue); // B=0x87, G=0x45, R=0x1C
        Assert.Equal(8865052, navy.OleValue);
    }

    [Theory]
    [InlineData("rgb(28,69,135)")]
    [InlineData("RGB(28,69,135)")]
    [InlineData("  rgb( 28 , 69 , 135 )  ")]
    [InlineData("rgb(028,069,135)")]
    [InlineData("#1C4587")]
    [InlineData("#1c4587")]
    [InlineData(" #1C4587 ")]
    public void Parse_accepts_rgb_and_hex_forms(string text)
    {
        Assert.Equal(OleColor.FromRgb(28, 69, 135), OleColor.Parse(text));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("navy")]
    [InlineData("1C4587")]
    [InlineData("#1C458")]
    [InlineData("#1C45877")]
    [InlineData("#GGGGGG")]
    [InlineData("#1C4587FF")]
    [InlineData("rgb(28,69)")]
    [InlineData("rgb(28,69,135,0)")]
    [InlineData("rgb(256,0,0)")]
    [InlineData("rgb(0,999,0)")]
    [InlineData("rgb(-1,0,0)")]
    [InlineData("rgb(1.5,0,0)")]
    [InlineData("rgb 28,69,135")]
    [InlineData("rgba(28,69,135,1)")]
    [InlineData("nofill")]
    public void Parse_rejects_invalid_text(string text)
    {
        var ex = Assert.Throws<FormatException>(() => OleColor.Parse(text));
        Assert.Contains("Invalid color", ex.Message);
    }

    [Fact]
    public void Parse_reports_out_of_range_component()
    {
        var ex = Assert.Throws<FormatException>(() => OleColor.Parse("rgb(300,0,0)"));
        Assert.Contains("0-255", ex.Message);
    }

    [Fact]
    public void Parse_rejects_null()
    {
        Assert.Throws<ArgumentNullException>(() => OleColor.Parse(null!));
    }

    [Theory]
    [InlineData(-1, 0, 0, "r")]
    [InlineData(256, 0, 0, "r")]
    [InlineData(0, -1, 0, "g")]
    [InlineData(0, 256, 0, "g")]
    [InlineData(0, 0, -1, "b")]
    [InlineData(0, 0, 256, "b")]
    public void FromRgb_rejects_components_outside_byte_range(int r, int g, int b, string param)
    {
        var ex = Assert.Throws<ArgumentOutOfRangeException>(() => OleColor.FromRgb(r, g, b));
        Assert.Equal(param, ex.ParamName);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(-4142)] // xlNone is not a color
    [InlineData(0x1000000)]
    [InlineData(int.MaxValue)]
    [InlineData(int.MinValue)]
    public void FromOle_rejects_values_outside_rgb_range(int value)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => OleColor.FromOle(value));
    }

    [Fact]
    public void FromOle_accepts_bounds()
    {
        Assert.Equal(OleColor.FromRgb(0, 0, 0), OleColor.FromOle(0));
        Assert.Equal(OleColor.FromRgb(255, 255, 255), OleColor.FromOle(OleColor.MaxOleValue));
    }

    [Fact]
    public void NoFill_is_a_distinct_sentinel()
    {
        var noFill = OleColor.NoFill;

        Assert.True(noFill.IsNoFill);
        Assert.Equal(OleColor.NoFill, noFill);
        Assert.True(noFill == OleColor.NoFill);
        Assert.NotEqual(OleColor.FromRgb(255, 255, 255), noFill);
        Assert.NotEqual(OleColor.FromRgb(0, 0, 0), noFill);
        Assert.NotEqual(default, noFill);
        Assert.Equal("none", noFill.ToString());
    }

    [Theory]
    [InlineData("none")]
    [InlineData("None")]
    [InlineData(" NONE ")]
    public void Parse_none_returns_NoFill(string text)
    {
        Assert.Equal(OleColor.NoFill, OleColor.Parse(text));
        Assert.Equal(OleColor.NoFill, OleColor.Parse(OleColor.NoFill.ToString()));
    }

    [Fact]
    public void NoFill_has_no_color_value()
    {
        var noFill = OleColor.NoFill;

        Assert.Throws<InvalidOperationException>(() => noFill.OleValue);
        Assert.Throws<InvalidOperationException>(() => noFill.R);
        Assert.Throws<InvalidOperationException>(() => noFill.G);
        Assert.Throws<InvalidOperationException>(() => noFill.B);
    }

    [Fact]
    public void Default_is_black()
    {
        OleColor color = default;

        Assert.False(color.IsNoFill);
        Assert.Equal(0, color.OleValue);
        Assert.Equal(OleColor.FromRgb(0, 0, 0), color);
    }

    [Fact]
    public void Equality_is_by_value()
    {
        var a = OleColor.FromRgb(28, 69, 135);
        var b = OleColor.Parse("#1C4587");
        var c = OleColor.FromRgb(28, 69, 136);

        Assert.True(a == b);
        Assert.False(a != b);
        Assert.True(a.Equals((object)b));
        Assert.Equal(a.GetHashCode(), b.GetHashCode());
        Assert.True(a != c);
        Assert.False(a.Equals(c));
        Assert.False(a.Equals("#1C4587"));
        Assert.False(a.Equals(null));
    }
}
