using Modelwright.Core.Formatting;
using Xunit;

namespace Modelwright.Core.Tests.Formatting;

public class CycleValueTests
{
    private static readonly CycleValue Black = CycleValue.FromColor(OleColor.FromRgb(0, 0, 0));

    [Fact]
    public void Automatic_is_known_but_neither_a_color_nor_a_number_format()
    {
        var automatic = CycleValue.Automatic;

        Assert.True(automatic.IsAutomatic);
        Assert.False(automatic.IsUnknown);
        Assert.False(automatic.IsColor);
        Assert.False(automatic.IsNumberFormat);
        Assert.Null(automatic.Color);
        Assert.Null(automatic.NumberFormat);
        Assert.Equal("automatic", automatic.ToString());
    }

    [Fact]
    public void Automatic_differs_from_explicit_black_and_unknown()
    {
        Assert.NotEqual(Black, CycleValue.Automatic);
        Assert.NotEqual(CycleValue.Unknown, CycleValue.Automatic);
        Assert.Equal(CycleValue.Automatic, CycleValue.Automatic);
        Assert.True(CycleValue.Automatic == CycleValue.Automatic);
        Assert.Equal(CycleValue.Automatic.GetHashCode(), CycleValue.Automatic.GetHashCode());
        Assert.False(Black.IsAutomatic);
        Assert.False(CycleValue.Unknown.IsAutomatic);
    }

    [Fact]
    public void Cycles_do_not_accept_automatic()
    {
        var cycle = new CycleDefinition("font", "Font", CycleKind.FontColor, new[] { new ColorItem("Black", OleColor.FromRgb(0, 0, 0)) });

        Assert.False(cycle.Accepts(CycleValue.Automatic));
        Assert.Equal(-1, new CycleEngine().IndexOf(cycle, CycleValue.Automatic));
    }
}
