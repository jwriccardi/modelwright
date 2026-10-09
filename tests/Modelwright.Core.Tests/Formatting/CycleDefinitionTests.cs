using System;
using Modelwright.Core.Formatting;
using Xunit;

namespace Modelwright.Core.Tests.Formatting;

public class CycleDefinitionTests
{
    private static readonly ColorItem Blue = new ColorItem("Blue", OleColor.FromRgb(0, 0, 255));

    [Fact]
    public void Valid_cycles_have_no_problems()
    {
        Assert.Empty(new CycleDefinition("N", "Number", CycleKind.NumberFormat, new CycleItem[] { new NumberFormatItem("Zero", "0") }).Validate());
        Assert.Empty(new CycleDefinition("F", "Font", CycleKind.FontColor, new CycleItem[] { Blue }).Validate());
        Assert.Empty(new CycleDefinition("K", "Fill", CycleKind.FillColor, new CycleItem[] { Blue, new ColorItem("None", OleColor.NoFill) }).Validate());
    }

    [Fact]
    public void Reports_an_empty_cycle()
    {
        var problems = new CycleDefinition("Empty", "Empty", CycleKind.FontColor, Array.Empty<CycleItem>()).Validate();

        var problem = Assert.Single(problems);
        Assert.Contains("cycle 'Empty'", problem);
        Assert.Contains("no items", problem);
    }

    [Fact]
    public void Reports_blank_id_display_name_item_name_and_code()
    {
        var cycle = new CycleDefinition(" ", "", CycleKind.NumberFormat, new CycleItem[] { new NumberFormatItem("", " ") });

        var problems = cycle.Validate();

        Assert.Equal(4, problems.Count);
        Assert.Contains(problems, p => p.Contains("id is blank"));
        Assert.Contains(problems, p => p.Contains("display name is blank"));
        Assert.Contains(problems, p => p.Contains("item 1: the name is blank"));
        Assert.Contains(problems, p => p.Contains("item 1: the number format code is blank"));
    }

    [Fact]
    public void Reports_items_of_the_wrong_kind()
    {
        var numberWithColor = new CycleDefinition("N", "N", CycleKind.NumberFormat, new CycleItem[] { Blue });
        var fontWithFormat = new CycleDefinition("F", "F", CycleKind.FontColor, new CycleItem[] { new NumberFormatItem("Zero", "0") });
        var fillWithFormat = new CycleDefinition("K", "K", CycleKind.FillColor, new CycleItem[] { new NumberFormatItem("Zero", "0") });

        Assert.Contains("color item does not belong in a number format cycle", Assert.Single(numberWithColor.Validate()));
        Assert.Contains("number format item does not belong in a font color cycle", Assert.Single(fontWithFormat.Validate()));
        Assert.Contains("number format item does not belong in a fill color cycle", Assert.Single(fillWithFormat.Validate()));
    }

    [Fact]
    public void Reports_no_fill_in_a_font_cycle()
    {
        var cycle = new CycleDefinition("F", "Font", CycleKind.FontColor, new CycleItem[] { Blue, new ColorItem("None", OleColor.NoFill) });

        var problem = Assert.Single(cycle.Validate());

        Assert.Contains("item 2", problem);
        Assert.Contains("only valid in a fill color cycle", problem);
    }

    [Fact]
    public void Reports_a_number_format_code_longer_than_excel_accepts()
    {
        var longest = new CycleDefinition("N", "N", CycleKind.NumberFormat, new CycleItem[] { new NumberFormatItem("Max", new string('0', 255)) });
        var tooLong = new CycleDefinition("N", "N", CycleKind.NumberFormat, new CycleItem[]
        {
            new NumberFormatItem("Zero", "0"),
            new NumberFormatItem("Long", new string('#', 256)),
        });

        Assert.Equal(255, NumberFormatItem.MaxCodeLength);
        Assert.Empty(longest.Validate());
        Assert.Equal(
            "cycle 'N', item 2: the number format code is 256 characters; Excel accepts at most 255.",
            Assert.Single(tooLong.Validate()));
    }

    [Fact]
    public void HasSameItems_compares_id_kind_and_item_values_in_order()
    {
        var cycle = new CycleDefinition("N", "Number", CycleKind.NumberFormat, new CycleItem[]
        {
            new NumberFormatItem("Zero", "0"),
            new NumberFormatItem("One", "0.0"),
        });

        // Names, display name and provisional do not matter.
        Assert.True(cycle.HasSameItems(new CycleDefinition("N", "Renamed", CycleKind.NumberFormat, new CycleItem[]
        {
            new NumberFormatItem("Nought", "0"),
            new NumberFormatItem("Tenths", "0.0"),
        }, provisional: true)));
        Assert.True(cycle.HasSameItems(cycle));

        Assert.False(cycle.HasSameItems(null));
        Assert.False(cycle.HasSameItems(new CycleDefinition("n", "Number", CycleKind.NumberFormat, cycle.Items)));
        Assert.False(cycle.HasSameItems(new CycleDefinition("N", "Number", CycleKind.NumberFormat, new CycleItem[]
        {
            new NumberFormatItem("One", "0.0"),
            new NumberFormatItem("Zero", "0"),
        })));
        Assert.False(cycle.HasSameItems(new CycleDefinition("N", "Number", CycleKind.NumberFormat, new CycleItem[] { new NumberFormatItem("Zero", "0") })));
        Assert.False(cycle.HasSameItems(new CycleDefinition("N", "Number", CycleKind.NumberFormat, new CycleItem[]
        {
            new NumberFormatItem("Zero", "0"),
            new NumberFormatItem("One", "0.0_)"),
        })));

        var font = new CycleDefinition("F", "Font", CycleKind.FontColor, new CycleItem[] { Blue });
        Assert.False(font.HasSameItems(new CycleDefinition("F", "Font", CycleKind.FillColor, new CycleItem[] { Blue })));
        Assert.False(font.HasSameItems(new CycleDefinition("F", "Font", CycleKind.FontColor, new CycleItem[] { new ColorItem("Blue", OleColor.FromRgb(0, 0, 254)) })));
    }

    [Fact]
    public void Constructor_rejects_nulls_and_copies_items()
    {
        var items = new CycleItem[] { Blue };
        var cycle = new CycleDefinition("F", "Font", CycleKind.FontColor, items, provisional: true);
        items[0] = new ColorItem("Other", OleColor.FromRgb(1, 2, 3));

        Assert.Same(Blue, cycle.Items[0]);
        Assert.True(cycle.Provisional);
        Assert.Equal("F (FontColor, 1 items)", cycle.ToString());

        Assert.Throws<ArgumentNullException>(() => new CycleDefinition(null!, "x", CycleKind.FontColor, items));
        Assert.Throws<ArgumentNullException>(() => new CycleDefinition("x", null!, CycleKind.FontColor, items));
        Assert.Throws<ArgumentNullException>(() => new CycleDefinition("x", "x", CycleKind.FontColor, null!));
        Assert.Throws<ArgumentException>(() => new CycleDefinition("x", "x", CycleKind.FontColor, new CycleItem[] { null! }));
    }

    [Fact]
    public void Accepts_values_of_its_kind_and_unknown()
    {
        var number = new CycleDefinition("N", "N", CycleKind.NumberFormat, new CycleItem[] { new NumberFormatItem("Zero", "0") });
        var fill = new CycleDefinition("K", "K", CycleKind.FillColor, new CycleItem[] { Blue });

        Assert.True(number.Accepts(CycleValue.FromNumberFormat("0")));
        Assert.True(number.Accepts(CycleValue.Unknown));
        Assert.False(number.Accepts(CycleValue.FromColor(OleColor.NoFill)));
        Assert.True(fill.Accepts(CycleValue.FromColor(OleColor.NoFill)));
        Assert.False(fill.Accepts(CycleValue.FromNumberFormat("0")));
    }

    [Fact]
    public void Items_expose_their_values()
    {
        var format = new NumberFormatItem("Zero", "0");

        Assert.Equal(CycleValue.FromNumberFormat("0"), format.Value);
        Assert.Equal(CycleValue.FromColor(OleColor.FromRgb(0, 0, 255)), Blue.Value);
        Assert.Equal("Zero: 0", format.ToString());
        Assert.Equal("Blue: #0000FF", Blue.ToString());
        Assert.Throws<ArgumentNullException>(() => new NumberFormatItem(null!, "0"));
        Assert.Throws<ArgumentNullException>(() => new NumberFormatItem("Zero", null!));
        Assert.Throws<ArgumentNullException>(() => new ColorItem(null!, OleColor.NoFill));
    }

    [Fact]
    public void CycleValue_kinds_and_equality()
    {
        var format = CycleValue.FromNumberFormat("0");
        var black = CycleValue.FromColor(OleColor.FromRgb(0, 0, 0));
        var noFill = CycleValue.FromColor(OleColor.NoFill);

        Assert.True(CycleValue.Unknown.IsUnknown);
        Assert.Equal(default, CycleValue.Unknown);
        Assert.Null(CycleValue.Unknown.NumberFormat);
        Assert.Null(CycleValue.Unknown.Color);
        Assert.Equal("(unknown)", CycleValue.Unknown.ToString());

        Assert.True(format.IsNumberFormat);
        Assert.False(format.IsColor);
        Assert.False(format.IsUnknown);
        Assert.Equal("0", format.NumberFormat);
        Assert.Null(format.Color);
        Assert.Equal("0", format.ToString());

        Assert.True(black.IsColor);
        Assert.False(black.IsUnknown);
        Assert.Equal(OleColor.FromRgb(0, 0, 0), black.Color);
        Assert.Equal("#000000", black.ToString());
        Assert.Equal("none", noFill.ToString());

        // Black (OLE 0) is neither unknown nor no fill, and a "0" format is not black.
        Assert.NotEqual(CycleValue.Unknown, black);
        Assert.NotEqual(noFill, black);
        Assert.NotEqual(format, black);
        Assert.True(format == CycleValue.FromNumberFormat("0"));
        Assert.True(format != CycleValue.FromNumberFormat("0.0"));
        Assert.True(format.Equals((object)CycleValue.FromNumberFormat("0")));
        Assert.False(format.Equals("0"));
        Assert.Equal(format.GetHashCode(), CycleValue.FromNumberFormat("0").GetHashCode());
        Assert.Equal(noFill.GetHashCode(), CycleValue.FromColor(OleColor.NoFill).GetHashCode());
        Assert.Equal(0, CycleValue.Unknown.GetHashCode());
        Assert.Throws<ArgumentNullException>(() => CycleValue.FromNumberFormat(null!));
    }
}
