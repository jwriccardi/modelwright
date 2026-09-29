using System;
using ExcelModelingToolkit.Core.Formatting;
using Xunit;

namespace ExcelModelingToolkit.Core.Tests.Formatting;

public class CycleEngineTests
{
    private const string SelectionA = "Book1.xlsx|Sheet1|$A$1:$B$4";
    private const string SelectionB = "Book1.xlsx|Sheet1|$C$1";

    private static readonly CycleDefinition Formats = new CycleDefinition("Fmt", "Formats", CycleKind.NumberFormat, new CycleItem[]
    {
        new NumberFormatItem("Zero", "0"),
        new NumberFormatItem("One", "0.0"),
        new NumberFormatItem("Two", "0.00"),
    });

    private static readonly OleColor Blue = OleColor.FromRgb(0, 0, 255);
    private static readonly OleColor Black = OleColor.FromRgb(0, 0, 0);
    private static readonly OleColor White = OleColor.FromRgb(255, 255, 255);
    private static readonly OleColor Navy = OleColor.FromRgb(28, 69, 135);

    private static readonly CycleDefinition Fills = new CycleDefinition("Fill", "Fill", CycleKind.FillColor, new CycleItem[]
    {
        new ColorItem("Light Blue", OleColor.FromRgb(201, 218, 248)),
        new ColorItem("Navy", Navy),
        new ColorItem("No Fill", OleColor.NoFill),
    });

    private static readonly CycleDefinition Fonts = new CycleDefinition("Font", "Font", CycleKind.FontColor, new CycleItem[]
    {
        new ColorItem("Blue", Blue),
        new ColorItem("White", White),
        new ColorItem("Black", Black),
    });

    private readonly CycleEngine _engine = new CycleEngine();

    private static CycleValue Format(string code) => CycleValue.FromNumberFormat(code);

    private static CycleValue Color(OleColor color) => CycleValue.FromColor(color);

    [Fact]
    public void No_match_applies_the_first_item()
    {
        var step = _engine.Next(Formats, Format("General"), SelectionA, null);

        Assert.Equal(0, step.Index);
        Assert.Equal("Zero", step.Item.Name);
        Assert.Equal(CycleStepReason.NoMatch, step.Reason);
    }

    [Theory]
    [InlineData("0", 1)]
    [InlineData("0.0", 2)]
    [InlineData("0.00", 0)] // wraps
    public void Match_on_item_k_applies_k_plus_1(string current, int expected)
    {
        var step = _engine.Next(Formats, Format(current), SelectionA, null);

        Assert.Equal(expected, step.Index);
        Assert.Equal(CycleStepReason.Matched, step.Reason);
    }

    [Fact]
    public void Number_format_match_is_ordinal()
    {
        var cycle = new CycleDefinition("D", "Date", CycleKind.NumberFormat, new CycleItem[]
        {
            new NumberFormatItem("ISO", "yyyy-mm-dd"),
            new NumberFormatItem("US", "mm-dd-yyyy"),
        });

        Assert.Equal(CycleStepReason.NoMatch, _engine.Next(cycle, Format("YYYY-MM-DD"), SelectionA, null).Reason);
        Assert.Equal(1, _engine.Next(cycle, Format("yyyy-mm-dd"), SelectionA, null).Index);
    }

    [Fact]
    public void Repeated_presses_on_the_same_selection_walk_the_cycle_and_wrap()
    {
        var current = Format("General");
        CycleState? state = null;
        var applied = new int[7];
        for (var press = 0; press < applied.Length; press++)
        {
            var step = _engine.Next(Formats, current, SelectionA, state);
            applied[press] = step.Index;
            state = step.State;
            current = step.Item.Value; // what Excel now reports
        }

        Assert.Equal(new[] { 0, 1, 2, 0, 1, 2, 0 }, applied);
    }

    [Fact]
    public void Returned_state_records_the_press()
    {
        var step = _engine.Next(Formats, Format("0"), SelectionA, null);

        Assert.Equal("Fmt", step.State.CycleId);
        Assert.Equal(SelectionA, step.State.SelectionKey);
        Assert.Equal(1, step.State.LastIndex);
        Assert.Equal(Format("0.0"), step.State.LastAppliedValue);
    }

    [Fact]
    public void Same_selection_advances_by_position_even_when_items_repeat()
    {
        // Matching alone would loop on the first "0" forever; the same-selection rule reaches the later items.
        var cycle = new CycleDefinition("Dup", "Dup", CycleKind.NumberFormat, new CycleItem[]
        {
            new NumberFormatItem("A", "0"),
            new NumberFormatItem("B", "0.0"),
            new NumberFormatItem("C", "0"),
            new NumberFormatItem("D", "0.00"),
        });

        var first = _engine.Next(cycle, Format("0.0"), SelectionA, null); // matched B -> C
        Assert.Equal(2, first.Index);

        var second = _engine.Next(cycle, Format("0"), SelectionA, first.State);
        Assert.Equal(3, second.Index);
        Assert.Equal(CycleStepReason.SameSelection, second.Reason);

        var fresh = _engine.Next(cycle, Format("0"), SelectionA, null); // no state: matches A -> B
        Assert.Equal(1, fresh.Index);
    }

    [Fact]
    public void Manual_change_on_the_same_selection_falls_back_to_matching()
    {
        var first = _engine.Next(Formats, Format("General"), SelectionA, null);
        Assert.Equal(0, first.Index);

        // The user typed a different format by hand; it matches "Two".
        var second = _engine.Next(Formats, Format("0.00"), SelectionA, first.State);
        Assert.Equal(0, second.Index);
        Assert.Equal(CycleStepReason.Matched, second.Reason);

        // An unrecognized manual format restarts the cycle.
        var third = _engine.Next(Formats, Format("#,##0"), SelectionA, second.State);
        Assert.Equal(0, third.Index);
        Assert.Equal(CycleStepReason.NoMatch, third.Reason);
    }

    [Fact]
    public void A_different_selection_ignores_the_previous_state()
    {
        var first = _engine.Next(Formats, Format("0"), SelectionA, null);

        var other = _engine.Next(Formats, Format("0"), SelectionB, first.State);

        Assert.Equal(1, other.Index);
        Assert.Equal(CycleStepReason.Matched, other.Reason);
    }

    [Fact]
    public void State_of_another_cycle_is_ignored()
    {
        var other = new CycleDefinition("Other", "Other", CycleKind.NumberFormat, Formats.Items);
        var first = _engine.Next(other, Format("0"), SelectionA, null);

        var step = _engine.Next(Formats, Format("0.0"), SelectionA, first.State);

        Assert.Equal(CycleStepReason.Matched, step.Reason);
        Assert.Equal(2, step.Index);
    }

    [Fact]
    public void State_with_an_index_past_the_end_is_ignored()
    {
        var stale = new CycleState("Fmt", SelectionA, 5, Format("0"));

        var step = _engine.Next(Formats, Format("0"), SelectionA, stale);

        Assert.Equal(CycleStepReason.Matched, step.Reason);
        Assert.Equal(1, step.Index);
    }

    [Fact]
    public void Unknown_value_without_state_applies_the_first_item()
    {
        var step = _engine.Next(Fonts, CycleValue.Unknown, SelectionA, null);

        Assert.Equal(0, step.Index);
        Assert.Equal(CycleStepReason.NoMatch, step.Reason);
    }

    [Fact]
    public void Unknown_value_on_the_same_selection_keeps_advancing()
    {
        var first = _engine.Next(Fonts, Color(Black), SelectionA, null); // Black -> Blue (wrap)
        Assert.Equal(0, first.Index);

        var second = _engine.Next(Fonts, CycleValue.Unknown, SelectionA, first.State);

        Assert.Equal(1, second.Index);
        Assert.Equal(CycleStepReason.SameSelection, second.Reason);
    }

    [Fact]
    public void Unknown_value_on_another_selection_applies_the_first_item()
    {
        var first = _engine.Next(Fonts, Color(Blue), SelectionA, null);

        var step = _engine.Next(Fonts, CycleValue.Unknown, SelectionB, first.State);

        Assert.Equal(0, step.Index);
    }

    [Fact]
    public void Read_back_alias_is_recognized_as_the_item()
    {
        // Excel reports "0.0" back as "0.0_)" (a hypothetical normalization).
        var step = _engine.Next(Formats, Format("General"), SelectionA, null); // applies "0"
        var afterZero = _engine.RecordReadBack(Formats, step.State, "0");
        var next = _engine.Next(Formats, Format("0"), SelectionA, afterZero); // applies "0.0"
        Assert.Equal(1, next.Index);

        var state = _engine.RecordReadBack(Formats, next.State, "0.0_)");
        Assert.Equal(Format("0.0_)"), state.LastAppliedValue);
        Assert.Equal(1, state.LastIndex);

        // Same selection: continues from the alias.
        Assert.Equal(2, _engine.Next(Formats, Format("0.0_)"), SelectionA, state).Index);

        // Another selection with that normalized format: the alias matches item "One".
        var elsewhere = _engine.Next(Formats, Format("0.0_)"), SelectionB, null);
        Assert.Equal(2, elsewhere.Index);
        Assert.Equal(CycleStepReason.Matched, elsewhere.Reason);
        Assert.Equal(1, _engine.IndexOf(Formats, Format("0.0_)")));
    }

    [Fact]
    public void Same_selection_accepts_an_alias_of_the_last_item()
    {
        var step = _engine.Next(Formats, Format("0"), SelectionB, null); // applies "0.0"
        _engine.RecordReadBack(Formats, step.State, "0.0 ");

        // The state still holds the item's own code, but the cell reads back as the alias.
        var next = _engine.Next(Formats, Format("0.0 "), SelectionB, step.State);

        Assert.Equal(CycleStepReason.SameSelection, next.Reason);
        Assert.Equal(2, next.Index);
    }

    [Fact]
    public void Exact_code_beats_an_alias_of_an_earlier_item()
    {
        var step = _engine.Next(Formats, Format("General"), SelectionA, null); // applies "0"
        _engine.RecordReadBack(Formats, step.State, "0.00"); // pathological: "0" read back as item Two's code

        Assert.Equal(2, _engine.IndexOf(Formats, Format("0.00")));
    }

    [Fact]
    public void Aliases_are_per_cycle_and_per_engine()
    {
        var copy = new CycleDefinition("Copy", "Copy", CycleKind.NumberFormat, Formats.Items);
        var step = _engine.Next(Formats, Format("General"), SelectionA, null);
        _engine.RecordReadBack(Formats, step.State, "zero!");

        Assert.Equal(0, _engine.IndexOf(Formats, Format("zero!")));
        Assert.Equal(-1, _engine.IndexOf(copy, Format("zero!")));
        Assert.Equal(-1, new CycleEngine().IndexOf(Formats, Format("zero!")));
    }

    [Fact]
    public void Read_back_equal_to_the_code_adds_no_alias()
    {
        var step = _engine.Next(Formats, Format("General"), SelectionA, null);

        var state = _engine.RecordReadBack(Formats, step.State, "0");

        Assert.Equal(Format("0"), state.LastAppliedValue);
        Assert.Equal(-1, _engine.IndexOf(Formats, Format("General")));
    }

    [Fact]
    public void RecordReadBack_rejects_bad_arguments()
    {
        var step = _engine.Next(Formats, Format("General"), SelectionA, null);
        var fontStep = _engine.Next(Fonts, Color(Blue), SelectionA, null);

        Assert.Throws<ArgumentNullException>(() => _engine.RecordReadBack(null!, step.State, "0"));
        Assert.Throws<ArgumentNullException>(() => _engine.RecordReadBack(Formats, null!, "0"));
        Assert.Throws<ArgumentNullException>(() => _engine.RecordReadBack(Formats, step.State, null!));
        Assert.Throws<ArgumentException>(() => _engine.RecordReadBack(Fonts, fontStep.State, "0"));
        Assert.Throws<ArgumentException>(() => _engine.RecordReadBack(Formats, new CycleState("Other", SelectionA, 0, Format("0")), "0"));
        Assert.Throws<ArgumentException>(() => _engine.RecordReadBack(Formats, new CycleState("Fmt", SelectionA, 3, Format("0")), "0"));
    }

    [Fact]
    public void Number_format_read_back_value_learns_an_alias()
    {
        var step = _engine.Next(Formats, Format("General"), SelectionA, null); // applies "0"

        var state = _engine.RecordReadBack(Formats, step.State, Format("0_)"));

        Assert.Equal(Format("0_)"), state.LastAppliedValue);
        Assert.Equal(0, _engine.IndexOf(Formats, Format("0_)")));
    }

    [Fact]
    public void Color_read_back_becomes_the_last_applied_value()
    {
        // Excel maps the applied blue to a nearby color (e.g. a legacy palette); the next press still continues.
        var nearBlue = OleColor.FromRgb(0, 0, 250);
        var step = _engine.Next(Fonts, Color(Black), SelectionA, null); // applies Blue

        var state = _engine.RecordReadBack(Fonts, step.State, Color(nearBlue));
        var next = _engine.Next(Fonts, Color(nearBlue), SelectionA, state);

        Assert.Equal(Color(nearBlue), state.LastAppliedValue);
        Assert.Equal(0, state.LastIndex);
        Assert.Equal(CycleStepReason.SameSelection, next.Reason);
        Assert.Equal(1, next.Index);
        Assert.Equal(-1, _engine.IndexOf(Fonts, Color(nearBlue))); // colors learn no aliases
    }

    [Fact]
    public void Fill_read_back_of_no_fill_is_recorded()
    {
        var step = _engine.Next(Fills, Color(Navy), SelectionA, null); // applies No Fill

        var state = _engine.RecordReadBack(Fills, step.State, Color(OleColor.NoFill));

        Assert.Equal(Color(OleColor.NoFill), state.LastAppliedValue);
        Assert.Equal(2, state.LastIndex);
    }

    [Fact]
    public void Unknown_read_back_keeps_the_state()
    {
        var fontStep = _engine.Next(Fonts, Color(Black), SelectionA, null);
        var formatStep = _engine.Next(Formats, Format("0"), SelectionA, null);

        Assert.Same(fontStep.State, _engine.RecordReadBack(Fonts, fontStep.State, CycleValue.Unknown));
        Assert.Same(formatStep.State, _engine.RecordReadBack(Formats, formatStep.State, CycleValue.Unknown));
    }

    [Fact]
    public void Read_back_value_of_the_wrong_kind_is_rejected()
    {
        var fontStep = _engine.Next(Fonts, Color(Blue), SelectionA, null);
        var formatStep = _engine.Next(Formats, Format("0"), SelectionA, null);

        Assert.Throws<ArgumentException>(() => _engine.RecordReadBack(Fonts, fontStep.State, Format("0")));
        Assert.Throws<ArgumentException>(() => _engine.RecordReadBack(Formats, formatStep.State, Color(Blue)));
        Assert.Throws<ArgumentException>(() => _engine.RecordReadBack(Fonts, formatStep.State, Color(Blue)));
        Assert.Throws<ArgumentNullException>(() => _engine.RecordReadBack(null!, fontStep.State, Color(Blue)));
        Assert.Throws<ArgumentNullException>(() => _engine.RecordReadBack(Fonts, null!, Color(Blue)));
    }

    [Fact]
    public void RetainAliases_keeps_aliases_of_codes_still_in_the_cycle()
    {
        var zero = _engine.Next(Formats, Format("General"), SelectionA, null); // applies "0"
        _engine.RecordReadBack(Formats, zero.State, "zero!");
        var one = _engine.Next(Formats, Format("0"), SelectionB, null); // applies "0.0"
        _engine.RecordReadBack(Formats, one.State, "one!");

        // "0" is still there (renamed and moved); "0.0" is gone.
        var edited = new CycleDefinition("Fmt", "Formats", CycleKind.NumberFormat, new CycleItem[]
        {
            new NumberFormatItem("Two", "0.00"),
            new NumberFormatItem("Nought", "0"),
        });
        _engine.RetainAliases(new[] { edited });

        Assert.Equal(1, _engine.IndexOf(edited, Format("zero!")));
        Assert.Equal(-1, _engine.IndexOf(Formats, Format("one!")));
    }

    [Fact]
    public void RetainAliases_keeps_everything_when_the_cycles_are_unchanged()
    {
        var step = _engine.Next(Formats, Format("General"), SelectionA, null);
        _engine.RecordReadBack(Formats, step.State, "zero!");

        _engine.RetainAliases(new[] { Formats, Fonts });

        Assert.Equal(0, _engine.IndexOf(Formats, Format("zero!")));
    }

    [Fact]
    public void RetainAliases_drops_removed_cycles_and_cycles_that_changed_kind()
    {
        var step = _engine.Next(Formats, Format("General"), SelectionA, null);
        _engine.RecordReadBack(Formats, step.State, "zero!");

        _engine.RetainAliases(new[] { Fonts });
        Assert.Equal(-1, _engine.IndexOf(Formats, Format("zero!")));

        _engine.RecordReadBack(Formats, step.State, "zero!");
        _engine.RetainAliases(new[] { new CycleDefinition("Fmt", "Now a font", CycleKind.FontColor, new CycleItem[] { new ColorItem("Blue", Blue) }) });
        Assert.Equal(-1, _engine.IndexOf(Formats, Format("zero!")));

        Assert.Throws<ArgumentNullException>(() => _engine.RetainAliases(null!));
        Assert.Throws<ArgumentException>(() => _engine.RetainAliases(new CycleDefinition[] { null! }));
    }

    [Fact]
    public void Fill_no_fill_is_matched_and_wraps_to_the_first_item()
    {
        var step = _engine.Next(Fills, Color(OleColor.NoFill), SelectionA, null);

        Assert.Equal(0, step.Index);
        Assert.Equal(CycleStepReason.Matched, step.Reason);
    }

    [Fact]
    public void Fill_navy_advances_to_no_fill()
    {
        var step = _engine.Next(Fills, Color(Navy), SelectionA, null);

        Assert.Equal(2, step.Index);
        Assert.True(((ColorItem)step.Item).Color.IsNoFill);
        Assert.Equal(Color(OleColor.NoFill), step.State.LastAppliedValue);
    }

    [Fact]
    public void White_fill_is_not_no_fill()
    {
        // Excel reports a no-fill cell's Interior.Color as white; the adapter maps it to NoFill via the pattern.
        // A genuine white fill must not be mistaken for the "No Fill" item.
        var step = _engine.Next(Fills, Color(White), SelectionA, null);

        Assert.Equal(CycleStepReason.NoMatch, step.Reason);
        Assert.Equal(0, step.Index);
    }

    [Fact]
    public void Fill_cycle_walks_through_no_fill_on_one_selection()
    {
        var current = Color(Navy);
        CycleState? state = null;
        var names = new string[4];
        for (var press = 0; press < names.Length; press++)
        {
            var step = _engine.Next(Fills, current, SelectionA, state);
            names[press] = step.Item.Name;
            state = step.State;
            current = step.Item.Value;
        }

        Assert.Equal(new[] { "No Fill", "Light Blue", "Navy", "No Fill" }, names);
    }

    [Fact]
    public void Font_black_and_white_are_distinct()
    {
        Assert.Equal(1, _engine.Next(Fonts, Color(Blue), SelectionA, null).Index);
        Assert.Equal(2, _engine.Next(Fonts, Color(White), SelectionA, null).Index);
        Assert.Equal(0, _engine.Next(Fonts, Color(Black), SelectionA, null).Index);
        Assert.Equal(0, _engine.Next(Fonts, Color(OleColor.FromRgb(1, 0, 0)), SelectionA, null).Index);
    }

    [Fact]
    public void Blue_black_toggle_alternates()
    {
        var toggle = new CycleDefinition("BB", "Blue/Black", CycleKind.FontColor, new CycleItem[]
        {
            new ColorItem("Blue", Blue),
            new ColorItem("Black", Black),
        });

        Assert.Equal("Black", _engine.Next(toggle, Color(Blue), SelectionA, null).Item.Name);
        Assert.Equal("Blue", _engine.Next(toggle, Color(Black), SelectionA, null).Item.Name);
        Assert.Equal("Blue", _engine.Next(toggle, Color(OleColor.FromRgb(0, 128, 0)), SelectionA, null).Item.Name);
        Assert.Equal("Blue", _engine.Next(toggle, CycleValue.Unknown, SelectionA, null).Item.Name);
    }

    [Fact]
    public void Single_item_cycle_always_applies_it()
    {
        var one = new CycleDefinition("One", "One", CycleKind.FontColor, new CycleItem[] { new ColorItem("Blue", Blue) });

        var first = _engine.Next(one, Color(Blue), SelectionA, null);
        var second = _engine.Next(one, Color(Blue), SelectionA, first.State);

        Assert.Equal(0, first.Index);
        Assert.Equal(0, second.Index);
    }

    [Fact]
    public void Rejects_a_value_of_the_wrong_kind()
    {
        Assert.Throws<ArgumentException>(() => _engine.Next(Formats, Color(Blue), SelectionA, null));
        Assert.Throws<ArgumentException>(() => _engine.Next(Fonts, Format("0"), SelectionA, null));
    }

    [Fact]
    public void Rejects_an_empty_cycle_and_null_arguments()
    {
        var empty = new CycleDefinition("Empty", "Empty", CycleKind.NumberFormat, Array.Empty<CycleItem>());

        Assert.Throws<ArgumentException>(() => _engine.Next(empty, Format("0"), SelectionA, null));
        Assert.Throws<ArgumentNullException>(() => _engine.Next(null!, Format("0"), SelectionA, null));
        Assert.Throws<ArgumentNullException>(() => _engine.Next(Formats, Format("0"), null!, null));
        Assert.Throws<ArgumentNullException>(() => _engine.IndexOf(null!, Format("0")));
    }

    [Fact]
    public void IndexOf_finds_items_and_never_matches_unknown()
    {
        Assert.Equal(2, _engine.IndexOf(Formats, Format("0.00")));
        Assert.Equal(-1, _engine.IndexOf(Formats, Format("0.000")));
        Assert.Equal(-1, _engine.IndexOf(Formats, CycleValue.Unknown));
        Assert.Equal(2, _engine.IndexOf(Fills, Color(OleColor.NoFill)));
    }

    [Fact]
    public void CycleState_rejects_bad_arguments()
    {
        Assert.Throws<ArgumentNullException>(() => new CycleState(null!, SelectionA, 0, CycleValue.Unknown));
        Assert.Throws<ArgumentNullException>(() => new CycleState("Fmt", null!, 0, CycleValue.Unknown));
        Assert.Throws<ArgumentOutOfRangeException>(() => new CycleState("Fmt", SelectionA, -1, CycleValue.Unknown));
    }
}
