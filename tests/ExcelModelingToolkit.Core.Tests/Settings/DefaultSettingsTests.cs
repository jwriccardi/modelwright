using System.Linq;
using System.Text;
using ExcelModelingToolkit.Core.Formatting;
using ExcelModelingToolkit.Core.Keys;
using ExcelModelingToolkit.Core.Settings;
using Xunit;

namespace ExcelModelingToolkit.Core.Tests.Settings;

public class DefaultSettingsTests
{
    private const char EnDash = '–';

    private static readonly ToolkitSettings Defaults = ToolkitSettings.Defaults();

    /// <summary>
    /// The default keymap (Macabacus factory ExcelShortcuts) and the exact OnKey strings the spike proved
    /// (docs/PLAN.md section 4.2).
    /// </summary>
    public static TheoryData<string, string, string> Keymap => new TheoryData<string, string, string>
    {
        { ActionIds.NumberCycle, "Ctrl+Shift+1", "^+1" },
        { ActionIds.DateCycle, "Ctrl+Shift+2", "^+2" },
        { ActionIds.CurrencyCycle, "Ctrl+Shift+4", "^+4" },
        { ActionIds.PercentCycle, "Ctrl+Shift+5", "^+5" },
        { ActionIds.MultipleCycle, "Ctrl+Shift+8", "^+8" },
        { ActionIds.BinaryCycle, "Ctrl+Shift+Y", "^+y" },
        { ActionIds.RatioCycle, "Alt+Shift+;", "%+;" },
        { ActionIds.FontColorCycle, "Ctrl+'", "^'" },
        { ActionIds.FillColorCycle, "Ctrl+Shift+K", "^+k" },
        { ActionIds.BlueBlackToggle, "Ctrl+;", "^;" },
        { ActionIds.TraceIn, "Ctrl+Shift+[", "^+{[}" },
        { ActionIds.LastAuditedCell, "Ctrl+Shift+\\", "^+\\" },
        { ActionIds.About, "Ctrl+Alt+Shift+F12", "^%+{F12}" },
    };

    [Theory]
    [MemberData(nameof(Keymap))]
    public void Default_keys_render_the_exact_onkey_strings(string actionId, string key, string onKey)
    {
        Assert.Equal(key, Defaults.Keymap[actionId]);
        Assert.Equal(onKey, KeyChord.Parse(Defaults.Keymap[actionId]).ToOnKeyString());
    }

    [Fact]
    public void Keymap_binds_every_action_and_nothing_else()
    {
        Assert.Equal(ActionIds.All.OrderBy(a => a), Defaults.Keymap.Keys.OrderBy(a => a));
        Assert.Equal(13, Defaults.Keymap.Count);
    }

    [Fact]
    public void Defaults_are_valid()
    {
        Assert.Empty(Defaults.Validate());
        Assert.Equal(ToolkitSettings.CurrentSchemaVersion, Defaults.SchemaVersion);
        Assert.Equal(10000, Defaults.UndoCellCap);
        Assert.True(Defaults.DiagnosticsLog);
    }

    [Fact]
    public void Every_cycle_action_has_a_cycle_in_ribbon_order()
    {
        Assert.Equal(ActionIds.CycleActions, Defaults.Cycles.Select(c => c.Id));
    }

    [Fact]
    public void Display_names_match_the_ribbon_labels()
    {
        Assert.Equal(
            new[] { "Number", "Date", "Currency", "Percent", "Multiple", "Binary", "Ratio", "Font Color", "Fill Color", "Blue/Black" },
            Defaults.Cycles.Select(c => c.DisplayName));
    }

    [Fact]
    public void No_default_cycle_is_provisional()
    {
        Assert.DoesNotContain(Defaults.Cycles, c => c.Provisional);
        Assert.DoesNotContain("\"provisional\"", Defaults.ToJson());
    }

    [Fact]
    public void General_number_cycle_is_the_factory_list()
    {
        AssertFormats(
            ActionIds.NumberCycle,
            ("Comma 0 Dec Lg Align", @"_(#,##0_)_%;(#,##0)_%;_(""–""_)_%;_(@_)_%"),
            ("Comma 1 Dec Lg Align", @"_(#,##0.0_)_%;(#,##0.0)_%;_(""–""_)_%;_(@_)_%"),
            ("Comma 2 Dec Lg Align", @"_(#,##0.00_)_%;(#,##0.00)_%;_(""–""_)_%;_(@_)_%"),
            ("Comma 0 Dec No Align", @"#,##0;(#,##0);""–"";@"));
    }

    [Fact]
    public void Date_cycle_is_the_factory_list()
    {
        // The backslash escapes the letter: 2024 shows as 2024A.
        AssertFormats(
            ActionIds.DateCycle,
            ("m/d/yyyy", @"m/d/yyyy;@"),
            ("Date Text Long", @"mmmm d, yyyy;@"),
            ("Date Actual Year", @"0000\A"),
            ("Date Estimated Year", @"0000\E"));
    }

    [Fact]
    public void Currency_cycle_is_the_factory_local_currency_list()
    {
        AssertFormats(
            ActionIds.CurrencyCycle,
            ("USD 0 Dec Lg Align", @"_([$$]#,##0_)_%;([$$]#,##0)_%;_(""–""_)_%;_(@_)_%"),
            ("USD 1 Dec Lg Align", @"_([$$]#,##0.0_)_%;([$$]#,##0.0)_%;_(""–""_)_%;_(@_)_%"),
            ("USD 2 Dec Lg Align", @"_([$$]#,##0.00_)_%;([$$]#,##0.00)_%;_(""–""_)_%;_(@_)_%"),
            ("USD 0 Dec No Align", @"[$$]#,##0;([$$]#,##0);""–"";@"),
            ("EUR 0 Dec Lg Align", @"_([$€-2]#,##0_)_%;([$€-2]#,##0)_%;_(""–""_)_%;_(@_)_%"),
            ("GBP 0 Dec Lg Align", @"_([$£-809]#,##0_)_%;([$£-809]#,##0)_%;_(""–""_)_%;_(@_)_%"),
            ("YEN 0 Dec Lg Align", @"_([$¥-2]#,##0_)_%;([$¥-2]#,##0)_%;_(""–""_)_%;_(@_)_%"));
    }

    [Fact]
    public void Percent_cycle_is_the_factory_list()
    {
        AssertFormats(
            ActionIds.PercentCycle,
            ("Percent Aligned Neg Pct", @"_(#,##0.0%_);(#,##0.0%);_(""–""_)_%;_(@_)_%"),
            ("Percent Unaligned", @"#,##0.0%;(#,##0.0%);""–"";@"),
            ("Hard Percent Aligned Neg Pct", @"_(#,##0.0""%""_);(#,##0.0""%"");_(""–""_)_%;_(@_)_%"),
            ("Hard Percent Unaligned", @"#,##0.0""%"";(#,##0.0""%"");""–"";@"),
            ("SOFR +", @"""S""+0_)_%;""S""-0_)_%;""S""+0_)_%"),
            ("LIBOR +", @"L+0_)_%;L-0_)_%;L+0_)_%"));
    }

    [Fact]
    public void Multiple_cycle_is_the_factory_list()
    {
        // The x is unquoted and _' pads by an apostrophe's width, as in Macabacus.
        AssertFormats(
            ActionIds.MultipleCycle,
            ("Mult 1 Decimal Aligned Neg Pct", @"_(0.0x_)_)_';_((0.0x)_'_';_(""–""_)_%;_(@_)_%"),
            ("Mult 2 Decimal Aligned Neg Pct", @"_(0.00x_)_)_';_((0.00x)_'_';_(""–""_)_%;_(@_)_%"),
            ("Mult 1 Decimal Unaligned", @"0.0x;(0.0x);""–"""),
            ("Mult 2 Decimal Unaligned", @"0.00x;(0.00x);""–"""));
    }

    [Fact]
    public void Binary_cycle_is_the_factory_list()
    {
        AssertFormats(
            ActionIds.BinaryCycle,
            ("Yes/No", @"""Yes"";""ERROR"";""No"";""ERROR"""),
            ("Y/N", @"""Y"";""ERROR"";""N"";""ERROR"""),
            ("On/Off", @"""On"";""ERROR"";""Off"";""ERROR"""),
            ("True/False", @"""True"";""ERROR"";""False"";""ERROR"""));
    }

    [Fact]
    public void Ratio_cycle_is_the_factory_list()
    {
        AssertFormats(
            ActionIds.RatioCycle,
            ("Exchange Ratio", @"0.0\:1_);(0.0)\:1_);0.0\:1_);@_)"),
            ("Fraction 1", @"# ?/?"),
            ("Fraction 2", @"# ??/??"),
            ("Fraction 3", @"# ???/???"),
            ("Halves", @"# ?/2"),
            ("Thirds", @"# ?/3"));
    }

    [Fact]
    public void Currency_symbols_are_the_exact_characters()
    {
        var codes = Codes(ActionIds.CurrencyCycle);

        Assert.StartsWith("_([$\u20AC-2]", codes[4]); // euro sign
        Assert.StartsWith("_([$\u00A3-809]", codes[5]); // pound sign
        Assert.StartsWith("_([$\u00A5-2]", codes[6]); // yen sign
    }

    [Fact]
    public void Dashes_are_en_dashes_not_hyphens()
    {
        var codes = Defaults.Cycles
            .SelectMany(c => c.Items.OfType<NumberFormatItem>())
            .Select(i => i.Code)
            .Where(code => code.Contains("\"" + EnDash + "\""))
            .ToList();

        Assert.Equal('\u2013', EnDash);
        Assert.Equal(19, codes.Count); // 4 number + 7 currency + 4 percent + 4 multiple
        Assert.All(Defaults.Cycles.SelectMany(c => c.Items.OfType<NumberFormatItem>()), i => Assert.DoesNotContain("\"-\"", i.Code));
    }

    [Fact]
    public void En_dashes_are_written_to_json_as_utf8_bytes_e2_80_93()
    {
        var bytes = new UTF8Encoding(false).GetBytes(Defaults.ToJson());
        var count = 0;
        for (var i = 0; i + 2 < bytes.Length; i++)
        {
            if (bytes[i] == 0xE2 && bytes[i + 1] == 0x80 && bytes[i + 2] == 0x93)
            {
                count++;
            }
        }

        Assert.Equal(19, count);
    }

    [Fact]
    public void Font_color_cycle_is_the_owners_order()
    {
        Assert.Equal(
            new[] { ("Blue", "#0000FF"), ("Green", "#008000"), ("Purple", "#800080"), ("Red", "#FF0000"), ("White", "#FFFFFF"), ("Black", "#000000") },
            Colors(ActionIds.FontColorCycle));
        Assert.Equal(CycleKind.FontColor, Defaults.FindCycle(ActionIds.FontColorCycle)!.Kind);
    }

    [Fact]
    public void Fill_color_cycle_is_the_owners_order_ending_with_no_fill()
    {
        Assert.Equal(
            new[] { ("Light Blue", "#C9DAF8"), ("Light Cyan", "#D2F2FF"), ("Light Pink", "#F4CCCC"), ("Peach", "#FCE5CD"), ("Navy", "#1C4587"), ("No Fill", "none") },
            Colors(ActionIds.FillColorCycle));
        Assert.Equal(CycleKind.FillColor, Defaults.FindCycle(ActionIds.FillColorCycle)!.Kind);
    }

    [Fact]
    public void Blue_black_toggle_is_inputs_blue_then_default_black()
    {
        Assert.Equal(new[] { ("Blue", "#0000FF"), ("Black", "#000000") }, Colors(ActionIds.BlueBlackToggle));
        Assert.Equal(CycleKind.FontColor, Defaults.FindCycle(ActionIds.BlueBlackToggle)!.Kind);
    }

    [Fact]
    public void Each_call_returns_a_fresh_equal_instance()
    {
        var other = DefaultSettings.Create();

        Assert.NotSame(Defaults, other);
        Assert.Equal(Defaults.ToJson(), other.ToJson());
    }

    private static void AssertFormats(string cycleId, params (string Name, string Code)[] expected)
    {
        var cycle = Defaults.FindCycle(cycleId)!;

        Assert.Equal(CycleKind.NumberFormat, cycle.Kind);
        Assert.Equal(expected, cycle.Items.Cast<NumberFormatItem>().Select(i => (i.Name, i.Code)));
    }

    private static string[] Codes(string cycleId) =>
        Defaults.FindCycle(cycleId)!.Items.Cast<NumberFormatItem>().Select(i => i.Code).ToArray();

    private static (string, string)[] Colors(string cycleId) =>
        Defaults.FindCycle(cycleId)!.Items.Cast<ColorItem>().Select(i => (i.Name, i.Color.ToString())).ToArray();
}
