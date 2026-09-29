using System.Linq;
using ExcelModelingToolkit.Core.Formatting;
using ExcelModelingToolkit.Core.Keys;
using ExcelModelingToolkit.Core.Settings;
using Xunit;

namespace ExcelModelingToolkit.Core.Tests.Settings;

public class DefaultSettingsTests
{
    private const char EnDash = '–';

    private static readonly ToolkitSettings Defaults = ToolkitSettings.Defaults();

    /// <summary>The default keymap and the exact OnKey strings the spike proved (docs/PLAN.md section 4.2).</summary>
    public static TheoryData<string, string, string> Keymap => new TheoryData<string, string, string>
    {
        { ActionIds.NumberCycle, "Ctrl+Shift+1", "^+1" },
        { ActionIds.DateCycle, "Ctrl+Shift+2", "^+2" },
        { ActionIds.CurrencyCycle, "Ctrl+Shift+4", "^+4" },
        { ActionIds.PercentCycle, "Ctrl+Shift+5", "^+5" },
        { ActionIds.MultipleCycle, "Ctrl+Shift+8", "^+8" },
        { ActionIds.FontColorCycle, "Ctrl+'", "^'" },
        { ActionIds.FillColorCycle, "Ctrl+Shift+K", "^+k" },
        { ActionIds.BlueBlackToggle, "Ctrl+;", "^;" },
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
        Assert.Equal(9, Defaults.Keymap.Count);
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
            new[] { "Number", "Date", "Currency", "Percent", "Multiple", "Font Color", "Fill Color", "Blue/Black" },
            Defaults.Cycles.Select(c => c.DisplayName));
    }

    [Fact]
    public void Provisional_flag_is_set_on_exactly_date_currency_percent_and_multiple()
    {
        Assert.Equal(
            new[] { ActionIds.DateCycle, ActionIds.CurrencyCycle, ActionIds.PercentCycle, ActionIds.MultipleCycle },
            Defaults.Cycles.Where(c => c.Provisional).Select(c => c.Id));
    }

    [Fact]
    public void General_number_cycle_is_the_owners_exact_codes()
    {
        var cycle = Defaults.FindCycle(ActionIds.NumberCycle)!;

        Assert.Equal(CycleKind.NumberFormat, cycle.Kind);
        Assert.Equal(
            new[]
            {
                ("Comma 0 Dec Lg Align", "_(#,##0_)_%;(#,##0)_%;_(\"–\"_)_%;_(@_)_%"),
                ("Comma 1 Dec Lg Align", "_(#,##0.0_)_%;(#,##0.0)_%;_(\"–\"_)_%;_(@_)_%"),
                ("Comma 2 Dec Lg Align", "_(#,##0.00_)_%;(#,##0.00)_%;_(\"–\"_)_%;_(@_)_%"),
                ("Comma 0 Dec No Align", "#,##0;(#,##0);\"–\";@"),
            },
            cycle.Items.Cast<NumberFormatItem>().Select(i => (i.Name, i.Code)));
    }

    [Fact]
    public void Provisional_number_cycles_have_the_planned_codes()
    {
        Assert.Equal(
            new[] { "yyyy\"A\"", "yyyy\"E\"", "mm-dd-yyyy", "yyyy-mm-dd" },
            Codes(ActionIds.DateCycle));
        Assert.Equal(
            new[]
            {
                "_([$$]#,##0_)_%;([$$]#,##0)_%;_(\"–\"_)_%;_(@_)_%",
                "_([$$]#,##0.0_)_%;([$$]#,##0.0)_%;_(\"–\"_)_%;_(@_)_%",
                "_([$$]#,##0.00_)_%;([$$]#,##0.00)_%;_(\"–\"_)_%;_(@_)_%",
            },
            Codes(ActionIds.CurrencyCycle));
        Assert.Equal(
            new[]
            {
                "_(0.0%_);(0.0%);_(\"–\"_)_%;_(@_)_%",
                "_(0%_);(0%);_(\"–\"_)_%;_(@_)_%",
                "_(0.00%_);(0.00%);_(\"–\"_)_%;_(@_)_%",
            },
            Codes(ActionIds.PercentCycle));
        Assert.Equal(
            new[]
            {
                "_(0.0\"x\"_)_%;(0.0\"x\")_%;_(\"–\"_)_%;_(@_)_%",
                "_(0.00\"x\"_)_%;(0.00\"x\")_%;_(\"–\"_)_%;_(@_)_%",
                "_(0\"x\"_)_%;(0\"x\")_%;_(\"–\"_)_%;_(@_)_%",
            },
            Codes(ActionIds.MultipleCycle));
    }

    [Fact]
    public void Dashes_are_en_dashes_not_hyphens()
    {
        var codes = Defaults.Cycles
            .Where(c => c.Id != ActionIds.DateCycle)
            .SelectMany(c => c.Items.OfType<NumberFormatItem>())
            .Select(i => i.Code)
            .ToList();

        Assert.Equal(13, codes.Count); // 4 number + 3 currency + 3 percent + 3 multiple
        Assert.All(codes, code => Assert.Contains("\"" + EnDash + "\"", code));
        Assert.All(codes, code => Assert.DoesNotContain("\"-\"", code));
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

    private static string[] Codes(string cycleId) =>
        Defaults.FindCycle(cycleId)!.Items.Cast<NumberFormatItem>().Select(i => i.Code).ToArray();

    private static (string, string)[] Colors(string cycleId) =>
        Defaults.FindCycle(cycleId)!.Items.Cast<ColorItem>().Select(i => (i.Name, i.Color.ToString())).ToArray();
}
