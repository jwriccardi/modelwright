using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Modelwright.Core.Formatting;
using Modelwright.Core.Settings;
using Xunit;

namespace Modelwright.Core.Tests.Settings;

public class SettingsDraftTests
{
    private static SettingsDraft DefaultDraft() => new SettingsDraft(ToolkitSettings.Defaults());

    /// <summary>
    /// The defaults with Date, Currency, Percent and Multiple marked provisional, as earlier builds shipped them. No
    /// default cycle is provisional any more; these keep the provisional handling covered, which stays generic.
    /// </summary>
    private static SettingsDraft ProvisionalDraft()
    {
        var defaults = ToolkitSettings.Defaults();
        var provisional = new[] { ActionIds.DateCycle, ActionIds.CurrencyCycle, ActionIds.PercentCycle, ActionIds.MultipleCycle };
        var cycles = defaults.Cycles.Select(c =>
            provisional.Contains(c.Id) ? new CycleDefinition(c.Id, c.DisplayName, c.Kind, c.Items, provisional: true) : c);
        return new SettingsDraft(new ToolkitSettings(cycles, defaults.Keymap, defaults.UndoCellCap, defaults.DiagnosticsLog));
    }

    private static CycleDraft Cycle(SettingsDraft draft, string id) => draft.Cycles.Single(c => c.Id == id);

    private static string Json(SettingsDraft draft)
    {
        var json = draft.ExportJson(out var problems);
        Assert.Empty(problems);
        return json!;
    }

    [Fact]
    public void Unedited_draft_gives_back_the_same_settings()
    {
        var defaults = ToolkitSettings.Defaults();
        var draft = new SettingsDraft(defaults);

        var settings = draft.ToSettings(out var problems);

        Assert.Empty(problems);
        Assert.Equal(defaults.ToJson(), settings.ToJson());
        Assert.Equal(defaults.Cycles.Select(c => c.Id), draft.Cycles.Select(c => c.Id));
        Assert.Equal(defaults.UndoCellCap, draft.UndoCellCap);
        Assert.Equal(defaults.DiagnosticsLog, draft.DiagnosticsLog);
    }

    [Fact]
    public void Constructor_rejects_null() =>
        Assert.Throws<ArgumentNullException>(() => new SettingsDraft(null!));

    [Fact]
    public void Adds_number_format_and_color_items_at_the_given_position()
    {
        var draft = DefaultDraft();
        var number = Cycle(draft, ActionIds.NumberCycle);
        var font = Cycle(draft, ActionIds.FontColorCycle);
        var fill = Cycle(draft, ActionIds.FillColorCycle);

        Assert.Equal(1, number.AddItem(1));
        var fontCount = font.Items.Count;
        Assert.Equal(fontCount, font.AddItem(fontCount));
        Assert.Equal(0, fill.AddItem(0));

        var format = Assert.IsType<NumberFormatItem>(number.Items[1]);
        Assert.Equal(CycleDraft.NewFormatName, format.Name);
        Assert.Equal(CycleDraft.NewFormatCode, format.Code);
        Assert.Equal(5, number.Items.Count);
        var fontItem = Assert.IsType<ColorItem>(font.Items[font.Items.Count - 1]);
        Assert.Equal(CycleDraft.NewColorName, fontItem.Name);
        Assert.Equal(OleColor.FromRgb(0, 0, 0), fontItem.Color);
        Assert.Equal(OleColor.FromRgb(255, 255, 255), Assert.IsType<ColorItem>(fill.Items[0]).Color);
        draft.ToSettings(out var problems);
        Assert.Empty(problems);
    }

    [Fact]
    public void Add_rejects_an_index_outside_the_list()
    {
        var number = Cycle(DefaultDraft(), ActionIds.NumberCycle);

        Assert.Throws<ArgumentOutOfRangeException>(() => number.AddItem(-1));
        Assert.Throws<ArgumentOutOfRangeException>(() => number.AddItem(5));
    }

    [Fact]
    public void Removes_and_moves_items()
    {
        var number = Cycle(DefaultDraft(), ActionIds.NumberCycle);
        var names = number.Items.Select(i => i.Name).ToArray(); // 0 Dec, 1 Dec, 2 Dec, 0 Dec No Align

        Assert.Equal(0, number.MoveItemUp(1));
        Assert.Equal(new[] { names[1], names[0], names[2], names[3] }, number.Items.Select(i => i.Name));
        Assert.Equal(0, number.MoveItemUp(0));
        Assert.Equal(3, number.MoveItemDown(2));
        Assert.Equal(3, number.MoveItemDown(3));
        Assert.Equal(new[] { names[1], names[0], names[3], names[2] }, number.Items.Select(i => i.Name));

        number.RemoveItem(0);

        Assert.Equal(new[] { names[0], names[3], names[2] }, number.Items.Select(i => i.Name));
        Assert.Throws<ArgumentOutOfRangeException>(() => number.RemoveItem(3));
        Assert.Throws<ArgumentOutOfRangeException>(() => number.MoveItemUp(-1));
        Assert.Throws<ArgumentOutOfRangeException>(() => number.MoveItemDown(3));
    }

    [Fact]
    public void Renames_and_edits_codes_and_colors()
    {
        var draft = DefaultDraft();
        var number = Cycle(draft, ActionIds.NumberCycle);
        var fill = Cycle(draft, ActionIds.FillColorCycle);

        number.RenameItem(0, "Thousands");
        number.SetCode(0, "#,##0,\"k\"");
        fill.RenameItem(4, "Dark");
        fill.SetColor(4, OleColor.FromRgb(1, 2, 3));
        fill.SetColor(0, OleColor.NoFill);

        var settings = draft.ToSettings(out var problems);
        Assert.Empty(problems);
        var format = (NumberFormatItem)settings.FindCycle(ActionIds.NumberCycle)!.Items[0];
        Assert.Equal("Thousands", format.Name);
        Assert.Equal("#,##0,\"k\"", format.Code);
        var dark = (ColorItem)settings.FindCycle(ActionIds.FillColorCycle)!.Items[4];
        Assert.Equal("Dark", dark.Name);
        Assert.Equal(OleColor.FromRgb(1, 2, 3), dark.Color);
        Assert.True(((ColorItem)settings.FindCycle(ActionIds.FillColorCycle)!.Items[0]).Color.IsNoFill);
    }

    [Fact]
    public void Codes_only_in_number_cycles_and_no_fill_only_in_fill_cycles()
    {
        var draft = DefaultDraft();
        var number = Cycle(draft, ActionIds.NumberCycle);
        var font = Cycle(draft, ActionIds.FontColorCycle);

        Assert.Throws<InvalidOperationException>(() => font.SetCode(0, "0.0"));
        Assert.Throws<InvalidOperationException>(() => number.SetColor(0, OleColor.FromRgb(0, 0, 255)));
        Assert.Throws<ArgumentException>(() => font.SetColor(0, OleColor.NoFill));
        Assert.Throws<ArgumentNullException>(() => number.SetCode(0, null!));
        Assert.Throws<ArgumentNullException>(() => number.RenameItem(0, null!));
        Assert.False(font.AllowsNoFill);
        Assert.True(Cycle(draft, ActionIds.FillColorCycle).AllowsNoFill);
        Assert.Equal(OleColor.FromRgb(0, 0, 255), ((ColorItem)font.Items[0]).Color);
    }

    public static TheoryData<string, Action<CycleDraft>> ItemEdits => new TheoryData<string, Action<CycleDraft>>
    {
        { "add", c => c.AddItem(0) },
        { "remove", c => c.RemoveItem(0) },
        { "move up", c => c.MoveItemUp(1) },
        { "move down", c => c.MoveItemDown(0) },
        { "rename", c => c.RenameItem(0, "Mine") },
        { "code", c => c.SetCode(0, "0.0%") },
    };

    [Theory]
    [MemberData(nameof(ItemEdits))]
    public void Editing_a_provisional_cycles_items_clears_provisional(string edit, Action<CycleDraft> apply)
    {
        var draft = ProvisionalDraft();
        var percent = Cycle(draft, ActionIds.PercentCycle);
        Assert.True(percent.Provisional);

        apply(percent);

        Assert.False(percent.Provisional, edit);
        Assert.False(draft.ToSettings(out _).FindCycle(ActionIds.PercentCycle)!.Provisional);
        Assert.True(Cycle(draft, ActionIds.DateCycle).Provisional); // other cycles keep their flag
        Assert.Equal(3, CountOf(Json(draft), "\"provisional\": true"));
    }

    [Fact]
    public void Edits_that_change_nothing_keep_provisional()
    {
        var draft = ProvisionalDraft();
        var percent = Cycle(draft, ActionIds.PercentCycle);
        var first = (NumberFormatItem)percent.Items[0];

        percent.RenameItem(0, first.Name);
        percent.SetCode(0, first.Code);
        percent.MoveItemUp(0);
        percent.MoveItemDown(percent.Items.Count - 1);
        draft.SetKey(ActionIds.PercentCycle, "Ctrl+Alt+5");
        draft.UndoCellCap = 5;

        Assert.True(percent.Provisional);
    }

    [Fact]
    public void Validation_reports_problems_with_the_settings_file_rules()
    {
        var draft = DefaultDraft();
        var number = Cycle(draft, ActionIds.NumberCycle);
        number.RenameItem(0, " ");
        number.SetCode(1, string.Empty);
        var blueBlack = Cycle(draft, ActionIds.BlueBlackToggle);
        blueBlack.RemoveItem(0);
        blueBlack.RemoveItem(0);
        draft.UndoCellCap = 0;
        draft.SetKey(ActionIds.DateCycle, "Ctrl+Shift+1");
        draft.SetKey(ActionIds.About, "Ctrl+Nope");

        var settings = draft.ToSettings(out var problems);

        Assert.Equal(settings.Validate(), problems);
        Assert.Contains("cycle 'NumberCycle', item 1: the name is blank.", problems);
        Assert.Contains("cycle 'NumberCycle', item 2: the number format code is blank.", problems);
        Assert.Contains("cycle 'BlueBlackToggle': it has no items; a cycle needs at least one.", problems);
        Assert.Contains("undoCellCap is 0; it must be at least 1.", problems);
        Assert.Contains("keymap: Ctrl+Shift+1 is assigned to both 'NumberCycle' and 'DateCycle'.", problems);
        Assert.Contains(problems, p => p.StartsWith("keymap: 'About': Invalid key chord \"Ctrl+Nope\"", StringComparison.Ordinal));
        Assert.Equal(6, problems.Count);
        Assert.Null(draft.ExportJson(out var exportProblems));
        Assert.Equal(problems, exportProblems);
    }

    [Fact]
    public void Too_long_code_is_reported()
    {
        var draft = DefaultDraft();
        Cycle(draft, ActionIds.NumberCycle).SetCode(0, new string('0', NumberFormatItem.MaxCodeLength + 1));

        draft.ToSettings(out var problems);

        Assert.Single(problems);
        Assert.Contains("Excel accepts at most 255", problems[0]);
    }

    [Fact]
    public void Keys_are_stored_in_display_form_when_valid_and_as_typed_otherwise()
    {
        var draft = DefaultDraft();

        draft.SetKey(ActionIds.FillColorCycle, "  shift+control+k ");
        draft.SetKey(ActionIds.DateCycle, " Ctrl+Nope ");

        Assert.Equal("Ctrl+Shift+K", draft.GetKey(ActionIds.FillColorCycle));
        Assert.Equal("Ctrl+Nope", draft.GetKey(ActionIds.DateCycle));
    }

    [Fact]
    public void Clearing_a_key_unbinds_it_and_the_file_keeps_the_empty_key()
    {
        var draft = DefaultDraft();

        draft.ClearKey(ActionIds.BlueBlackToggle);

        Assert.Equal(string.Empty, draft.GetKey(ActionIds.BlueBlackToggle));
        Assert.Null(draft.KeyProblem(ActionIds.BlueBlackToggle));
        var settings = draft.ToSettings(out var problems);
        Assert.Empty(problems);
        Assert.Equal(string.Empty, settings.Keymap[ActionIds.BlueBlackToggle]);
        Assert.Contains("\"BlueBlackToggle\": \"\"", settings.ToJson());
    }

    [Fact]
    public void Key_problems_name_the_parse_error_or_the_other_action()
    {
        var draft = DefaultDraft();
        Assert.All(draft.ShortcutActions, a => Assert.Null(draft.KeyProblem(a)));

        draft.SetKey(ActionIds.DateCycle, "Ctrl+Shift+K"); // Fill Color's key
        draft.SetKey(ActionIds.About, "K");

        Assert.Equal("Ctrl+Shift+K is also assigned to Fill Color.", draft.KeyProblem(ActionIds.DateCycle));
        Assert.Equal("Ctrl+Shift+K is also assigned to Date.", draft.KeyProblem(ActionIds.FillColorCycle));
        Assert.StartsWith("Invalid key chord \"K\": it has no modifier.", draft.KeyProblem(ActionIds.About));
        Assert.Null(draft.KeyProblem(ActionIds.NumberCycle));

        draft.SetKey(ActionIds.DateCycle, "Ctrl+Shift+2");

        Assert.Null(draft.KeyProblem(ActionIds.FillColorCycle));
    }

    [Fact]
    public void Unknown_actions_cannot_be_bound()
    {
        var draft = DefaultDraft();

        Assert.Throws<ArgumentException>(() => draft.SetKey("Nope", "Ctrl+Q"));
        Assert.Throws<ArgumentNullException>(() => draft.SetKey(ActionIds.About, null!));
        Assert.Throws<ArgumentNullException>(() => draft.GetKey(null!));
        Assert.Equal(ActionIds.All, draft.ShortcutActions);
    }

    [Fact]
    public void An_action_missing_from_the_keymap_is_unbound_until_set()
    {
        var draft = new SettingsDraft(new ToolkitSettings(ToolkitSettings.Defaults().Cycles, new Dictionary<string, string>()));

        Assert.Equal(string.Empty, draft.GetKey(ActionIds.About));
        Assert.DoesNotContain("\"About\"", Json(draft));
        Assert.Equal(string.Empty, draft.GetKey(ActionIds.BinaryCycle));
        Assert.Contains("\"BinaryCycle\": \"\"", Json(draft)); // written, so loading the file keeps it unbound

        draft.SetKey(ActionIds.About, "Ctrl+Alt+Shift+F11");

        Assert.Contains("\"About\": \"Ctrl+Alt+Shift+F11\"", Json(draft));
    }

    [Fact]
    public void Binary_and_ratio_are_listed_with_their_cycles_and_keys()
    {
        var draft = DefaultDraft();

        Assert.Equal(ActionIds.All, draft.ShortcutActions);
        Assert.Equal("Binary", draft.ActionDisplayName(ActionIds.BinaryCycle));
        Assert.Equal("Ratio", draft.ActionDisplayName(ActionIds.RatioCycle));
        Assert.Equal("Ctrl+Shift+Y", draft.GetKey(ActionIds.BinaryCycle));
        Assert.Equal("Alt+Shift+;", draft.GetKey(ActionIds.RatioCycle));
        Assert.Null(draft.KeyProblem(ActionIds.BinaryCycle));
        Assert.Null(draft.KeyProblem(ActionIds.RatioCycle));
        Assert.Equal(4, Cycle(draft, ActionIds.BinaryCycle).Items.Count);
        Assert.Equal(6, Cycle(draft, ActionIds.RatioCycle).Items.Count);
    }

    [Fact]
    public void Trace_actions_are_listed_with_their_names_and_keys()
    {
        var draft = DefaultDraft();

        Assert.Equal("Trace In", draft.ActionDisplayName(ActionIds.TraceIn));
        Assert.Equal("Last Audited Cell", draft.ActionDisplayName(ActionIds.LastAuditedCell));
        Assert.Equal("Ctrl+Shift+[", draft.GetKey(ActionIds.TraceIn));
        Assert.Equal("Ctrl+Shift+\\", draft.GetKey(ActionIds.LastAuditedCell));
        Assert.Null(draft.KeyProblem(ActionIds.TraceIn));
        Assert.Null(draft.KeyProblem(ActionIds.LastAuditedCell));

        draft.SetKey(ActionIds.About, "Ctrl+{");

        Assert.Equal("Ctrl+{ is pressed with the same keys as Ctrl+Shift+[ (Trace In) on a US keyboard.", draft.KeyProblem(ActionIds.About));
    }

    [Fact]
    public void Action_display_names_come_from_the_cycles()
    {
        var draft = DefaultDraft();

        Assert.Equal("Number", draft.ActionDisplayName(ActionIds.NumberCycle));
        Assert.Equal("Blue/Black", draft.ActionDisplayName(ActionIds.BlueBlackToggle));
        Assert.Equal("About", draft.ActionDisplayName(ActionIds.About));
        Assert.Equal("Other", draft.ActionDisplayName("Other"));
    }

    [Fact]
    public void General_options_are_written()
    {
        var draft = DefaultDraft();

        draft.UndoCellCap = 42;
        draft.DiagnosticsLog = false;

        var settings = draft.ToSettings(out var problems);
        Assert.Empty(problems);
        Assert.Equal(42, settings.UndoCellCap);
        Assert.False(settings.DiagnosticsLog);
    }

    [Fact]
    public void Keyboard_shortcuts_switch_is_loaded_and_written()
    {
        var draft = new SettingsDraft(ToolkitSettings.Defaults().WithUseKeyboardShortcuts(false));
        Assert.False(draft.UseKeyboardShortcuts);

        draft.UseKeyboardShortcuts = true;
        Assert.True(draft.ToSettings(out _).UseKeyboardShortcuts);

        draft.UseKeyboardShortcuts = false;
        var settings = draft.ToSettings(out var problems);
        Assert.Empty(problems);
        Assert.False(settings.UseKeyboardShortcuts);
        Assert.NotEqual(new SettingsDraft(ToolkitSettings.Defaults()).ContentJson(), draft.ContentJson()); // Cancel would discard a change.

        draft.ResetToDefaults();
        Assert.True(draft.UseKeyboardShortcuts);
    }

    [Fact]
    public void Reset_restores_the_factory_defaults()
    {
        var draft = ProvisionalDraft();
        var percent = Cycle(draft, ActionIds.PercentCycle);
        percent.RemoveItem(0);
        draft.SetKey(ActionIds.About, "Ctrl+Alt+A");
        draft.UndoCellCap = 3;
        draft.DiagnosticsLog = false;

        draft.ResetToDefaults();

        Assert.Equal(ToolkitSettings.Defaults().ToJson(), Json(draft));
        Assert.DoesNotContain(draft.Cycles, c => c.Provisional); // the factory lists replace the placeholders
        Assert.NotSame(percent, Cycle(draft, ActionIds.PercentCycle));
    }

    [Fact]
    public void Export_then_import_gives_identical_json()
    {
        var draft = DefaultDraft();
        var number = Cycle(draft, ActionIds.NumberCycle);
        number.AddItem(4);
        number.RenameItem(4, "Thousands \"k\" – €");
        number.SetCode(4, "#,##0,\"k\";(#,##0,\"k\");\"–\"");
        number.MoveItemUp(4);
        var fill = Cycle(draft, ActionIds.FillColorCycle);
        fill.SetColor(1, OleColor.Parse("#ABCDEF"));
        fill.MoveItemDown(0);
        Cycle(draft, ActionIds.DateCycle).RemoveItem(0);
        draft.SetKey(ActionIds.About, "ctrl+alt+shift+f11");
        draft.ClearKey(ActionIds.MultipleCycle);
        draft.UndoCellCap = 777;
        draft.DiagnosticsLog = false;
        var exported = Json(draft);

        var imported = DefaultDraft();
        var problems = imported.ImportJson(exported);

        Assert.Empty(problems);
        Assert.Equal(exported, Json(imported));
        Assert.NotEqual(ToolkitSettings.Defaults().ToJson(), exported);

        // Through bytes too, as the dialog reads a file (UTF-8 without a byte order mark, as it writes one).
        var fromFile = DefaultDraft();
        Assert.Empty(fromFile.ImportFile(new UTF8Encoding(false).GetBytes(exported)));
        Assert.Equal(exported, Json(fromFile));
    }

    [Theory]
    [InlineData("not json")]
    [InlineData("{ \"schemaVersion\": 2 }")]
    [InlineData("{ \"schemaVersion\": 1, \"keymap\": {}, \"cycles\": [], \"extra\": 1 }")]
    [InlineData("{ \"schemaVersion\": 1, \"undoCellCap\": 0, \"keymap\": {}, \"cycles\": [] }")]
    public void Failed_import_reports_problems_and_leaves_the_draft_unchanged(string json)
    {
        var draft = DefaultDraft();
        Cycle(draft, ActionIds.NumberCycle).RemoveItem(0);
        draft.UndoCellCap = 99;
        var before = Json(draft);

        var problems = draft.ImportJson(json);

        Assert.NotEmpty(problems);
        Assert.Equal(before, Json(draft));
    }

    [Fact]
    public void Failed_file_import_reports_encoding_and_size_problems()
    {
        var draft = DefaultDraft();
        var before = Json(draft);

        var badUtf8 = draft.ImportFile(new byte[] { 0x7B, 0xFF, 0x7D });
        var tooLarge = draft.ImportFile(new byte[ToolkitSettings.MaxFileBytes + 1]);

        Assert.Equal("settings.json is not valid UTF-8; save it as UTF-8.", Assert.Single(badUtf8));
        Assert.Contains("the limit is", Assert.Single(tooLarge));
        Assert.Equal(before, Json(draft));
        Assert.Throws<ArgumentNullException>(() => draft.ImportJson(null!));
        Assert.Throws<ArgumentNullException>(() => draft.ImportFile(null!));
    }

    [Fact]
    public void Import_replaces_cycles_keys_and_options()
    {
        var custom = new ToolkitSettings(
            new[]
            {
                new CycleDefinition(ActionIds.FontColorCycle, "Font", CycleKind.FontColor, new CycleItem[]
                {
                    new ColorItem("Teal", OleColor.FromRgb(0, 128, 128)),
                }),
            },
            new Dictionary<string, string> { [ActionIds.FontColorCycle] = "Ctrl+Alt+F" },
            undoCellCap: 12,
            diagnosticsLog: false);
        var draft = DefaultDraft();

        Assert.Empty(draft.ImportJson(custom.ToJson()));

        Assert.Equal(new[] { ActionIds.FontColorCycle }, draft.Cycles.Select(c => c.Id));
        Assert.Equal("Ctrl+Alt+F", draft.GetKey(ActionIds.FontColorCycle));
        Assert.Equal(string.Empty, draft.GetKey(ActionIds.NumberCycle));
        Assert.Equal(12, draft.UndoCellCap);
        Assert.False(draft.DiagnosticsLog);
        Assert.Equal(custom.ToJson(), Json(draft));
    }

    [Fact]
    public void Editing_the_draft_does_not_change_the_settings_it_came_from()
    {
        var original = ToolkitSettings.Defaults();
        var json = original.ToJson();
        var draft = new SettingsDraft(original);

        Cycle(draft, ActionIds.NumberCycle).RemoveItem(0);
        draft.SetKey(ActionIds.About, string.Empty);

        Assert.Equal(json, original.ToJson());
    }

    [Fact]
    public void Edits_undone_by_hand_keep_provisional()
    {
        var draft = ProvisionalDraft();
        var percent = Cycle(draft, ActionIds.PercentCycle);
        var first = (NumberFormatItem)percent.Items[0];
        var before = draft.ContentJson();

        percent.RenameItem(0, "Mine");
        Assert.False(percent.Provisional);
        percent.RenameItem(0, first.Name);
        Assert.True(percent.Provisional);

        percent.SetCode(0, "0.0%");
        Assert.False(percent.Provisional);
        percent.SetCode(0, first.Code);
        Assert.True(percent.Provisional);

        percent.MoveItemDown(0);
        Assert.False(percent.Provisional);
        percent.MoveItemUp(1);
        Assert.True(percent.Provisional);

        percent.AddItem(1);
        Assert.False(percent.Provisional);
        percent.RemoveItem(1);
        Assert.True(percent.Provisional);

        Assert.Equal(before, draft.ContentJson());
        Assert.Equal(4, CountOf(Json(draft), "\"provisional\": true"));
    }

    [Fact]
    public void Same_values_under_another_name_are_not_the_provisional_items()
    {
        var percent = Cycle(ProvisionalDraft(), ActionIds.PercentCycle);
        var first = percent.Items[0].Name;

        percent.RenameItem(0, first.ToUpperInvariant());

        Assert.False(percent.Provisional);
    }

    [Fact]
    public void A_cycle_that_was_not_provisional_never_becomes_provisional()
    {
        var number = Cycle(DefaultDraft(), ActionIds.NumberCycle);

        Assert.False(number.Provisional);
        number.RenameItem(0, "Mine");
        number.RenameItem(0, ToolkitSettings.Defaults().FindCycle(ActionIds.NumberCycle)!.Items[0].Name);

        Assert.False(number.Provisional);
    }

    [Fact]
    public void Content_json_tells_whether_anything_changed()
    {
        var draft = DefaultDraft();
        var before = draft.ContentJson();

        draft.UndoCellCap = 0; // invalid contents are still written
        Assert.NotEqual(before, draft.ContentJson());
        Assert.Contains("\"undoCellCap\": 0", draft.ContentJson());

        draft.UndoCellCap = ToolkitSettings.DefaultUndoCellCap;
        Assert.Equal(before, draft.ContentJson());
    }

    [Fact]
    public void Keys_from_the_settings_are_stored_in_display_form()
    {
        var keymap = new Dictionary<string, string>
        {
            [ActionIds.FillColorCycle] = "shift+control+k",
            [ActionIds.DateCycle] = "Ctrl+Nope",
            [ActionIds.About] = string.Empty,
        };

        var draft = new SettingsDraft(new ToolkitSettings(ToolkitSettings.Defaults().Cycles, keymap));

        Assert.Equal("Ctrl+Shift+K", draft.GetKey(ActionIds.FillColorCycle));
        Assert.Equal("Ctrl+Nope", draft.GetKey(ActionIds.DateCycle));
        Assert.Equal(string.Empty, draft.GetKey(ActionIds.About));

        var before = draft.ContentJson();
        draft.SetKey(ActionIds.FillColorCycle, "Ctrl+Shift+K"); // re-entering the key changes nothing
        Assert.Equal(before, draft.ContentJson());
    }

    [Fact]
    public void Key_problems_name_a_key_pressed_the_same_way_on_a_us_keyboard()
    {
        var draft = DefaultDraft();
        draft.SetKey(ActionIds.TraceIn, string.Empty); // frees its default, Ctrl+Shift+[

        draft.SetKey(ActionIds.DateCycle, "Ctrl+{");
        draft.SetKey(ActionIds.About, "Ctrl+Shift+[");

        Assert.Equal("Ctrl+{ is pressed with the same keys as Ctrl+Shift+[ (About) on a US keyboard.", draft.KeyProblem(ActionIds.DateCycle));
        Assert.Equal("Ctrl+Shift+[ is pressed with the same keys as Ctrl+{ (Date) on a US keyboard.", draft.KeyProblem(ActionIds.About));
        draft.ToSettings(out var problems);
        Assert.Contains("keymap: Ctrl+{ ('DateCycle') and Ctrl+Shift+[ ('About') are the same keys on a US keyboard (Ctrl+Shift+[).", problems);

        draft.SetKey(ActionIds.NumberCycle, "Ctrl+Shift+[");

        // An exact duplicate is named first.
        Assert.Equal("Ctrl+Shift+[ is also assigned to Number.", draft.KeyProblem(ActionIds.About));

        draft.SetKey(ActionIds.NumberCycle, "Ctrl+Shift+1");
        draft.SetKey(ActionIds.About, "Ctrl+[");

        Assert.Null(draft.KeyProblem(ActionIds.DateCycle));
        Assert.Null(draft.KeyProblem(ActionIds.About));
    }

    [Fact]
    public void File_import_problems_name_the_file()
    {
        var draft = DefaultDraft();

        var problems = draft.ImportFile(new byte[] { 0x7B, 0xFF, 0x7D }, "team-settings.json");

        Assert.Equal("team-settings.json is not valid UTF-8; save it as UTF-8.", Assert.Single(problems));
        Assert.Throws<ArgumentNullException>(() => draft.ImportFile(new byte[0], null!));
    }

    [Fact]
    public void Finds_another_item_with_the_same_value()
    {
        var draft = DefaultDraft();
        var number = Cycle(draft, ActionIds.NumberCycle);
        var font = Cycle(draft, ActionIds.FontColorCycle);

        Assert.All(Enumerable.Range(0, number.Items.Count), i => Assert.Equal(-1, number.IndexOfSameValue(i)));

        number.AddItem(4);
        number.AddItem(5); // two "General" items
        font.AddItem(0); // black, like the last item

        Assert.Equal(5, number.IndexOfSameValue(4));
        Assert.Equal(4, number.IndexOfSameValue(5));
        Assert.Equal(font.Items.Count - 1, font.IndexOfSameValue(0));
        Assert.Equal(0, font.IndexOfSameValue(font.Items.Count - 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => number.IndexOfSameValue(6));
    }

    private static int CountOf(string text, string value)
    {
        var count = 0;
        for (var i = text.IndexOf(value, StringComparison.Ordinal); i >= 0; i = text.IndexOf(value, i + 1, StringComparison.Ordinal))
        {
            count++;
        }

        return count;
    }
}
