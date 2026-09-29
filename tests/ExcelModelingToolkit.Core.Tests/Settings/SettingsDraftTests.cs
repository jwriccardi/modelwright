using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelModelingToolkit.Core.Formatting;
using ExcelModelingToolkit.Core.Settings;
using Xunit;

namespace ExcelModelingToolkit.Core.Tests.Settings;

public class SettingsDraftTests
{
    private static SettingsDraft DefaultDraft() => new SettingsDraft(ToolkitSettings.Defaults());

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
        var draft = DefaultDraft();
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
        var draft = DefaultDraft();
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

        draft.SetKey(ActionIds.About, "Ctrl+Alt+Shift+F11");

        Assert.Contains("\"About\": \"Ctrl+Alt+Shift+F11\"", Json(draft));
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
    public void Reset_restores_the_factory_defaults()
    {
        var draft = DefaultDraft();
        var percent = Cycle(draft, ActionIds.PercentCycle);
        percent.RemoveItem(0);
        draft.SetKey(ActionIds.About, "Ctrl+Alt+A");
        draft.UndoCellCap = 3;
        draft.DiagnosticsLog = false;

        draft.ResetToDefaults();

        Assert.Equal(ToolkitSettings.Defaults().ToJson(), Json(draft));
        Assert.True(Cycle(draft, ActionIds.PercentCycle).Provisional);
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
