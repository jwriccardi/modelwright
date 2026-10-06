using System.Collections.Generic;
using System.Linq;
using ExcelModelingToolkit.Core.Formatting;
using ExcelModelingToolkit.Core.Settings;
using Xunit;

namespace ExcelModelingToolkit.Core.Tests.Settings;

/// <summary>Loading a settings file written by an earlier build (<see cref="ToolkitSettings.FromJson"/>).</summary>
public class SettingsUpgradeTests
{
    /// <summary>
    /// The shape of the file the Phase 3a build wrote: nine actions (no Binary or Ratio), and provisional
    /// placeholders for Date, Currency, Percent and Multiple (shortened here).
    /// </summary>
    private const string EarlierBuildJson = @"{
  ""schemaVersion"": 1,
  ""diagnosticsLog"": false,
  ""undoCellCap"": 500,
  ""keymap"": {
    ""NumberCycle"": ""Ctrl+Shift+1"",
    ""DateCycle"": ""Ctrl+Shift+2"",
    ""CurrencyCycle"": ""Ctrl+Shift+4"",
    ""PercentCycle"": ""Ctrl+Shift+5"",
    ""MultipleCycle"": ""Ctrl+Shift+8"",
    ""FontColorCycle"": ""Ctrl+'"",
    ""FillColorCycle"": ""Ctrl+Shift+K"",
    ""BlueBlackToggle"": """",
    ""About"": ""Ctrl+Alt+Shift+F12""
  },
  ""cycles"": [
    { ""id"": ""NumberCycle"", ""displayName"": ""My Numbers"", ""kind"": ""numberFormat"",
      ""items"": [ { ""name"": ""Plain"", ""code"": ""0"" } ] },
    { ""id"": ""DateCycle"", ""displayName"": ""Date"", ""kind"": ""numberFormat"", ""provisional"": true,
      ""items"": [ { ""name"": ""Month-Day-Year"", ""code"": ""mm-dd-yyyy"" } ] },
    { ""id"": ""CurrencyCycle"", ""displayName"": ""Currency"", ""kind"": ""numberFormat"", ""provisional"": true,
      ""items"": [ { ""name"": ""Currency 0 Dec"", ""code"": ""$#,##0"" } ] },
    { ""id"": ""PercentCycle"", ""displayName"": ""Percent"", ""kind"": ""numberFormat"", ""provisional"": true,
      ""items"": [ { ""name"": ""Percent 1 Dec"", ""code"": ""0.0%"" } ] },
    { ""id"": ""MultipleCycle"", ""displayName"": ""Multiple"", ""kind"": ""numberFormat"", ""provisional"": true,
      ""items"": [ { ""name"": ""Multiple 1 Dec"", ""code"": ""0.0x"" } ] },
    { ""id"": ""FontColorCycle"", ""displayName"": ""Font Color"", ""kind"": ""fontColor"",
      ""items"": [ { ""name"": ""Blue"", ""color"": ""rgb(0,0,255)"" } ] },
    { ""id"": ""FillColorCycle"", ""displayName"": ""Fill Color"", ""kind"": ""fillColor"",
      ""items"": [ { ""name"": ""No Fill"", ""color"": ""none"" } ] },
    { ""id"": ""BlueBlackToggle"", ""displayName"": ""Blue/Black"", ""kind"": ""fontColor"",
      ""items"": [ { ""name"": ""Blue"", ""color"": ""rgb(0,0,255)"" } ] }
  ]
}";

    private static readonly ToolkitSettings Defaults = ToolkitSettings.Defaults();

    [Fact]
    public void Earlier_file_loads_with_the_new_actions_on_their_default_keys()
    {
        var result = ToolkitSettings.FromJson(EarlierBuildJson);

        Assert.Equal(SettingsLoadOutcome.Loaded, result.Outcome);
        Assert.Empty(result.Problems);
        Assert.Equal("Ctrl+Shift+Y", result.Settings.Keymap[ActionIds.BinaryCycle]);
        Assert.Equal("Alt+Shift+;", result.Settings.Keymap[ActionIds.RatioCycle]);
        Assert.Contains("keymap: 'BinaryCycle' is new: added with its default key Ctrl+Shift+Y and the default Binary cycle.", result.Notes);
        Assert.Contains("keymap: 'RatioCycle' is new: added with its default key Alt+Shift+; and the default Ratio cycle.", result.Notes);
    }

    [Fact]
    public void Earlier_file_gets_the_new_default_cycles_appended()
    {
        var settings = ToolkitSettings.FromJson(EarlierBuildJson).Settings;

        Assert.Equal(
            new[] { "NumberCycle", "DateCycle", "CurrencyCycle", "PercentCycle", "MultipleCycle", "FontColorCycle", "FillColorCycle", "BlueBlackToggle", "BinaryCycle", "RatioCycle" },
            settings.Cycles.Select(c => c.Id));
        Assert.True(settings.FindCycle(ActionIds.BinaryCycle)!.HasSameItems(Defaults.FindCycle(ActionIds.BinaryCycle)));
        Assert.True(settings.FindCycle(ActionIds.RatioCycle)!.HasSameItems(Defaults.FindCycle(ActionIds.RatioCycle)));
    }

    [Fact]
    public void Provisional_placeholders_are_replaced_in_place_by_the_default_lists()
    {
        var result = ToolkitSettings.FromJson(EarlierBuildJson);

        foreach (var id in new[] { ActionIds.DateCycle, ActionIds.CurrencyCycle, ActionIds.PercentCycle, ActionIds.MultipleCycle })
        {
            var cycle = result.Settings.FindCycle(id)!;
            Assert.False(cycle.Provisional);
            Assert.True(cycle.HasSameItems(Defaults.FindCycle(id)), id);
        }

        Assert.Contains("cycle 'DateCycle': replaced the provisional placeholder list with the default Date list.", result.Notes);
        Assert.Equal(6, result.Notes.Count); // four placeholders, two new actions
    }

    [Fact]
    public void Everything_else_in_an_earlier_file_is_kept()
    {
        var settings = ToolkitSettings.FromJson(EarlierBuildJson).Settings;

        var number = settings.FindCycle(ActionIds.NumberCycle)!;
        Assert.Equal("My Numbers", number.DisplayName);
        Assert.Equal("0", ((NumberFormatItem)Assert.Single(number.Items)).Code);
        Assert.Single(settings.FindCycle(ActionIds.FontColorCycle)!.Items);
        Assert.Equal(string.Empty, settings.Keymap[ActionIds.BlueBlackToggle]);
        Assert.Equal("Ctrl+Shift+1", settings.Keymap[ActionIds.NumberCycle]);
        Assert.Equal(11, settings.Keymap.Count);
        Assert.Equal(500, settings.UndoCellCap);
        Assert.False(settings.DiagnosticsLog);
    }

    [Fact]
    public void Upgraded_settings_save_and_reload_without_further_changes()
    {
        var upgraded = ToolkitSettings.FromJson(EarlierBuildJson).Settings;

        var reloaded = ToolkitSettings.FromJson(upgraded.ToJson());

        Assert.Equal(SettingsLoadOutcome.Loaded, reloaded.Outcome);
        Assert.Empty(reloaded.Notes);
        Assert.Equal(upgraded.ToJson(), reloaded.Settings.ToJson());
    }

    [Fact]
    public void Current_defaults_load_without_notes()
    {
        var result = ToolkitSettings.FromJson(Defaults.ToJson());

        Assert.Empty(result.Problems);
        Assert.Empty(result.Notes);
    }

    [Fact]
    public void A_new_action_the_file_lists_as_unbound_stays_unbound_without_a_cycle()
    {
        var json = EarlierBuildJson.Replace("\"About\":", "\"BinaryCycle\": \"\", \"About\":");

        var result = ToolkitSettings.FromJson(json);

        Assert.Empty(result.Problems);
        Assert.Equal(string.Empty, result.Settings.Keymap[ActionIds.BinaryCycle]);
        Assert.Null(result.Settings.FindCycle(ActionIds.BinaryCycle));
        Assert.DoesNotContain(result.Notes, n => n.Contains("BinaryCycle"));
        Assert.Equal("Alt+Shift+;", result.Settings.Keymap[ActionIds.RatioCycle]);
    }

    [Fact]
    public void A_new_action_whose_default_key_is_taken_is_left_unbound()
    {
        var json = EarlierBuildJson.Replace("\"NumberCycle\": \"Ctrl+Shift+1\"", "\"NumberCycle\": \"shift+ctrl+y\"");

        var result = ToolkitSettings.FromJson(json);

        Assert.Equal(SettingsLoadOutcome.Loaded, result.Outcome);
        Assert.Empty(result.Problems);
        Assert.Equal(string.Empty, result.Settings.Keymap[ActionIds.BinaryCycle]);
        Assert.NotNull(result.Settings.FindCycle(ActionIds.BinaryCycle)); // ready to bind
        Assert.Contains("keymap: 'BinaryCycle' is new and left unbound: its default key Ctrl+Shift+Y is assigned to 'NumberCycle'.", result.Notes);
    }

    [Fact]
    public void A_users_own_cycle_with_a_new_actions_id_is_kept_and_bound()
    {
        var json = EarlierBuildJson.Replace(
            "\"cycles\": [",
            "\"cycles\": [ { \"id\": \"RatioCycle\", \"displayName\": \"Mine\", \"kind\": \"numberFormat\", \"items\": [ { \"name\": \"R\", \"code\": \"0.00\" } ] },");

        var result = ToolkitSettings.FromJson(json);

        Assert.Empty(result.Problems);
        Assert.Equal("Mine", result.Settings.FindCycle(ActionIds.RatioCycle)!.DisplayName);
        Assert.Single(result.Settings.Cycles, c => c.Id == ActionIds.RatioCycle);
        Assert.Equal("Alt+Shift+;", result.Settings.Keymap[ActionIds.RatioCycle]);
        Assert.Contains("keymap: 'RatioCycle' is new: added with its default key Alt+Shift+;.", result.Notes);
    }

    [Fact]
    public void A_cycle_id_that_differs_only_in_case_leaves_the_new_action_unbound()
    {
        var settings = new ToolkitSettings(
            new[] { new CycleDefinition("binarycycle", "Mine", CycleKind.NumberFormat, new CycleItem[] { new NumberFormatItem("Z", "0") }) },
            new Dictionary<string, string> { [ActionIds.RatioCycle] = string.Empty });
        var json = settings.ToJson().Replace("\"BinaryCycle\": \"\",", string.Empty);

        var result = ToolkitSettings.FromJson(json);

        Assert.Empty(result.Problems);
        Assert.Equal(string.Empty, result.Settings.Keymap[ActionIds.BinaryCycle]);
        Assert.Equal(new[] { "binarycycle" }, result.Settings.Cycles.Select(c => c.Id));
        Assert.Equal("keymap: 'BinaryCycle' is new and left unbound: a cycle's id differs from it only in case.", Assert.Single(result.Notes));
    }

    [Fact]
    public void A_provisional_cycle_without_a_default_keeps_its_items()
    {
        var settings = new ToolkitSettings(
            new[] { new CycleDefinition("MyCycle", "Mine", CycleKind.NumberFormat, new CycleItem[] { new NumberFormatItem("Z", "0") }, provisional: true) },
            new Dictionary<string, string>());

        var result = ToolkitSettings.FromJson(settings.ToJson());

        Assert.Empty(result.Notes);
        Assert.True(Assert.Single(result.Settings.Cycles).Provisional);
    }

    [Fact]
    public void A_provisional_cycle_of_another_kind_than_its_default_is_kept()
    {
        var settings = new ToolkitSettings(
            new[] { new CycleDefinition(ActionIds.DateCycle, "Dates", CycleKind.FontColor, new CycleItem[] { new ColorItem("Red", OleColor.FromRgb(255, 0, 0)) }, provisional: true) },
            new Dictionary<string, string> { [ActionIds.DateCycle] = "Ctrl+Shift+2" });

        var result = ToolkitSettings.FromJson(settings.ToJson());

        Assert.Empty(result.Notes);
        Assert.Equal(CycleKind.FontColor, Assert.Single(result.Settings.Cycles).Kind);
    }

    [Fact]
    public void A_rejected_file_has_no_notes()
    {
        var result = ToolkitSettings.FromJson(EarlierBuildJson.Replace("\"Ctrl+Shift+2\"", "\"Ctrl+Shift+1\""));

        Assert.Equal(SettingsLoadOutcome.Rejected, result.Outcome);
        Assert.Empty(result.Notes);
    }
}
