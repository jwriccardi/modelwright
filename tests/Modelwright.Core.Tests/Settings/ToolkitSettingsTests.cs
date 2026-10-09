using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Modelwright.Core.Formatting;
using Modelwright.Core.Keys;
using Modelwright.Core.Settings;
using Xunit;

namespace Modelwright.Core.Tests.Settings;

public class ToolkitSettingsTests
{
    private const string MinimalJson = @"{
  ""schemaVersion"": 1,
  ""keymap"": { ""FontColorCycle"": ""Ctrl+'"" },
  ""cycles"": [
    { ""id"": ""FontColorCycle"", ""displayName"": ""Font"", ""kind"": ""fontColor"",
      ""items"": [ { ""name"": ""Blue"", ""color"": ""rgb(0,0,255)"" }, { ""name"": ""Black"", ""color"": ""#000000"" } ] }
  ]
}";

    [Fact]
    public void Defaults_round_trip_through_json()
    {
        var defaults = ToolkitSettings.Defaults();
        var json = defaults.ToJson();

        var result = ToolkitSettings.FromJson(json);

        Assert.Empty(result.Problems);
        Assert.Equal(SettingsLoadOutcome.Loaded, result.Outcome);
        AssertEquivalent(defaults, result.Settings);
        Assert.Equal(json, result.Settings.ToJson());
    }

    [Fact]
    public void Custom_settings_round_trip_through_json()
    {
        var settings = new ToolkitSettings(
            new[]
            {
                // Not a default cycle's id: a provisional default cycle would be brought up to date (SettingsUpgradeTests).
                new CycleDefinition("MyNumbers", "Numbers \"quoted\" \\ tab\t", CycleKind.NumberFormat, new CycleItem[]
                {
                    new NumberFormatItem("Slash", "m/d/yyyy"),
                    new NumberFormatItem("Unicode", "0.0\"€\";(0.0\"€\");\"–\""),
                }, provisional: true),
                new CycleDefinition("FillColorCycle", "Fill", CycleKind.FillColor, new CycleItem[]
                {
                    new ColorItem("None", OleColor.NoFill),
                    new ColorItem("Gray", OleColor.FromRgb(128, 128, 128)),
                }),
            },
            new Dictionary<string, string> { ["FillColorCycle"] = "Ctrl+Alt+Shift+1", ["About"] = "" },
            undoCellCap: 250,
            diagnosticsLog: false);

        var json = settings.ToJson();

        var result = ToolkitSettings.FromJson(json);

        Assert.Empty(result.Problems);
        Assert.Empty(result.Notes);
        Assert.Equal(json, result.Settings.ToJson());
        Assert.Equal(1, CountOf(json, "\"provisional\": true"));

        // The actions added later are written even when unbound, so they read back unbound.
        Assert.Equal(string.Empty, result.Settings.Keymap[ActionIds.BinaryCycle]);
        Assert.Equal(string.Empty, result.Settings.Keymap[ActionIds.RatioCycle]);
        Assert.Equal(new[] { "MyNumbers", "FillColorCycle" }, result.Settings.Cycles.Select(c => c.Id));
    }

    [Fact]
    public void Json_is_indented_and_readable()
    {
        var json = ToolkitSettings.Defaults().ToJson();

        Assert.StartsWith("{\r\n  \"schemaVersion\": 1,\r\n  \"diagnosticsLog\": true,\r\n  \"undoCellCap\": 10000,\r\n  \"keymap\": {\r\n", json);
        Assert.EndsWith("}\r\n", json);
        Assert.Contains("    \"NumberCycle\": \"Ctrl+Shift+1\",\r\n", json);
        Assert.Contains("        { \"name\": \"Comma 0 Dec No Align\", \"code\": \"#,##0;(#,##0);\\\"–\\\";@\" }", json);
        Assert.Contains("        { \"name\": \"Navy\", \"color\": \"rgb(28,69,135)\" },", json);
        Assert.Contains("        { \"name\": \"No Fill\", \"color\": \"none\" }", json);
        Assert.Contains("    \"BinaryCycle\": \"Ctrl+Shift+Y\",\r\n    \"RatioCycle\": \"Alt+Shift+;\",\r\n", json);
        Assert.DoesNotContain("\"provisional\"", json); // written only when true
        Assert.DoesNotContain("\\u2013", json); // written as the character itself
    }

    [Fact]
    public void Minimal_file_uses_defaults_for_optional_properties()
    {
        var result = ToolkitSettings.FromJson(MinimalJson);

        Assert.Empty(result.Problems);
        Assert.Equal(SettingsLoadOutcome.Loaded, result.Outcome);
        Assert.Equal(ToolkitSettings.DefaultUndoCellCap, result.Settings.UndoCellCap);
        Assert.True(result.Settings.DiagnosticsLog);
        Assert.True(result.Settings.UseKeyboardShortcuts);
        var cycle = result.Settings.Cycles[0];
        Assert.Equal("FontColorCycle", cycle.Id);
        Assert.False(cycle.Provisional);
        Assert.Equal(OleColor.FromRgb(0, 0, 255), ((ColorItem)cycle.Items[0]).Color);
        Assert.Equal("Ctrl+'", result.Settings.Keymap["FontColorCycle"]);
    }

    [Fact]
    public void Keyboard_shortcuts_switch_round_trips_and_is_written_only_when_off()
    {
        var off = ToolkitSettings.Defaults().WithUseKeyboardShortcuts(false);

        var json = off.ToJson();
        var result = ToolkitSettings.FromJson(json);

        Assert.Empty(result.Problems);
        Assert.False(result.Settings.UseKeyboardShortcuts);
        Assert.Equal(json, result.Settings.ToJson());
        Assert.Contains("  \"undoCellCap\": 10000,\r\n  \"useKeyboardShortcuts\": false,\r\n  \"keymap\": {", json);
        AssertEquivalent(off, result.Settings);

        var on = result.Settings.WithUseKeyboardShortcuts(true);
        Assert.True(on.UseKeyboardShortcuts);
        Assert.DoesNotContain("useKeyboardShortcuts", on.ToJson());
        Assert.Equal(ToolkitSettings.Defaults().ToJson(), on.ToJson());
    }

    [Theory]
    [InlineData("true", true)]
    [InlineData("false", false)]
    public void Keyboard_shortcuts_switch_is_read(string value, bool expected)
    {
        var json = MinimalJson.Replace("\"schemaVersion\": 1,", "\"schemaVersion\": 1, \"useKeyboardShortcuts\": " + value + ",");

        var result = ToolkitSettings.FromJson(json);

        Assert.Empty(result.Problems);
        Assert.Equal(expected, result.Settings.UseKeyboardShortcuts);
    }

    [Fact]
    public void Kind_is_case_insensitive()
    {
        var result = ToolkitSettings.FromJson(MinimalJson.Replace("fontColor", "FONTCOLOR"));

        Assert.Empty(result.Problems);
        Assert.Equal(CycleKind.FontColor, result.Settings.Cycles[0].Kind);
    }

    [Theory]
    [InlineData("")]
    [InlineData("{")]
    [InlineData("{ \"schemaVersion\": 1, }")]
    [InlineData("not json")]
    public void Malformed_json_falls_back_to_defaults(string json)
    {
        var result = ToolkitSettings.FromJson(json);

        AssertDefaults(result);
        Assert.StartsWith("the settings are not valid JSON: line 1, column", Assert.Single(result.Problems));
    }

    [Theory]
    [InlineData("[]")]
    [InlineData("42")]
    [InlineData("null")]
    public void Non_object_json_falls_back_to_defaults(string json)
    {
        var result = ToolkitSettings.FromJson(json);

        AssertDefaults(result);
        Assert.Contains("must be a JSON object", Assert.Single(result.Problems));
    }

    [Theory]
    [InlineData("2", "schemaVersion 2 is not supported")]
    [InlineData("0", "schemaVersion 0 is not supported")]
    [InlineData("1.0", "schemaVersion 1.0 is not supported")]
    [InlineData("\"1\"", "schemaVersion \"1\" is not supported")]
    [InlineData("null", "schemaVersion null is not supported")]
    public void Unknown_schema_version_falls_back_to_defaults(string version, string expected)
    {
        var result = ToolkitSettings.FromJson(MinimalJson.Replace("\"schemaVersion\": 1", "\"schemaVersion\": " + version));

        AssertDefaults(result);
        Assert.StartsWith(expected, Assert.Single(result.Problems));
    }

    [Fact]
    public void Missing_schema_version_falls_back_to_defaults()
    {
        var result = ToolkitSettings.FromJson(MinimalJson.Replace("\"schemaVersion\": 1,", string.Empty));

        AssertDefaults(result);
        Assert.StartsWith("schemaVersion is missing", Assert.Single(result.Problems));
    }

    [Fact]
    public void Bad_key_is_reported()
    {
        var result = ToolkitSettings.FromJson(MinimalJson.Replace("\"Ctrl+'\"", "\"Ctrl+Foo\""));

        AssertDefaults(result);
        var problem = Assert.Single(result.Problems);
        Assert.StartsWith("keymap: 'FontColorCycle': Invalid key chord \"Ctrl+Foo\"", problem);
    }

    [Fact]
    public void Duplicate_chord_is_reported()
    {
        var settings = WithKeymap(new Dictionary<string, string>
        {
            [ActionIds.NumberCycle] = "Ctrl+Shift+K",
            [ActionIds.FillColorCycle] = "shift+ctrl+k",
        });

        var problem = Assert.Single(settings.Validate());

        Assert.Equal("keymap: Ctrl+Shift+K is assigned to both 'NumberCycle' and 'FillColorCycle'.", problem);
    }

    [Theory]
    [InlineData("Ctrl+{", "Ctrl+Shift+[")]
    [InlineData("Ctrl++", "Ctrl+Shift+=")]
    [InlineData("Ctrl+Shift+{", "Ctrl+{")]
    public void Keys_pressed_the_same_way_on_a_us_keyboard_are_reported(string first, string second)
    {
        var settings = WithKeymap(new Dictionary<string, string>
        {
            [ActionIds.NumberCycle] = first,
            [ActionIds.FillColorCycle] = second,
        });

        var problem = Assert.Single(settings.Validate());

        var usKeys = KeyChord.Parse(first).ToUsKeys();
        Assert.Equal(
            $"keymap: {KeyChord.Parse(first)} ('NumberCycle') and {KeyChord.Parse(second)} ('FillColorCycle') are the same keys on a US keyboard ({usKeys}).",
            problem);
    }

    [Fact]
    public void Shifted_and_unshifted_keys_of_different_keys_are_not_duplicates()
    {
        var settings = WithKeymap(new Dictionary<string, string>
        {
            [ActionIds.NumberCycle] = "Ctrl+{",
            [ActionIds.FillColorCycle] = "Ctrl+[",
            [ActionIds.DateCycle] = "Ctrl+Shift+]",
        });

        Assert.Empty(settings.Validate());
    }

    [Fact]
    public void Duplicate_chord_in_a_file_falls_back_to_defaults()
    {
        var json = ToolkitSettings.Defaults().ToJson().Replace("\"Ctrl+Shift+2\"", "\"Ctrl+Shift+1\"");

        var result = ToolkitSettings.FromJson(json);

        AssertDefaults(result);
        Assert.Contains("Ctrl+Shift+1 is assigned to both 'NumberCycle' and 'DateCycle'", Assert.Single(result.Problems));
    }

    [Fact]
    public void Empty_cycle_is_reported()
    {
        var json = MinimalJson.Replace(
            "[ { \"name\": \"Blue\", \"color\": \"rgb(0,0,255)\" }, { \"name\": \"Black\", \"color\": \"#000000\" } ]",
            "[]");

        var result = ToolkitSettings.FromJson(json);

        AssertDefaults(result);
        Assert.Equal("cycle 'FontColorCycle': it has no items; a cycle needs at least one.", Assert.Single(result.Problems));
    }

    [Fact]
    public void Duplicate_cycle_ids_are_reported_ignoring_case()
    {
        var item = new CycleItem[] { new ColorItem("Blue", OleColor.FromRgb(0, 0, 255)) };
        var settings = new ToolkitSettings(
            new[]
            {
                new CycleDefinition("FontColorCycle", "A", CycleKind.FontColor, item),
                new CycleDefinition("fontcolorcycle", "B", CycleKind.FontColor, item),
            },
            new Dictionary<string, string>());

        var problem = Assert.Single(settings.Validate());

        Assert.Equal("cycle 'fontcolorcycle': another cycle has the same id (ids are compared ignoring case).", problem);
    }

    [Fact]
    public void Unknown_action_is_reported()
    {
        var settings = WithKeymap(new Dictionary<string, string> { ["numberCycle"] = "Ctrl+Shift+1" });

        var problem = Assert.Single(settings.Validate());

        Assert.StartsWith("keymap: 'numberCycle' is not an action. Known actions: NumberCycle, DateCycle,", problem);
    }

    [Fact]
    public void Bound_cycle_action_without_a_cycle_is_reported()
    {
        var settings = new ToolkitSettings(
            Array.Empty<CycleDefinition>(),
            new Dictionary<string, string> { [ActionIds.DateCycle] = "Ctrl+Shift+2", [ActionIds.NumberCycle] = "", [ActionIds.About] = "Ctrl+Alt+Shift+F12" });

        var problem = Assert.Single(settings.Validate());

        Assert.Equal("keymap: 'DateCycle' has a key but there is no cycle with id 'DateCycle'.", problem);
    }

    [Fact]
    public void Empty_key_leaves_an_action_unbound()
    {
        var settings = WithKeymap(new Dictionary<string, string> { [ActionIds.NumberCycle] = "", [ActionIds.DateCycle] = "" });

        Assert.Empty(settings.Validate());
    }

    [Fact]
    public void Undo_cap_below_one_is_reported()
    {
        var settings = new ToolkitSettings(Array.Empty<CycleDefinition>(), new Dictionary<string, string>(), undoCellCap: 0);

        Assert.Equal("undoCellCap is 0; it must be at least 1.", Assert.Single(settings.Validate()));
    }

    [Fact]
    public void Cycle_problems_are_included()
    {
        var settings = new ToolkitSettings(
            new[] { new CycleDefinition("X", "X", CycleKind.FontColor, new CycleItem[] { new ColorItem("None", OleColor.NoFill) }) },
            new Dictionary<string, string>());

        Assert.Contains("only valid in a fill color cycle", Assert.Single(settings.Validate()));
    }

    [Theory]
    [InlineData("\"keymapp\": {},", "settings: unknown property \"keymapp\"")]
    [InlineData("\"undoCellCap\": \"many\",", "undoCellCap must be a whole number, not \"many\".")]
    [InlineData("\"undoCellCap\": 1.5,", "undoCellCap must be a whole number, not 1.5.")]
    [InlineData("\"undoCellCap\": 99999999999,", "undoCellCap must be a whole number, not 99999999999.")]
    [InlineData("\"diagnosticsLog\": \"yes\",", "diagnosticsLog must be true or false, not \"yes\".")]
    [InlineData("\"diagnosticsLog\": 1,", "diagnosticsLog must be true or false, not 1.")]
    [InlineData("\"useKeyboardShortcuts\": \"no\",", "useKeyboardShortcuts must be true or false, not \"no\".")]
    [InlineData("\"useKeyboardShortcuts\": 0,", "useKeyboardShortcuts must be true or false, not 0.")]
    [InlineData("\"useKeyboardShortcuts\": null,", "useKeyboardShortcuts must be true or false, not null.")]
    public void Bad_top_level_properties_are_reported(string insert, string expected)
    {
        var json = MinimalJson.Replace("\"schemaVersion\": 1,", "\"schemaVersion\": 1, " + insert);

        var result = ToolkitSettings.FromJson(json);

        AssertDefaults(result);
        Assert.StartsWith(expected, Assert.Single(result.Problems));
    }

    [Theory]
    [InlineData("\"keymap\": { \"FontColorCycle\": \"Ctrl+'\" },", "\"keymap\": [],", "keymap must be an object")]
    [InlineData("\"keymap\": { \"FontColorCycle\": \"Ctrl+'\" },", "", "keymap is missing.")]
    [InlineData("\"Ctrl+'\"", "null", "keymap: 'FontColorCycle' must be a key string")]
    [InlineData("\"Ctrl+'\"", "7", "keymap: 'FontColorCycle' must be a key string")]
    public void Bad_keymap_is_reported(string find, string replace, string expected)
    {
        var result = ToolkitSettings.FromJson(MinimalJson.Replace(find, replace));

        AssertDefaults(result);
        Assert.StartsWith(expected, Assert.Single(result.Problems));
    }

    [Fact]
    public void Missing_or_wrongly_typed_cycles_are_reported()
    {
        var missing = "{ \"schemaVersion\": 1, \"keymap\": {} }";
        var notArray = "{ \"schemaVersion\": 1, \"keymap\": {}, \"cycles\": {} }";
        var notObject = "{ \"schemaVersion\": 1, \"keymap\": {}, \"cycles\": [ 5 ] }";

        Assert.Equal("cycles is missing.", Assert.Single(ToolkitSettings.FromJson(missing).Problems));
        Assert.StartsWith("cycles must be an array", Assert.Single(ToolkitSettings.FromJson(notArray).Problems));
        Assert.Equal("cycle 1 must be an object.", Assert.Single(ToolkitSettings.FromJson(notObject).Problems));
    }

    [Theory]
    [InlineData("\"id\": \"FontColorCycle\", ", "", "cycle 1: id is missing.")]
    [InlineData("\"id\": \"FontColorCycle\"", "\"id\": 3", "cycle 1: id must be a string, not 3.")]
    [InlineData("\"displayName\": \"Font\", ", "", "cycle 'FontColorCycle': displayName is missing.")]
    [InlineData("\"kind\": \"fontColor\"", "\"kind\": \"color\"", "cycle 'FontColorCycle': kind \"color\" is not one of numberFormat, fontColor, fillColor.")]
    [InlineData("\"kind\": \"fontColor\",", "\"kind\": \"fontColor\", \"provisional\": \"no\",", "cycle 'FontColorCycle': provisional must be true or false, not \"no\".")]
    [InlineData("\"kind\": \"fontColor\",", "\"kind\": \"fontColor\", \"colour\": 1,", "cycle 'FontColorCycle': unknown property \"colour\"")]
    [InlineData("\"items\": [", "\"elements\": [", "cycle 'FontColorCycle': unknown property \"elements\"")]
    public void Bad_cycle_properties_are_reported(string find, string replace, string expected)
    {
        var result = ToolkitSettings.FromJson(MinimalJson.Replace(find, replace));

        AssertDefaults(result);
        Assert.StartsWith(expected, result.Problems[0]);
    }

    [Fact]
    public void Items_that_are_not_an_array_are_reported()
    {
        var json = MinimalJson.Replace(
            "[ { \"name\": \"Blue\", \"color\": \"rgb(0,0,255)\" }, { \"name\": \"Black\", \"color\": \"#000000\" } ]",
            "\"blue\"");

        Assert.Equal("cycle 'FontColorCycle': items must be an array.", Assert.Single(ToolkitSettings.FromJson(json).Problems));
    }

    [Theory]
    [InlineData("\"rgb(0,0,255)\"", "\"blue\"", "cycle 'FontColorCycle', item 1: Invalid color \"blue\"")]
    [InlineData("\"rgb(0,0,255)\"", "\"rgb(0,0,256)\"", "cycle 'FontColorCycle', item 1: Invalid color \"rgb(0,0,256)\"")]
    [InlineData("\"color\": \"rgb(0,0,255)\"", "\"code\": \"0\"", "cycle 'FontColorCycle', item 1: unknown property \"code\"")]
    [InlineData("{ \"name\": \"Blue\", \"color\": \"rgb(0,0,255)\" }", "\"Blue\"", "cycle 'FontColorCycle', item 1 must be an object.")]
    [InlineData("\"name\": \"Blue\", ", "", "cycle 'FontColorCycle', item 1: name is missing.")]
    [InlineData("\"color\": \"rgb(0,0,255)\"", "\"color\": 255", "cycle 'FontColorCycle', item 1: color must be a string, not 255.")]
    public void Bad_color_items_are_reported(string find, string replace, string expected)
    {
        var result = ToolkitSettings.FromJson(MinimalJson.Replace(find, replace));

        AssertDefaults(result);
        Assert.StartsWith(expected, result.Problems[0]);
    }

    [Theory]
    [InlineData("{ \"name\": \"Zero\" }", "cycle 'N', item 1: code is missing.")]
    [InlineData("{ \"name\": \"Zero\", \"color\": \"none\" }", "cycle 'N', item 1: unknown property \"color\"")]
    [InlineData("{ \"name\": \"Zero\", \"code\": \"\" }", "cycle 'N', item 1: the number format code is blank.")]
    [InlineData("{ \"name\": \"Zero\", \"code\": [] }", "cycle 'N', item 1: code must be a string, not an array.")]
    [InlineData("{ \"name\": {}, \"code\": \"0\" }", "cycle 'N', item 1: name must be a string, not an object.")]
    public void Bad_number_format_items_are_reported(string item, string expected)
    {
        var json = "{ \"schemaVersion\": 1, \"keymap\": {}, \"cycles\": [ { \"id\": \"N\", \"displayName\": \"N\", \"kind\": \"numberFormat\", \"items\": [ " + item + " ] } ] }";

        var result = ToolkitSettings.FromJson(json);

        AssertDefaults(result);
        Assert.StartsWith(expected, result.Problems[0]);
    }

    [Fact]
    public void All_problems_are_reported_together()
    {
        var json = MinimalJson
            .Replace("\"Ctrl+'\"", "\"Ctrl+Foo\"")
            .Replace("\"schemaVersion\": 1,", "\"schemaVersion\": 1, \"undoCellCap\": 0,");

        var result = ToolkitSettings.FromJson(json);

        AssertDefaults(result);
        Assert.Equal(2, result.Problems.Count);
    }

    [Fact]
    public void Byte_order_mark_is_accepted()
    {
        var result = ToolkitSettings.FromJson("\uFEFF" + MinimalJson);

        Assert.Empty(result.Problems);
    }

    [Fact]
    public void FindCycle_is_ordinal()
    {
        var defaults = ToolkitSettings.Defaults();

        Assert.NotNull(defaults.FindCycle("NumberCycle"));
        Assert.Null(defaults.FindCycle("numbercycle"));
        Assert.Null(defaults.FindCycle("Nope"));
    }

    [Fact]
    public void Constructor_rejects_bad_arguments()
    {
        var none = new Dictionary<string, string>();

        Assert.Throws<ArgumentNullException>(() => new ToolkitSettings(null!, none));
        Assert.Throws<ArgumentNullException>(() => new ToolkitSettings(Array.Empty<CycleDefinition>(), null!));
        Assert.Throws<ArgumentException>(() => new ToolkitSettings(new CycleDefinition[] { null! }, none));
        Assert.Throws<ArgumentException>(() => new ToolkitSettings(
            Array.Empty<CycleDefinition>(),
            new[] { new KeyValuePair<string, string>("About", "Ctrl+A"), new KeyValuePair<string, string>("About", "Ctrl+B") }));
        Assert.Throws<ArgumentException>(() => new ToolkitSettings(
            Array.Empty<CycleDefinition>(),
            new[] { new KeyValuePair<string, string>("About", null!) }));
        Assert.Throws<ArgumentNullException>(() => ToolkitSettings.FromJson(null!));
    }

    [Fact]
    public void SettingsLoadResult_rejects_nulls()
    {
        Assert.Throws<ArgumentNullException>(() => new SettingsLoadResult(null!, Array.Empty<string>(), SettingsLoadOutcome.Loaded));
        Assert.Throws<ArgumentNullException>(() => new SettingsLoadResult(ToolkitSettings.Defaults(), null!, SettingsLoadOutcome.Loaded));
        Assert.Throws<ArgumentNullException>(() => SettingsLoadResult.Rejected(null!));
        Assert.Empty(new SettingsLoadResult(ToolkitSettings.Defaults(), Array.Empty<string>(), SettingsLoadOutcome.Loaded).Notes);
        Assert.Empty(SettingsLoadResult.Rejected(new[] { "bad" }).Notes);
    }

    [Fact]
    public void Rejected_result_carries_the_defaults_and_the_problems()
    {
        var result = SettingsLoadResult.Rejected(new[] { "bad" });

        AssertDefaults(result);
        Assert.Equal("bad", Assert.Single(result.Problems));
    }

    [Fact]
    public void File_bytes_in_utf8_load()
    {
        var result = ToolkitSettings.FromFileBytes(Encoding.UTF8.GetBytes(MinimalJson.Replace("\"Font\"", "\"Font – €\"")));

        Assert.Empty(result.Problems);
        Assert.Equal(SettingsLoadOutcome.Loaded, result.Outcome);
        Assert.Equal("Font – €", result.Settings.Cycles[0].DisplayName);
    }

    [Fact]
    public void File_bytes_with_a_utf8_byte_order_mark_load()
    {
        var bytes = new byte[] { 0xEF, 0xBB, 0xBF }.Concat(Encoding.UTF8.GetBytes(MinimalJson)).ToArray();

        var result = ToolkitSettings.FromFileBytes(bytes);

        Assert.Equal(SettingsLoadOutcome.Loaded, result.Outcome);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void File_bytes_in_utf16_with_a_byte_order_mark_load(bool bigEndian)
    {
        var encoding = new UnicodeEncoding(bigEndian, byteOrderMark: true);
        var bytes = encoding.GetPreamble().Concat(encoding.GetBytes(MinimalJson.Replace("\"Font\"", "\"Font –\""))).ToArray();

        var result = ToolkitSettings.FromFileBytes(bytes);

        Assert.Equal(SettingsLoadOutcome.Loaded, result.Outcome);
        Assert.Equal("Font –", result.Settings.Cycles[0].DisplayName);
    }

    [Fact]
    public void File_bytes_that_are_not_utf8_are_rejected_not_garbled()
    {
        // "Font é" saved in Windows-1252: 0xE9 alone is not valid UTF-8.
        var bytes = Encoding.UTF8.GetBytes(MinimalJson.Replace("\"Font\"", "\"Font #\""));
        bytes[Array.IndexOf(bytes, (byte)'#')] = 0xE9;

        var result = ToolkitSettings.FromFileBytes(bytes);

        AssertDefaults(result);
        Assert.Equal("settings.json is not valid UTF-8; save it as UTF-8.", Assert.Single(result.Problems));
    }

    [Fact]
    public void File_bytes_with_a_utf16_byte_order_mark_but_invalid_utf16_are_rejected()
    {
        // A lone low surrogate (0xDC00, little-endian) after the byte order mark.
        var bytes = new byte[] { 0xFF, 0xFE, 0x00, 0xDC };

        var result = ToolkitSettings.FromFileBytes(bytes);

        AssertDefaults(result);
        Assert.Equal("settings.json is not valid UTF-16; save it as UTF-8.", Assert.Single(result.Problems));
    }

    [Fact]
    public void File_bytes_over_the_size_limit_are_rejected_before_decoding()
    {
        var bytes = new byte[ToolkitSettings.MaxFileBytes + 1];
        bytes[0] = 0xE9; // invalid UTF-8, which would be reported if the bytes were decoded

        var result = ToolkitSettings.FromFileBytes(bytes);

        AssertDefaults(result);
        Assert.Equal(
            "settings.json is 1,048,577 bytes; the limit is 1,048,576 bytes (1 MB). It is probably not a settings file.",
            Assert.Single(result.Problems));
        Assert.Throws<ArgumentNullException>(() => ToolkitSettings.FromFileBytes(null!));
    }

    [Fact]
    public void File_problems_name_the_given_file()
    {
        var badUtf8 = ToolkitSettings.FromFileBytes(new byte[] { 0x7B, 0xFF, 0x7D }, "team.json");
        var badUtf16 = ToolkitSettings.FromFileBytes(new byte[] { 0xFF, 0xFE, 0x00, 0xDC }, "team.json");
        var tooLarge = ToolkitSettings.FromFileBytes(new byte[ToolkitSettings.MaxFileBytes + 1], "team.json");

        Assert.Equal("team.json is not valid UTF-8; save it as UTF-8.", Assert.Single(badUtf8.Problems));
        Assert.Equal("team.json is not valid UTF-16; save it as UTF-8.", Assert.Single(badUtf16.Problems));
        Assert.StartsWith("team.json is 1,048,577 bytes;", Assert.Single(tooLarge.Problems));
        Assert.StartsWith("x.json is 1,048,577 bytes;", ToolkitSettings.CheckFileSize(ToolkitSettings.MaxFileBytes + 1, "x.json"));
        Assert.Throws<ArgumentNullException>(() => ToolkitSettings.FromFileBytes(new byte[0], null!));
        Assert.Throws<ArgumentNullException>(() => ToolkitSettings.CheckFileSize(0, null!));
    }

    [Fact]
    public void File_size_check_allows_up_to_one_megabyte()
    {
        Assert.Null(ToolkitSettings.CheckFileSize(0));
        Assert.Null(ToolkitSettings.CheckFileSize(ToolkitSettings.MaxFileBytes));
        Assert.NotNull(ToolkitSettings.CheckFileSize(ToolkitSettings.MaxFileBytes + 1));
    }

    [Fact]
    public void Number_format_code_over_255_characters_is_reported()
    {
        var json = "{ \"schemaVersion\": 1, \"keymap\": {}, \"cycles\": [ { \"id\": \"N\", \"displayName\": \"N\", \"kind\": \"numberFormat\", \"items\": [ " +
            "{ \"name\": \"Long\", \"code\": \"" + new string('0', 256) + "\" } ] } ] }";

        var result = ToolkitSettings.FromJson(json);

        AssertDefaults(result);
        Assert.Equal(
            "cycle 'N', item 1: the number format code is 256 characters; Excel accepts at most 255.",
            Assert.Single(result.Problems));
    }

    [Fact]
    public void ActionIds_classify_actions()
    {
        Assert.True(ActionIds.IsKnown(ActionIds.About));
        Assert.False(ActionIds.IsCycle(ActionIds.About));
        Assert.True(ActionIds.IsCycle(ActionIds.BlueBlackToggle));
        Assert.False(ActionIds.IsKnown("about"));
        Assert.True(ActionIds.IsCycle(ActionIds.BinaryCycle));
        Assert.True(ActionIds.IsCycle(ActionIds.RatioCycle));
        Assert.True(ActionIds.IsKnown(ActionIds.TraceIn));
        Assert.True(ActionIds.IsKnown(ActionIds.LastAuditedCell));
        Assert.False(ActionIds.IsCycle(ActionIds.TraceIn));
        Assert.False(ActionIds.IsCycle(ActionIds.LastAuditedCell));
        Assert.Equal(13, ActionIds.All.Count);
        Assert.Equal(10, ActionIds.CycleActions.Count);
        Assert.Equal(new[] { ActionIds.TraceIn, ActionIds.LastAuditedCell }, ActionIds.TraceActions);
        Assert.Equal(new[] { ActionIds.TraceIn, ActionIds.LastAuditedCell, ActionIds.About }, ActionIds.All.Skip(10));
        Assert.Equal(new[] { ActionIds.BinaryCycle, ActionIds.RatioCycle, ActionIds.TraceIn, ActionIds.LastAuditedCell }, ActionIds.AddedLater);
        Assert.All(ActionIds.AddedLater, a => Assert.True(ActionIds.IsKnown(a)));
    }

    [Theory]
    [InlineData(ActionIds.TraceIn, "Trace In")]
    [InlineData(ActionIds.LastAuditedCell, "Last Audited Cell")]
    [InlineData(ActionIds.About, "About")]
    [InlineData(ActionIds.NumberCycle, ActionIds.NumberCycle)]
    [InlineData("Unknown", "Unknown")]
    public void ActionIds_name_the_actions_that_are_not_cycles(string actionId, string expected)
    {
        Assert.Equal(expected, ActionIds.DisplayName(actionId));
    }

    private static ToolkitSettings WithKeymap(Dictionary<string, string> keymap) =>
        new ToolkitSettings(ToolkitSettings.Defaults().Cycles, keymap);

    private static void AssertDefaults(SettingsLoadResult result)
    {
        Assert.Equal(SettingsLoadOutcome.Rejected, result.Outcome);
        Assert.NotEmpty(result.Problems);
        Assert.Equal(ToolkitSettings.Defaults().ToJson(), result.Settings.ToJson());
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

    private static void AssertEquivalent(ToolkitSettings expected, ToolkitSettings actual)
    {
        Assert.Equal(expected.SchemaVersion, actual.SchemaVersion);
        Assert.Equal(expected.UndoCellCap, actual.UndoCellCap);
        Assert.Equal(expected.DiagnosticsLog, actual.DiagnosticsLog);
        Assert.Equal(expected.UseKeyboardShortcuts, actual.UseKeyboardShortcuts);
        Assert.Equal(expected.Keymap.OrderBy(k => k.Key), actual.Keymap.OrderBy(k => k.Key));
        Assert.Equal(expected.Cycles.Count, actual.Cycles.Count);
        for (var i = 0; i < expected.Cycles.Count; i++)
        {
            var e = expected.Cycles[i];
            var a = actual.Cycles[i];
            Assert.Equal(e.Id, a.Id);
            Assert.Equal(e.DisplayName, a.DisplayName);
            Assert.Equal(e.Kind, a.Kind);
            Assert.Equal(e.Provisional, a.Provisional);
            Assert.Equal(e.Items.Select(item => (item.GetType(), item.Name, item.Value)), a.Items.Select(item => (item.GetType(), item.Name, item.Value)));
        }
    }
}
