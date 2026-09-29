using System;
using System.Globalization;
using System.Linq;
using ExcelModelingToolkit.Core.Formatting;
using ExcelModelingToolkit.Core.Keys;
using ExcelModelingToolkit.Core.Settings;
using Xunit;

namespace ExcelModelingToolkit.Core.Tests.Settings;

/// <summary>
/// Settings, colors and keys must parse and print the same whatever the user's culture: de-DE uses a decimal
/// comma and "." as the group separator; tr-TR upper-cases "i" to a dotted "İ" and lower-cases "I" to a dotless "ı".
/// </summary>
public class CultureInvarianceTests
{
    public static readonly TheoryData<string> Cultures = new TheoryData<string> { "de-DE", "tr-TR" };

    private const string CapJson = @"{
  ""schemaVersion"": 1,
  ""undoCellCap"": UNDO_CAP,
  ""keymap"": { ""FillColorCycle"": ""ctrl+shift+i"" },
  ""cycles"": [
    { ""id"": ""FillColorCycle"", ""displayName"": ""Fill"", ""kind"": ""FILLCOLOR"",
      ""items"": [ { ""name"": ""Navy"", ""color"": ""RGB(28,69,135)"" }, { ""name"": ""None"", ""color"": ""NONE"" } ] },
    { ""id"": ""NumberCycle"", ""displayName"": ""Number"", ""kind"": ""NumberFormat"",
      ""items"": [ { ""name"": ""Comma"", ""code"": ""#,##0.00;(#,##0.00)"" } ] }
  ]
}";

    [Theory]
    [MemberData(nameof(Cultures))]
    public void Settings_json_round_trip_is_culture_independent(string culture)
    {
        static string Run()
        {
            var json = ToolkitSettings.Defaults().ToJson();
            var result = ToolkitSettings.FromJson(json);
            var custom = ToolkitSettings.FromJson(CapJson.Replace("UNDO_CAP", "2500"));
            return string.Join(
                "|",
                json,
                result.Outcome,
                result.Settings.ToJson(),
                custom.Outcome,
                string.Join(";", custom.Problems),
                custom.Settings.ToJson());
        }

        var invariant = InCulture(string.Empty, Run);

        Assert.Equal(invariant, InCulture(culture, Run));
        Assert.Contains("\"undoCellCap\": 10000,", invariant);
        Assert.Contains("|Loaded|", invariant);
        Assert.Contains("\"color\": \"rgb(28,69,135)\"", invariant);
    }

    [Theory]
    [MemberData(nameof(Cultures))]
    public void Color_parsing_is_culture_independent(string culture)
    {
        var inputs = new[] { "rgb(28,69,135)", "RGB( 28 , 69 , 135 )", "#1c4587", "#1C4587", "none", "NONE", "None", "rgb(1.5,2,3)", "rgb(1,2,3,4)", "#1c45" };

        string Run() => string.Join("|", inputs.Select(input =>
        {
            try
            {
                var color = OleColor.Parse(input);
                return color.IsNoFill ? "none" : color.OleValue.ToString(CultureInfo.InvariantCulture) + "=" + color;
            }
            catch (FormatException ex)
            {
                return "error: " + ex.Message;
            }
        }));

        var invariant = InCulture(string.Empty, Run);

        Assert.Equal(invariant, InCulture(culture, Run));
        Assert.StartsWith("8865052=#1C4587|8865052=#1C4587|8865052=#1C4587|8865052=#1C4587|none|none|none|error: ", invariant);
    }

    [Theory]
    [MemberData(nameof(Cultures))]
    public void Key_parsing_is_culture_independent(string culture)
    {
        var inputs = new[] { "ctrl+shift+i", "CTRL+SHIFT+I", "Ctrl+I", "ctrl+alt+shift+f12", "Ctrl+'", "Ctrl+;", "Alt+Shift+Insert", "Ctrl+İ" };

        string Run() => string.Join("|", inputs.Select(input =>
        {
            try
            {
                var chord = KeyChord.Parse(input);
                return chord.ToDisplayString() + "=" + chord.ToOnKeyString();
            }
            catch (FormatException ex)
            {
                return "error: " + ex.Message;
            }
        }));

        var invariant = InCulture(string.Empty, Run);

        Assert.Equal(invariant, InCulture(culture, Run));
        Assert.StartsWith("Ctrl+Shift+I=^+i|Ctrl+Shift+I=^+i|Ctrl+I=^i|Ctrl+Alt+Shift+F12=^%+{F12}|", invariant);
    }

    /// <summary>
    /// undoCellCap must be a plain JSON integer. A quoted number and exponent or decimal forms (1e4, 10000.0) are
    /// rejected, even where they denote a whole number, so the file has one spelling for each value.
    /// </summary>
    [Theory]
    [MemberData(nameof(Cultures))]
    public void Numeric_fields_parse_the_same_in_every_culture(string culture)
    {
        var inputs = new[] { "10000", "\"10000\"", "1e4", "1E4", "10000.0", "-5", "0", "2147483648" };

        string Run() => string.Join("\n", inputs.Select(input =>
        {
            var result = ToolkitSettings.FromJson(CapJson.Replace("UNDO_CAP", input));
            return result.Outcome == SettingsLoadOutcome.Loaded
                ? "ok " + result.Settings.UndoCellCap.ToString(CultureInfo.InvariantCulture)
                : "rejected: " + string.Join("; ", result.Problems);
        }));

        var invariant = InCulture(string.Empty, Run);

        Assert.Equal(invariant, InCulture(culture, Run));
        Assert.Equal(
            new[]
            {
                "ok 10000",
                "rejected: undoCellCap must be a whole number, not \"10000\".",
                "rejected: undoCellCap must be a whole number, not 1e4.",
                "rejected: undoCellCap must be a whole number, not 1E4.",
                "rejected: undoCellCap must be a whole number, not 10000.0.",
                "rejected: undoCellCap is -5; it must be at least 1.",
                "rejected: undoCellCap is 0; it must be at least 1.",
                "rejected: undoCellCap must be a whole number, not 2147483648.",
            },
            invariant.Split('\n'));
    }

    [Fact]
    public void Culture_is_restored_after_each_run()
    {
        var before = CultureInfo.CurrentCulture.Name;

        Assert.Throws<InvalidOperationException>(() => InCulture<string>("tr-TR", () => throw new InvalidOperationException()));

        Assert.Equal(before, CultureInfo.CurrentCulture.Name);
    }

    /// <summary>Runs <paramref name="action"/> with the current (UI) culture set to <paramref name="name"/> ("" = invariant), then restores it.</summary>
    private static T InCulture<T>(string name, Func<T> action)
    {
        var culture = CultureInfo.CurrentCulture;
        var uiCulture = CultureInfo.CurrentUICulture;
        try
        {
            var target = name.Length == 0 ? CultureInfo.InvariantCulture : new CultureInfo(name);
            CultureInfo.CurrentCulture = target;
            CultureInfo.CurrentUICulture = target;
            return action();
        }
        finally
        {
            CultureInfo.CurrentCulture = culture;
            CultureInfo.CurrentUICulture = uiCulture;
        }
    }
}
