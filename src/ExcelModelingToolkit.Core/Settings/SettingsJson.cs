using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelModelingToolkit.Core.Formatting;
using ExcelModelingToolkit.Core.Json;

namespace ExcelModelingToolkit.Core.Settings;

/// <summary>
/// The settings file format (schema version 1):
/// <code>
/// {
///   "schemaVersion": 1,
///   "diagnosticsLog": true,              (optional, default true)
///   "undoCellCap": 10000,                (optional, default 10000)
///   "keymap": { "NumberCycle": "Ctrl+Shift+1", ... },   ("" leaves an action unbound; see below)
///   "cycles": [
///     { "id": "NumberCycle", "displayName": "Number", "kind": "numberFormat", "provisional": true,
///       "items": [ { "name": "...", "code": "..." } ] },
///     { "id": "FillColorCycle", "displayName": "Fill Color", "kind": "fillColor",
///       "items": [ { "name": "Navy", "color": "rgb(28,69,135)" }, { "name": "No Fill", "color": "none" } ] }
///   ]
/// }
/// </code>
/// <c>kind</c> is <c>numberFormat</c>, <c>fontColor</c> or <c>fillColor</c> (any case). Colors are
/// <c>rgb(r,g,b)</c>, <c>#RRGGBB</c> or <c>none</c> (<see cref="OleColor.Parse"/>). <c>provisional</c> is
/// optional (default false) and written only when true. Unknown properties are errors, so typos are reported.
/// An action missing from the keymap is unbound, except one added since files were first written
/// (<see cref="ActionIds.AddedLater"/>): the file predates it, so it gets its default
/// (<see cref="ToolkitSettings.FromJson"/>). The writer always lists those.
/// </summary>
internal static class SettingsJson
{
    private const string NumberFormatKind = "numberFormat";
    private const string FontColorKind = "fontColor";
    private const string FillColorKind = "fillColor";

    private static readonly string[] RootProperties = { "schemaVersion", "diagnosticsLog", "undoCellCap", "keymap", "cycles" };
    private static readonly string[] CycleProperties = { "id", "displayName", "kind", "provisional", "items" };
    private static readonly string[] NumberFormatItemProperties = { "name", "code" };
    private static readonly string[] ColorItemProperties = { "name", "color" };

    /// <summary>Writes <paramref name="settings"/> (whatever their validity) as indented JSON.</summary>
    public static string Write(ToolkitSettings settings)
    {
        // Every action added since files were first written is listed (unbound as ""), so that reading the file
        // back does not take a missing one for a file that predates it (SettingsUpgrade).
        var keymap = new JsonObject();
        var actionIds = settings.Keymap.Keys.Union(ActionIds.AddedLater, StringComparer.Ordinal);
        foreach (var actionId in ToolkitSettings.InKeymapOrder(actionIds))
        {
            keymap.Add(actionId, settings.Keymap.TryGetValue(actionId, out var key) ? key : string.Empty);
        }

        var cycles = new List<object?>();
        foreach (var cycle in settings.Cycles)
        {
            var items = new List<object?>();
            foreach (var item in cycle.Items)
            {
                var node = new JsonObject();
                node.Add("name", item.Name);
                switch (item)
                {
                    case NumberFormatItem format:
                        node.Add("code", format.Code);
                        break;
                    case ColorItem color:
                        node.Add("color", FormatColor(color.Color));
                        break;
                }

                items.Add(node);
            }

            var cycleNode = new JsonObject();
            cycleNode.Add("id", cycle.Id);
            cycleNode.Add("displayName", cycle.DisplayName);
            cycleNode.Add("kind", KindName(cycle.Kind));
            if (cycle.Provisional)
            {
                cycleNode.Add("provisional", true);
            }

            cycleNode.Add("items", items);
            cycles.Add(cycleNode);
        }

        var root = new JsonObject();
        root.Add("schemaVersion", ToolkitSettings.CurrentSchemaVersion);
        root.Add("diagnosticsLog", settings.DiagnosticsLog);
        root.Add("undoCellCap", settings.UndoCellCap);
        root.Add("keymap", keymap);
        root.Add("cycles", cycles);
        return JsonWriter.Write(root);
    }

    /// <summary>
    /// Reads settings JSON. Returns null and adds to <paramref name="problems"/> if the JSON is malformed, the
    /// schema version is missing or unsupported, or any property is missing, unknown or of the wrong type.
    /// Does not run <see cref="ToolkitSettings.Validate"/>.
    /// </summary>
    public static ToolkitSettings? Read(string json, List<string> problems)
    {
        object? rootNode;
        try
        {
            rootNode = JsonParser.Parse(json);
        }
        catch (FormatException ex)
        {
            problems.Add($"the settings are not valid JSON: {ex.Message}");
            return null;
        }

        if (!(rootNode is JsonObject root))
        {
            problems.Add("the settings must be a JSON object: { ... }.");
            return null;
        }

        // The schema version comes first: under another version, nothing else can be interpreted.
        if (!root.TryGet("schemaVersion", out var versionNode))
        {
            problems.Add($"schemaVersion is missing; this version of the add-in reads schemaVersion {ToolkitSettings.CurrentSchemaVersion}.");
            return null;
        }

        if (!TryGetInt(versionNode, out var version) || version != ToolkitSettings.CurrentSchemaVersion)
        {
            problems.Add($"schemaVersion {Describe(versionNode)} is not supported; this version of the add-in reads schemaVersion {ToolkitSettings.CurrentSchemaVersion}.");
            return null;
        }

        var start = problems.Count;
        CheckProperties(root, RootProperties, "settings", problems);
        var diagnosticsLog = ReadOptionalBool(root, "diagnosticsLog", true, problems);
        var undoCellCap = ReadOptionalInt(root, "undoCellCap", ToolkitSettings.DefaultUndoCellCap, problems);
        var keymap = ReadKeymap(root, problems);
        var cycles = ReadCycles(root, problems);
        return problems.Count == start ? new ToolkitSettings(cycles, keymap, undoCellCap, diagnosticsLog) : null;
    }

    private static List<KeyValuePair<string, string>> ReadKeymap(JsonObject root, List<string> problems)
    {
        var keymap = new List<KeyValuePair<string, string>>();
        if (!root.TryGet("keymap", out var node))
        {
            problems.Add("keymap is missing.");
            return keymap;
        }

        if (!(node is JsonObject map))
        {
            problems.Add("keymap must be an object: { \"ActionId\": \"Ctrl+Shift+1\", ... }.");
            return keymap;
        }

        foreach (var entry in map.Properties)
        {
            if (entry.Value is string key)
            {
                keymap.Add(new KeyValuePair<string, string>(entry.Key, key));
            }
            else
            {
                problems.Add($"keymap: '{entry.Key}' must be a key string such as \"Ctrl+Shift+1\" (or \"\" to leave it unbound).");
            }
        }

        return keymap;
    }

    private static List<CycleDefinition> ReadCycles(JsonObject root, List<string> problems)
    {
        var cycles = new List<CycleDefinition>();
        if (!root.TryGet("cycles", out var node))
        {
            problems.Add("cycles is missing.");
            return cycles;
        }

        if (!(node is List<object?> array))
        {
            problems.Add("cycles must be an array: [ { ... }, ... ].");
            return cycles;
        }

        for (var i = 0; i < array.Count; i++)
        {
            var cycle = ReadCycle(array[i], i, problems);
            if (cycle is not null)
            {
                cycles.Add(cycle);
            }
        }

        return cycles;
    }

    private static CycleDefinition? ReadCycle(object? node, int index, List<string> problems)
    {
        var label = $"cycle {index + 1}";
        if (!(node is JsonObject obj))
        {
            problems.Add($"{label} must be an object.");
            return null;
        }

        var start = problems.Count;
        var id = ReadRequiredString(obj, "id", label, problems);
        if (id is not null)
        {
            label = $"cycle '{id}'";
        }

        CheckProperties(obj, CycleProperties, label, problems);
        var displayName = ReadRequiredString(obj, "displayName", label, problems);
        var kindText = ReadRequiredString(obj, "kind", label, problems);
        var provisional = ReadOptionalBool(obj, "provisional", false, problems, label);
        CycleKind? kind = null;
        if (kindText is not null)
        {
            kind = ParseKind(kindText);
            if (kind is null)
            {
                problems.Add($"{label}: kind \"{kindText}\" is not one of {NumberFormatKind}, {FontColorKind}, {FillColorKind}.");
            }
        }

        var items = new List<CycleItem>();
        if (!obj.TryGet("items", out var itemsNode))
        {
            problems.Add($"{label}: items is missing.");
        }
        else if (!(itemsNode is List<object?> itemArray))
        {
            problems.Add($"{label}: items must be an array.");
        }
        else if (kind is not null)
        {
            for (var i = 0; i < itemArray.Count; i++)
            {
                var item = ReadItem(itemArray[i], kind.Value, $"{label}, item {i + 1}", problems);
                if (item is not null)
                {
                    items.Add(item);
                }
            }
        }

        return problems.Count == start ? new CycleDefinition(id!, displayName!, kind!.Value, items, provisional) : null;
    }

    private static CycleItem? ReadItem(object? node, CycleKind kind, string label, List<string> problems)
    {
        if (!(node is JsonObject obj))
        {
            problems.Add($"{label} must be an object.");
            return null;
        }

        var start = problems.Count;
        var isFormat = kind == CycleKind.NumberFormat;
        CheckProperties(obj, isFormat ? NumberFormatItemProperties : ColorItemProperties, label, problems);
        var name = ReadRequiredString(obj, "name", label, problems);
        if (isFormat)
        {
            var code = ReadRequiredString(obj, "code", label, problems);
            return problems.Count == start ? new NumberFormatItem(name!, code!) : null;
        }

        var colorText = ReadRequiredString(obj, "color", label, problems);
        if (colorText is null)
        {
            return null;
        }

        try
        {
            var color = OleColor.Parse(colorText);
            return problems.Count == start ? new ColorItem(name!, color) : null;
        }
        catch (FormatException ex)
        {
            problems.Add($"{label}: {ex.Message}");
            return null;
        }
    }

    private static void CheckProperties(JsonObject obj, string[] allowed, string label, List<string> problems)
    {
        foreach (var property in obj.Properties)
        {
            if (Array.IndexOf(allowed, property.Key) < 0)
            {
                problems.Add($"{label}: unknown property \"{property.Key}\" (expected {string.Join(", ", allowed)}).");
            }
        }
    }

    private static string? ReadRequiredString(JsonObject obj, string name, string label, List<string> problems)
    {
        if (!obj.TryGet(name, out var node))
        {
            problems.Add($"{label}: {name} is missing.");
            return null;
        }

        if (node is string text)
        {
            return text;
        }

        problems.Add($"{label}: {name} must be a string, not {Describe(node)}.");
        return null;
    }

    private static bool ReadOptionalBool(JsonObject obj, string name, bool fallback, List<string> problems, string? label = null)
    {
        if (!obj.TryGet(name, out var node))
        {
            return fallback;
        }

        if (node is bool flag)
        {
            return flag;
        }

        problems.Add($"{(label is null ? string.Empty : label + ": ")}{name} must be true or false, not {Describe(node)}.");
        return fallback;
    }

    private static int ReadOptionalInt(JsonObject obj, string name, int fallback, List<string> problems)
    {
        if (!obj.TryGet(name, out var node))
        {
            return fallback;
        }

        if (TryGetInt(node, out var value))
        {
            return value;
        }

        problems.Add($"{name} must be a whole number, not {Describe(node)}.");
        return fallback;
    }

    private static bool TryGetInt(object? node, out int value)
    {
        value = 0;
        return node is JsonNumber number &&
            int.TryParse(number.Text, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out value);
    }

    private static string Describe(object? node) => node switch
    {
        null => "null",
        string text => $"\"{text}\"",
        bool flag => flag ? "true" : "false",
        JsonNumber number => number.Text,
        JsonObject _ => "an object",
        _ => "an array",
    };

    private static CycleKind? ParseKind(string text)
    {
        if (string.Equals(text, NumberFormatKind, StringComparison.OrdinalIgnoreCase))
        {
            return CycleKind.NumberFormat;
        }

        if (string.Equals(text, FontColorKind, StringComparison.OrdinalIgnoreCase))
        {
            return CycleKind.FontColor;
        }

        if (string.Equals(text, FillColorKind, StringComparison.OrdinalIgnoreCase))
        {
            return CycleKind.FillColor;
        }

        return null;
    }

    private static string KindName(CycleKind kind) => kind switch
    {
        CycleKind.NumberFormat => NumberFormatKind,
        CycleKind.FontColor => FontColorKind,
        CycleKind.FillColor => FillColorKind,
        _ => kind.ToString(),
    };

    private static string FormatColor(OleColor color) => color.IsNoFill
        ? "none"
        : string.Format(CultureInfo.InvariantCulture, "rgb({0},{1},{2})", color.R, color.G, color.B);
}
