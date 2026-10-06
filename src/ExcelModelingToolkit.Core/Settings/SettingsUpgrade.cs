using System;
using System.Collections.Generic;
using System.Linq;
using ExcelModelingToolkit.Core.Formatting;
using ExcelModelingToolkit.Core.Keys;

namespace ExcelModelingToolkit.Core.Settings;

/// <summary>
/// Brings settings read from a file written by an earlier build up to date with <see cref="DefaultSettings"/>,
/// without changing anything the user chose. The file itself is not touched; each change is described in a note.
/// </summary>
internal static class SettingsUpgrade
{
    /// <summary>
    /// Returns <paramref name="settings"/> with two kinds of change, each described in <paramref name="notes"/>
    /// (empty, and the same instance returned, if there is nothing to change):
    /// <list type="bullet">
    /// <item>
    /// An action of <see cref="ActionIds.AddedLater"/> that the keymap does not mention (the file predates it) gets
    /// its default key, or stays unbound if another action already uses that key; a cycle action also gets its
    /// default cycle, appended, if there is no cycle with its id. If a cycle's id differs from the action's only in
    /// case, the action stays unbound instead (adding the cycle would make the ids clash).
    /// </item>
    /// <item>
    /// A provisional cycle (<see cref="CycleDefinition.Provisional"/>: the placeholder list an earlier build
    /// shipped, never edited in the settings dialog) whose default is no longer provisional is replaced by that
    /// default, in place.
    /// </item>
    /// </list>
    /// </summary>
    public static ToolkitSettings Apply(ToolkitSettings settings, out IReadOnlyList<string> notes)
    {
        var defaults = DefaultSettings.Create();
        var found = new List<string>();
        var cycles = settings.Cycles.ToList();
        for (var i = 0; i < cycles.Count; i++)
        {
            var current = defaults.FindCycle(cycles[i].Id);
            if (cycles[i].Provisional && current is not null && !current.Provisional && current.Kind == cycles[i].Kind)
            {
                cycles[i] = current;
                found.Add($"cycle '{current.Id}': replaced the provisional placeholder list with the default {current.DisplayName} list.");
            }
        }

        var keymap = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var entry in settings.Keymap)
        {
            keymap.Add(entry.Key, entry.Value);
        }

        foreach (var actionId in ActionIds.AddedLater.Where(a => !keymap.ContainsKey(a)))
        {
            var key = defaults.Keymap[actionId];
            CycleDefinition? added = null;
            if (ActionIds.IsCycle(actionId) && settings.FindCycle(actionId) is null)
            {
                if (cycles.Any(c => string.Equals(c.Id, actionId, StringComparison.OrdinalIgnoreCase)))
                {
                    keymap.Add(actionId, string.Empty);
                    found.Add($"keymap: '{actionId}' is new and left unbound: a cycle's id differs from it only in case.");
                    continue;
                }

                added = defaults.FindCycle(actionId)!;
                cycles.Add(added);
            }

            var owner = OwnerOf(KeyChord.Parse(key), keymap);
            keymap.Add(actionId, owner is null ? key : string.Empty);
            found.Add(owner is null
                ? $"keymap: '{actionId}' is new: added with its default key {key}{(added is null ? string.Empty : $" and the default {added.DisplayName} cycle")}."
                : $"keymap: '{actionId}' is new and left unbound: its default key {key} is assigned to '{owner}'.");
        }

        notes = found;
        return found.Count == 0
            ? settings
            : new ToolkitSettings(cycles, keymap, settings.UndoCellCap, settings.DiagnosticsLog);
    }

    /// <summary>The action whose key in <paramref name="keymap"/> is <paramref name="chord"/>, or null. Keys that do not parse are skipped.</summary>
    private static string? OwnerOf(KeyChord chord, Dictionary<string, string> keymap)
    {
        foreach (var entry in keymap)
        {
            if (entry.Value.Length == 0)
            {
                continue;
            }

            try
            {
                if (KeyChord.Parse(entry.Value) == chord)
                {
                    return entry.Key;
                }
            }
            catch (FormatException)
            {
                // Validate reports it.
            }
        }

        return null;
    }
}
