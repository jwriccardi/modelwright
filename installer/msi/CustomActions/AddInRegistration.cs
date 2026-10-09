using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace Modelwright.Installer;

/// <summary>One OPEN / OPENn value of Excel's Options key.</summary>
public sealed class OpenEntry
{
    public OpenEntry(string name, int index, string value)
    {
        Name = name;
        Index = index;
        Value = value;
    }

    /// <summary>The value name as stored (OPEN, OPEN1, ...).</summary>
    public string Name { get; }

    /// <summary>The numeric suffix: 0 for OPEN, n for OPENn.</summary>
    public int Index { get; }

    public string Value { get; }
}

/// <summary>One registry change: set <see cref="Name"/> to <see cref="Value"/>, or delete it when Value is null.</summary>
public sealed class RegistryChange
{
    public RegistryChange(string name, string? value)
    {
        Name = name;
        Value = value;
    }

    public string Name { get; }

    public string? Value { get; }

    public bool IsDelete => Value is null;
}

/// <summary>
/// How Modelwright is registered with Excel: the next free OPEN/OPENn value of
/// HKCU\Software\Microsoft\Office\16.0\Excel\Options, holding <c>/R "&lt;path&gt;"</c>, with no gaps in the numbering.
/// The same rules as install/install.ps1 and install/uninstall.ps1 (keep them in step): an existing Modelwright
/// entry (or one from a pre-rename ModelingToolkit build, or the bare form Excel writes for files in its own AddIns
/// folder) is replaced in its slot and any further ones are dropped; if exactly the right entry is already there,
/// nothing changes. Pure logic, so it can be unit-tested; the registry access is in CustomActions.
/// </summary>
public static class AddInRegistration
{
    /// <summary>Excel's per-user Options key, relative to HKEY_CURRENT_USER.</summary>
    public const string OptionsSubKey = @"Software\Microsoft\Office\16.0\Excel\Options";

    /// <summary>Excel's list of add-ins that are known but not ticked, relative to HKEY_CURRENT_USER.</summary>
    public const string AddInManagerSubKey = @"Software\Microsoft\Office\16.0\Excel\Add-in Manager";

    // Same pattern as $ourAddInPattern in install.ps1 / uninstall.ps1 (PowerShell's -match is case-insensitive).
    private static readonly Regex OurAddIn = new Regex(
        @"^\s*(/R\s+)?""?([^""]*[\\/])?(Modelwright(32|64)|ModelingToolkit[^\\/""]*)\.xll""?\s*$",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex OpenName = new Regex(
        @"^OPEN(\d*)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    /// <summary>True when an OPENn value, or an Add-in Manager value name, refers to a Modelwright add-in.</summary>
    public static bool IsOurs(string value) => OurAddIn.IsMatch(value);

    /// <summary>The value an OPENn entry holds for the add-in at <paramref name="xllPath"/>.</summary>
    public static string OpenValue(string xllPath) => "/R \"" + xllPath + "\"";

    /// <summary>OPEN for slot 0, OPENn for slot n.</summary>
    public static string SlotName(int index) =>
        index == 0 ? "OPEN" : "OPEN" + index.ToString(CultureInfo.InvariantCulture);

    /// <summary>The OPEN, OPEN1, ... values among a key's values, sorted by number. Other values are ignored.</summary>
    public static IReadOnlyList<OpenEntry> GetOpenEntries(IEnumerable<KeyValuePair<string, string>> values)
    {
        var entries = new List<OpenEntry>();
        foreach (var pair in values)
        {
            var match = OpenName.Match(pair.Key);
            if (!match.Success)
            {
                continue;
            }

            var digits = match.Groups[1].Value;
            var index = 0;
            if (digits.Length > 0 && !int.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out index))
            {
                index = int.MaxValue;
            }

            entries.Add(new OpenEntry(pair.Key, index, pair.Value));
        }

        return entries.OrderBy(e => e.Index).ToList();
    }

    /// <summary>
    /// The values OPEN, OPEN1, ... should hold so that <paramref name="value"/> is registered: the first Modelwright
    /// entry is replaced in its slot and the others are dropped, or the value is appended. Null when exactly the
    /// right entry is already there and nothing needs to change.
    /// </summary>
    public static IReadOnlyList<string>? PlanRegister(IReadOnlyList<OpenEntry> entries, string value)
    {
        var ours = entries.Where(e => IsOurs(e.Value)).ToList();
        if (ours.Count == 1 && string.Equals(ours[0].Value, value, StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var values = new List<string>();
        var placed = false;
        foreach (var entry in entries)
        {
            if (!IsOurs(entry.Value))
            {
                values.Add(entry.Value);
            }
            else if (!placed)
            {
                values.Add(value);
                placed = true;
            }
        }

        if (!placed)
        {
            values.Add(value);
        }

        return values;
    }

    /// <summary>
    /// The values OPEN, OPEN1, ... should hold once every Modelwright entry is removed (the rest in their order, no
    /// gaps). Null when there is no Modelwright entry.
    /// </summary>
    public static IReadOnlyList<string>? PlanUnregister(IReadOnlyList<OpenEntry> entries)
    {
        if (!entries.Any(e => IsOurs(e.Value)))
        {
            return null;
        }

        return entries.Where(e => !IsOurs(e.Value)).Select(e => e.Value).ToList();
    }

    /// <summary>
    /// The changes that turn <paramref name="entries"/> into OPEN, OPEN1, ... holding <paramref name="values"/> in
    /// order, with no gaps. Values that are already right are left alone. Apply them in the order returned.
    /// </summary>
    public static IReadOnlyList<RegistryChange> Diff(IReadOnlyList<OpenEntry> entries, IReadOnlyList<string> values)
    {
        var current = new Dictionary<string, OpenEntry>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in entries)
        {
            current[entry.Name] = entry;
        }

        var changes = new List<RegistryChange>();
        var wanted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < values.Count; i++)
        {
            var name = SlotName(i);
            wanted.Add(name);
            current.TryGetValue(name, out var old);
            if (old is null || !string.Equals(old.Value, values[i], StringComparison.Ordinal))
            {
                // Registry names are case-insensitive: drop "open1" before writing "OPEN1" so the new name is used.
                if (old is not null && !string.Equals(old.Name, name, StringComparison.Ordinal))
                {
                    changes.Add(new RegistryChange(old.Name, null));
                }

                changes.Add(new RegistryChange(name, values[i]));
            }
        }

        foreach (var entry in entries)
        {
            if (!wanted.Contains(entry.Name))
            {
                changes.Add(new RegistryChange(entry.Name, null));
            }
        }

        return changes;
    }
}
