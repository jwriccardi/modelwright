using System;
using System.Collections.Generic;
using System.Linq;
using ExcelModelingToolkit.Core.Keys;

namespace ExcelModelingToolkit.Core.Settings;

/// <summary>
/// The settings dialog's editing model (docs/PLAN.md section 4.6): a mutable copy of <see cref="ToolkitSettings"/>
/// whose cycles' items (<see cref="CycleDraft"/>), keymap, undo cap and diagnostics setting can be changed, reset
/// to the defaults, exported and imported. Contents may be invalid while editing; <see cref="ToSettings"/> builds
/// the immutable settings and reports their problems with <see cref="ToolkitSettings.Validate"/>, the same rules
/// that check the settings file.
/// </summary>
public sealed class SettingsDraft
{
    private readonly List<CycleDraft> _cycles = new List<CycleDraft>();
    private readonly Dictionary<string, string> _keymap = new Dictionary<string, string>(StringComparer.Ordinal);

    /// <summary>
    /// Creates a draft holding a copy of <paramref name="settings"/>, with each key that <see cref="KeyChord.Parse"/>
    /// accepts in its display form, as <see cref="SetKey"/> stores it.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="settings"/> is null.</exception>
    public SettingsDraft(ToolkitSettings settings)
    {
        Load(settings ?? throw new ArgumentNullException(nameof(settings)));
    }

    /// <summary>The cycles, in file order. <see cref="ResetToDefaults"/> and a successful import replace them.</summary>
    public IReadOnlyList<CycleDraft> Cycles => _cycles;

    /// <summary>See <see cref="ToolkitSettings.UndoCellCap"/>. A value below 1 is reported by validation.</summary>
    public int UndoCellCap { get; set; }

    /// <summary>See <see cref="ToolkitSettings.DiagnosticsLog"/>.</summary>
    public bool DiagnosticsLog { get; set; }

    /// <summary>
    /// The actions the shortcuts editor lists: every action in <see cref="ActionIds.All"/> order, then any other
    /// action the keymap holds (ordinal order).
    /// </summary>
    public IReadOnlyList<string> ShortcutActions =>
        ActionIds.All
            .Concat(_keymap.Keys.Where(a => !ActionIds.IsKnown(a)).OrderBy(a => a, StringComparer.Ordinal))
            .ToArray();

    /// <summary>The key string bound to <paramref name="actionId"/>, or an empty string if it is unbound.</summary>
    public string GetKey(string actionId) =>
        _keymap.TryGetValue(actionId ?? throw new ArgumentNullException(nameof(actionId)), out var key) ? key : string.Empty;

    /// <summary>
    /// Binds <paramref name="key"/> (trimmed) to <paramref name="actionId"/>; an empty key leaves it unbound. A key
    /// that <see cref="KeyChord.Parse"/> accepts is stored in its display form (e.g. <c>ctrl+shift+k</c> becomes
    /// <c>Ctrl+Shift+K</c>); any other text is stored as typed, so <see cref="KeyProblem"/> and validation report it.
    /// </summary>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="actionId"/> is not one of <see cref="ShortcutActions"/>.</exception>
    public void SetKey(string actionId, string key)
    {
        if (actionId is null)
        {
            throw new ArgumentNullException(nameof(actionId));
        }

        if (key is null)
        {
            throw new ArgumentNullException(nameof(key));
        }

        if (!ActionIds.IsKnown(actionId) && !_keymap.ContainsKey(actionId))
        {
            throw new ArgumentException($"'{actionId}' is not an action.", nameof(actionId));
        }

        _keymap[actionId] = Canonical(key.Trim());
    }

    /// <summary>Leaves <paramref name="actionId"/> unbound (an empty key string, which the file keeps).</summary>
    /// <exception cref="ArgumentNullException"><paramref name="actionId"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="actionId"/> is not one of <see cref="ShortcutActions"/>.</exception>
    public void ClearKey(string actionId) => SetKey(actionId, string.Empty);

    /// <summary>
    /// What is wrong with <paramref name="actionId"/>'s key, for showing next to it while editing, or null if
    /// nothing is (an unbound action is fine): the <see cref="KeyChord.Parse"/> error, the other action that has
    /// the same key, or the other action whose key is pressed the same way on a US keyboard (e.g. <c>Ctrl+{</c> and
    /// <c>Ctrl+Shift+[</c>; see <see cref="KeyChord.ToUsKeys"/>).
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="actionId"/> is null.</exception>
    public string? KeyProblem(string actionId)
    {
        var key = GetKey(actionId);
        if (key.Length == 0)
        {
            return null;
        }

        if (!TryParse(key, out var chord, out var error))
        {
            return error;
        }

        string? sameUsKeys = null;
        foreach (var other in ShortcutActions)
        {
            if (string.Equals(other, actionId, StringComparison.Ordinal) || !TryParse(GetKey(other), out var otherChord, out _))
            {
                continue;
            }

            if (otherChord == chord)
            {
                return $"{chord} is also assigned to {ActionDisplayName(other)}.";
            }

            if (sameUsKeys is null && otherChord!.ToUsKeys() == chord!.ToUsKeys())
            {
                sameUsKeys = $"{chord} is pressed with the same keys as {otherChord} ({ActionDisplayName(other)}) on a US keyboard.";
            }
        }

        return sameUsKeys;
    }

    /// <summary>
    /// The name people see for <paramref name="actionId"/>: its cycle's display name, else
    /// <see cref="ActionIds.DisplayName"/> (<c>Trace In</c>, <c>About</c>, or the id itself).
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="actionId"/> is null.</exception>
    public string ActionDisplayName(string actionId)
    {
        if (actionId is null)
        {
            throw new ArgumentNullException(nameof(actionId));
        }

        var cycle = _cycles.FirstOrDefault(c => string.Equals(c.Id, actionId, StringComparison.Ordinal));
        if (cycle is not null)
        {
            return cycle.DisplayName;
        }

        return ActionIds.DisplayName(actionId);
    }

    /// <summary>
    /// The settings the draft now holds, and in <paramref name="problems"/> everything
    /// <see cref="ToolkitSettings.Validate"/> finds wrong with them (empty when they can be saved).
    /// </summary>
    public ToolkitSettings ToSettings(out IReadOnlyList<string> problems)
    {
        var settings = new ToolkitSettings(
            _cycles.Select(c => c.ToDefinition()),
            _keymap,
            UndoCellCap,
            DiagnosticsLog);
        problems = settings.Validate();
        return settings;
    }

    /// <summary>
    /// The draft as settings-file JSON (<see cref="ToolkitSettings.ToJson"/>, the serializer that writes
    /// settings.json), or null if the draft is invalid, with the problems in <paramref name="problems"/>. An exported
    /// file imports back (<see cref="ImportJson"/>) to identical JSON.
    /// </summary>
    public string? ExportJson(out IReadOnlyList<string> problems)
    {
        var settings = ToSettings(out problems);
        return problems.Count == 0 ? settings.ToJson() : null;
    }

    /// <summary>
    /// Replaces the draft's contents with settings JSON, read and validated exactly as the settings file is
    /// (<see cref="ToolkitSettings.FromJson"/>). Returns the problems; if there are any, the draft is unchanged.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="json"/> is null.</exception>
    public IReadOnlyList<string> ImportJson(string json) => Import(ToolkitSettings.FromJson(json));

    /// <summary>
    /// Like <see cref="ImportJson"/>, from a file's raw bytes: size and encoding are checked as for the settings
    /// file (<see cref="ToolkitSettings.FromFileBytes(byte[], string)"/>), and those problems name
    /// <paramref name="fileName"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    public IReadOnlyList<string> ImportFile(byte[] bytes, string fileName = ToolkitSettings.DefaultFileName) =>
        Import(ToolkitSettings.FromFileBytes(bytes, fileName));

    /// <summary>
    /// The draft as settings-file JSON whatever its validity, for telling whether anything was edited (two drafts
    /// with the same contents give the same text).
    /// </summary>
    public string ContentJson() => ToSettings(out _).ToJson();

    /// <summary>Replaces the draft's contents with the factory defaults (<see cref="ToolkitSettings.Defaults"/>).</summary>
    public void ResetToDefaults() => Load(ToolkitSettings.Defaults());

    private IReadOnlyList<string> Import(SettingsLoadResult result)
    {
        if (result.Outcome != SettingsLoadOutcome.Loaded)
        {
            return result.Problems;
        }

        Load(result.Settings);
        return Array.Empty<string>();
    }

    private void Load(ToolkitSettings settings)
    {
        _cycles.Clear();
        _cycles.AddRange(settings.Cycles.Select(c => new CycleDraft(c)));
        _keymap.Clear();
        foreach (var entry in settings.Keymap)
        {
            _keymap.Add(entry.Key, Canonical(entry.Value)); // As SetKey stores it, so re-entering a key changes nothing.
        }

        UndoCellCap = settings.UndoCellCap;
        DiagnosticsLog = settings.DiagnosticsLog;
    }

    /// <summary>A key that <see cref="KeyChord.Parse"/> accepts in its display form; any other text unchanged.</summary>
    private static string Canonical(string key) => TryParse(key, out var chord, out _) ? chord!.ToDisplayString() : key;

    private static bool TryParse(string key, out KeyChord? chord, out string? error)
    {
        chord = null;
        error = null;
        if (key.Length == 0)
        {
            return false;
        }

        try
        {
            chord = KeyChord.Parse(key);
            return true;
        }
        catch (FormatException ex)
        {
            error = ex.Message;
            return false;
        }
    }
}
