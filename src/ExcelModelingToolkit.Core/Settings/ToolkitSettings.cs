using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ExcelModelingToolkit.Core.Formatting;
using ExcelModelingToolkit.Core.Keys;

namespace ExcelModelingToolkit.Core.Settings;

/// <summary>
/// The user's settings (docs/PLAN.md section 4.6), stored as JSON with a versioned schema
/// (<see cref="CurrentSchemaVersion"/>). Immutable. The constructor accepts any non-null contents;
/// <see cref="Validate"/> reports what is wrong with them, and <see cref="FromJson"/> rejects a file that
/// cannot be used (returning <see cref="Defaults"/> with the problems).
/// </summary>
public sealed class ToolkitSettings
{
    /// <summary>The schema version this code reads and writes.</summary>
    public const int CurrentSchemaVersion = 1;

    /// <summary>Default <see cref="UndoCellCap"/>.</summary>
    public const int DefaultUndoCellCap = 10000;

    /// <summary>The largest settings file accepted, in bytes (1 MB). A real file is a few kilobytes.</summary>
    public const long MaxFileBytes = 1024 * 1024;

    private static readonly Encoding StrictUtf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);

    private readonly Dictionary<string, string> _keymap;

    /// <summary>Creates settings.</summary>
    /// <param name="cycles">The cycles, in the order they are written to the file.</param>
    /// <param name="keymap">
    /// Action id (see <see cref="ActionIds"/>) to key string (e.g. <c>Ctrl+Shift+1</c>); an empty string leaves
    /// the action unbound.
    /// </param>
    /// <param name="undoCellCap">
    /// The most format reads one undo snapshot may take (see <see cref="UndoCellCap"/>).
    /// </param>
    /// <param name="diagnosticsLog">True to write the per-command diagnostics log.</param>
    /// <exception cref="ArgumentNullException"><paramref name="cycles"/> or <paramref name="keymap"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="cycles"/> contains null, or <paramref name="keymap"/> has a null or repeated action id,
    /// or a null key.
    /// </exception>
    public ToolkitSettings(
        IEnumerable<CycleDefinition> cycles,
        IEnumerable<KeyValuePair<string, string>> keymap,
        int undoCellCap = DefaultUndoCellCap,
        bool diagnosticsLog = true)
    {
        var cycleList = (cycles ?? throw new ArgumentNullException(nameof(cycles))).ToArray();
        if (cycleList.Any(c => c is null))
        {
            throw new ArgumentException("A cycle cannot be null.", nameof(cycles));
        }

        _keymap = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var entry in keymap ?? throw new ArgumentNullException(nameof(keymap)))
        {
            if (entry.Key is null || entry.Value is null)
            {
                throw new ArgumentException("Keymap action ids and keys cannot be null.", nameof(keymap));
            }

            if (_keymap.ContainsKey(entry.Key))
            {
                throw new ArgumentException($"The keymap lists '{entry.Key}' more than once.", nameof(keymap));
            }

            _keymap.Add(entry.Key, entry.Value);
        }

        Cycles = cycleList;
        UndoCellCap = undoCellCap;
        DiagnosticsLog = diagnosticsLog;
    }

    /// <summary>The schema version (always <see cref="CurrentSchemaVersion"/> for a loaded or created instance).</summary>
    public int SchemaVersion => CurrentSchemaVersion;

    /// <summary>The cycles.</summary>
    public IReadOnlyList<CycleDefinition> Cycles { get; }

    /// <summary>Action id to key string; an empty string means unbound.</summary>
    public IReadOnlyDictionary<string, string> Keymap => _keymap;

    /// <summary>
    /// The most format reads one undo snapshot may take (<see cref="Undo.SnapshotPlanner"/>); it caps reads, not
    /// cells. A uniform region of any size costs one read (a whole column too); a mixed one is halved until each
    /// piece is uniform, so cells whose formats all differ cost up to about two reads each, and k uniform regions
    /// about k * log2(cells). Beyond the cap, the change is applied but cannot be undone (Ctrl+Z stops there and
    /// says so).
    /// </summary>
    public int UndoCellCap { get; }

    /// <summary>True to write the per-command diagnostics log.</summary>
    public bool DiagnosticsLog { get; }

    /// <summary>The factory defaults (the owner's Macabacus configuration); see <see cref="DefaultSettings"/>.</summary>
    public static ToolkitSettings Defaults() => DefaultSettings.Create();

    /// <summary>
    /// Parses settings JSON. If the JSON is malformed, has an unsupported <c>schemaVersion</c>, or fails
    /// <see cref="Validate"/>, returns a <see cref="SettingsLoadOutcome.Rejected"/> result carrying
    /// <see cref="Defaults"/> and the problems (never throws for bad content).
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="json"/> is null.</exception>
    public static SettingsLoadResult FromJson(string json)
    {
        if (json is null)
        {
            throw new ArgumentNullException(nameof(json));
        }

        var problems = new List<string>();
        var settings = SettingsJson.Read(json, problems);
        if (settings is not null)
        {
            problems.AddRange(settings.Validate());
        }

        return problems.Count == 0
            ? new SettingsLoadResult(settings!, problems, SettingsLoadOutcome.Loaded)
            : SettingsLoadResult.Rejected(problems);
    }

    /// <summary>
    /// Reads the settings file's raw bytes: rejects a file larger than <see cref="MaxFileBytes"/>, decodes it as
    /// strict UTF-8 (a UTF-8 byte order mark is skipped; a UTF-16 one selects strict UTF-16), then parses it with
    /// <see cref="FromJson"/>. Invalid bytes are rejected rather than replaced, so a file saved in another code
    /// page cannot silently load with garbled number formats. Never throws for bad content.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="bytes"/> is null.</exception>
    public static SettingsLoadResult FromFileBytes(byte[] bytes)
    {
        if (bytes is null)
        {
            throw new ArgumentNullException(nameof(bytes));
        }

        var tooLarge = CheckFileSize(bytes.LongLength);
        if (tooLarge is not null)
        {
            return SettingsLoadResult.Rejected(new[] { tooLarge });
        }

        string json;
        try
        {
            json = Decode(bytes);
        }
        catch (DecoderFallbackException)
        {
            return SettingsLoadResult.Rejected(new[] { InvalidEncodingProblem(bytes) });
        }

        return FromJson(json);
    }

    /// <summary>
    /// The problem to report for a settings file of <paramref name="length"/> bytes, or null if the size is
    /// acceptable (at most <see cref="MaxFileBytes"/>). Lets the caller check before reading the file.
    /// </summary>
    public static string? CheckFileSize(long length) =>
        length > MaxFileBytes
            ? string.Format(
                CultureInfo.InvariantCulture,
                "settings.json is {0:N0} bytes; the limit is {1:N0} bytes (1 MB). It is probably not a settings file.",
                length,
                MaxFileBytes)
            : null;

    /// <summary>Writes the settings as indented, human-editable JSON that <see cref="FromJson"/> reads back.</summary>
    public string ToJson() => SettingsJson.Write(this);

    /// <summary>The cycle with id <paramref name="id"/> (ordinal), or null.</summary>
    public CycleDefinition? FindCycle(string id) =>
        Cycles.FirstOrDefault(c => string.Equals(c.Id, id, StringComparison.Ordinal));

    /// <summary>
    /// Checks the settings and returns the problems found (empty if none): an undo cap below 1; each cycle's own
    /// problems (<see cref="CycleDefinition.Validate"/>); duplicate cycle ids (ignoring case); keymap entries for
    /// unknown actions, keys that <see cref="KeyChord.Parse"/> rejects, a key bound to two actions, and a bound
    /// cycle action with no cycle of that id.
    /// </summary>
    public IReadOnlyList<string> Validate()
    {
        var problems = new List<string>();
        if (UndoCellCap < 1)
        {
            problems.Add($"undoCellCap is {UndoCellCap}; it must be at least 1.");
        }

        var seenIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var cycle in Cycles)
        {
            problems.AddRange(cycle.Validate());
            if (!string.IsNullOrWhiteSpace(cycle.Id) && !seenIds.Add(cycle.Id))
            {
                problems.Add($"cycle '{cycle.Id}': another cycle has the same id (ids are compared ignoring case).");
            }
        }

        var chordOwners = new Dictionary<KeyChord, string>();
        foreach (var actionId in OrderedKeymapActions())
        {
            var key = _keymap[actionId];
            if (!ActionIds.IsKnown(actionId))
            {
                problems.Add($"keymap: '{actionId}' is not an action. Known actions: {string.Join(", ", ActionIds.All)}.");
                continue;
            }

            if (ActionIds.IsCycle(actionId) && key.Length > 0 && FindCycle(actionId) is null)
            {
                problems.Add($"keymap: '{actionId}' has a key but there is no cycle with id '{actionId}'.");
            }

            if (key.Length == 0)
            {
                continue;
            }

            KeyChord chord;
            try
            {
                chord = KeyChord.Parse(key);
            }
            catch (FormatException ex)
            {
                problems.Add($"keymap: '{actionId}': {ex.Message}");
                continue;
            }

            if (chordOwners.TryGetValue(chord, out var owner))
            {
                problems.Add($"keymap: {chord} is assigned to both '{owner}' and '{actionId}'.");
            }
            else
            {
                chordOwners.Add(chord, actionId);
            }
        }

        return problems;
    }

    private static string Decode(byte[] bytes)
    {
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
        {
            return StrictUtf8.GetString(bytes, 3, bytes.Length - 3);
        }

        if (IsUtf16Bom(bytes, out var bigEndian))
        {
            var utf16 = new UnicodeEncoding(bigEndian, byteOrderMark: false, throwOnInvalidBytes: true);
            return utf16.GetString(bytes, 2, bytes.Length - 2);
        }

        return StrictUtf8.GetString(bytes);
    }

    private static bool IsUtf16Bom(byte[] bytes, out bool bigEndian)
    {
        bigEndian = bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF;
        return bigEndian || (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE);
    }

    private static string InvalidEncodingProblem(byte[] bytes) =>
        IsUtf16Bom(bytes, out _)
            ? "settings.json is not valid UTF-16; save it as UTF-8."
            : "settings.json is not valid UTF-8; save it as UTF-8.";

    /// <summary>The keymap's action ids: known actions in <see cref="ActionIds.All"/> order, then any others ordinally.</summary>
    internal IEnumerable<string> OrderedKeymapActions() =>
        _keymap.Keys.OrderBy(OrderOf).ThenBy(k => k, StringComparer.Ordinal);

    private static int OrderOf(string actionId)
    {
        for (var i = 0; i < ActionIds.All.Count; i++)
        {
            if (string.Equals(ActionIds.All[i], actionId, StringComparison.Ordinal))
            {
                return i;
            }
        }

        return int.MaxValue;
    }
}
