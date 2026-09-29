using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using ExcelModelingToolkit.Core.Formatting;
using ExcelModelingToolkit.Core.Settings;
using ExcelModelingToolkit.Core.Undo;

namespace ExcelModelingToolkit.AddIn;

/// <summary>
/// In-memory state for the Excel session: the settings in use, the cycle engine (with its learned number format
/// aliases), each cycle's last <see cref="CycleState"/> and our formatting undo stack. Main thread only; nothing
/// here is persisted.
/// </summary>
internal static class Session
{
    /// <summary>The settings in use (the defaults until <see cref="Initialize"/> runs).</summary>
    public static ToolkitSettings Settings { get; private set; } = ToolkitSettings.Defaults();

    /// <summary>The cycle engine; kept across reloads, with the aliases of codes that are gone dropped.</summary>
    public static CycleEngine Engine { get; private set; } = new CycleEngine();

    /// <summary>Each cycle's state from its previous press, by cycle id.</summary>
    public static Dictionary<string, CycleState> States { get; } = new Dictionary<string, CycleState>(StringComparer.Ordinal);

    /// <summary>
    /// Our undo and redo stacks for the cycles' changes (docs/PLAN.md section 4.4). Kept across settings reloads;
    /// emptied when the add-in closes.
    /// </summary>
    public static UndoManager Undo { get; } = new UndoManager();

    /// <summary>
    /// Whether the settings file is what <see cref="Settings"/> came from: <see cref="SettingsLoadOutcome.Loaded"/>
    /// (loaded, or saved by the settings dialog) or <see cref="SettingsLoadOutcome.CreatedDefaults"/> (written with
    /// the defaults); or <see cref="SettingsLoadOutcome.Rejected"/>: the file has problems and is not in use, so
    /// saving over it must be confirmed.
    /// </summary>
    public static SettingsLoadOutcome SourceState { get; private set; } = SettingsLoadOutcome.Rejected;

    /// <summary>
    /// The SHA-256 of the settings file as we last loaded, created or saved it, or null if the file is not in use
    /// (<see cref="SourceState"/>) or could not be written. A file whose hash differs was changed by someone else.
    /// </summary>
    public static string? SourceHash { get; private set; }

    /// <summary>
    /// Loads the settings file at startup (restoring the newest backup if it is missing, else writing the defaults
    /// on first run; using the defaults if the file is rejected), after deleting temporary files that earlier saves
    /// left behind; starts a fresh engine and cycle states, applies the diagnostics setting, and logs any problems.
    /// Never throws.
    /// </summary>
    public static SettingsFileLoadResult Initialize()
    {
        var deleted = SettingsStore.DeleteStaleTemporaryFiles();
        var load = SettingsStore.Load();
        var result = load.Result;
        Settings = result.Settings;
        Engine = new CycleEngine();
        States.Clear();
        UseSource(load);
        DiagnosticsLog.Enabled = Settings.DiagnosticsLog;
        if (deleted > 0)
        {
            DiagnosticsLog.Write("SettingsTempFilesDeleted", deleted.ToString(CultureInfo.InvariantCulture));
        }

        LogLoad(load, result.Outcome == SettingsLoadOutcome.Rejected ? "rejected, using defaults" : Describe(result.Outcome));
        return load;
    }

    /// <summary>
    /// Reloads the settings file. If it is rejected, nothing changes (settings, engine, cycle states and the
    /// diagnostics setting stay as they were; the file is marked not in use) and <paramref name="keptPrevious"/> is
    /// true. Otherwise the new settings are used; the engine keeps its aliases for number format codes that are
    /// still in their cycle, and each cycle keeps its last state if its items are unchanged. A missing file is
    /// restored from the newest backup, or with no backup recreated with the defaults. Never throws.
    /// </summary>
    public static SettingsFileLoadResult Reload(out bool keptPrevious)
    {
        var load = SettingsStore.Load();
        UseSource(load);
        keptPrevious = load.Result.Outcome == SettingsLoadOutcome.Rejected;
        if (keptPrevious)
        {
            LogLoad(load, "Settings rejected, kept previous");
            return load;
        }

        Use(load.Result.Settings);
        LogLoad(load, Describe(load.Result.Outcome));
        return load;
    }

    /// <summary>
    /// Uses <paramref name="settings"/>, just saved to the settings file by the settings dialog
    /// (<paramref name="save"/>), exactly as a successful <see cref="Reload"/> would: the engine keeps its aliases
    /// for number format codes still in their cycle, and each cycle keeps its last state if its items are unchanged.
    /// The saved file is now the one in use. Logs the save first, while the previous diagnostics setting still
    /// applies, so turning the log off is itself logged. Never throws.
    /// </summary>
    public static void ApplySaved(ToolkitSettings settings, SettingsFileSaveResult save)
    {
        DiagnosticsLog.Write(
            "SettingsSaved",
            SettingsStore.FilePath,
            "source=dialog",
            "cycles=" + settings.Cycles.Count.ToString(CultureInfo.InvariantCulture),
            "undoCellCap=" + settings.UndoCellCap.ToString(CultureInfo.InvariantCulture),
            "diagnosticsLog=" + (settings.DiagnosticsLog ? "true" : "false"),
            "backup=" + (save.BackupPath ?? "none"));
        SourceState = SettingsLoadOutcome.Loaded;
        SourceHash = save.Hash;
        Use(settings);
    }

    /// <summary>
    /// A status-bar message: <paramref name="head"/>, then the backup a missing settings file was restored from,
    /// the first settings problem and the first key that could not be bound, if any (the rest are in the
    /// diagnostics log).
    /// </summary>
    public static string Summarize(string head, SettingsFileLoadResult? fileLoad, IReadOnlyList<string> keyFailures)
    {
        var message = head;
        if (fileLoad?.RestoredFrom is not null)
        {
            message += $". settings.json was missing; restored it from {Path.GetFileName(fileLoad.RestoredFrom)}";
        }

        var load = fileLoad?.Result;
        if (load is not null && load.Problems.Count > 0)
        {
            message += $". Settings problem{(load.Outcome == SettingsLoadOutcome.Rejected ? ", using defaults" : string.Empty)}: " +
                load.Problems[0] + More(load.Problems.Count);
        }

        if (keyFailures.Count > 0)
        {
            message += ". Could not register " + keyFailures[0] + More(keyFailures.Count);
        }

        return message;
    }

    /// <summary>Switches to <paramref name="settings"/>, keeping what still applies (see <see cref="Reload"/>).</summary>
    private static void Use(ToolkitSettings settings)
    {
        var previous = Settings;
        Settings = settings;
        Engine.RetainAliases(Settings.Cycles);
        foreach (var cycleId in new List<string>(States.Keys))
        {
            if (previous.FindCycle(cycleId)?.HasSameItems(Settings.FindCycle(cycleId)) != true)
            {
                States.Remove(cycleId);
            }
        }

        DiagnosticsLog.Enabled = Settings.DiagnosticsLog;
    }

    /// <summary>Records whether the file just loaded is in use, and its hash if it is.</summary>
    private static void UseSource(SettingsFileLoadResult load)
    {
        SourceState = load.Result.Outcome;
        SourceHash = load.Result.Outcome == SettingsLoadOutcome.Rejected ? null : load.Hash;
    }

    private static void LogLoad(SettingsFileLoadResult load, string outcome)
    {
        if (load.RestoredFrom is not null)
        {
            DiagnosticsLog.Write("SettingsRestoredFromBackup", load.RestoredFrom);
        }

        DiagnosticsLog.Write("Settings", SettingsStore.FilePath, outcome, $"problems={load.Result.Problems.Count}");
        foreach (var problem in load.Result.Problems)
        {
            DiagnosticsLog.Write("SettingsProblem", problem);
        }
    }

    private static string Describe(SettingsLoadOutcome outcome) =>
        outcome == SettingsLoadOutcome.CreatedDefaults ? "created defaults" : "loaded";

    private static string More(int count) =>
        count > 1 ? $" (+{count - 1} more in the diagnostics log)" : string.Empty;
}
