using System;
using System.Collections.Generic;
using ExcelModelingToolkit.Core.Formatting;
using ExcelModelingToolkit.Core.Settings;

namespace ExcelModelingToolkit.AddIn;

/// <summary>
/// In-memory state for the Excel session: the settings in use, the cycle engine (with its learned number format
/// aliases) and each cycle's last <see cref="CycleState"/>. Main thread only; nothing here is persisted.
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
    /// Loads the settings file at startup (writing the defaults on first run; using them if the file is rejected),
    /// starts a fresh engine and cycle states, applies the diagnostics setting, and logs any problems. Never throws.
    /// </summary>
    public static SettingsLoadResult Initialize()
    {
        var result = SettingsStore.Load();
        Settings = result.Settings;
        Engine = new CycleEngine();
        States.Clear();
        DiagnosticsLog.Enabled = Settings.DiagnosticsLog;
        LogLoad(result, result.Outcome == SettingsLoadOutcome.Rejected ? "rejected, using defaults" : Describe(result.Outcome));
        return result;
    }

    /// <summary>
    /// Reloads the settings file. If it is rejected, nothing changes (settings, engine, cycle states and the
    /// diagnostics setting stay as they were) and <paramref name="keptPrevious"/> is true. Otherwise the new
    /// settings are used; the engine keeps its aliases for number format codes that are still in their cycle, and
    /// each cycle keeps its last state if its items are unchanged. A missing file is recreated with the defaults
    /// (a deliberate reset). Never throws.
    /// </summary>
    public static SettingsLoadResult Reload(out bool keptPrevious)
    {
        var result = SettingsStore.Load();
        keptPrevious = result.Outcome == SettingsLoadOutcome.Rejected;
        if (keptPrevious)
        {
            LogLoad(result, "Settings rejected, kept previous");
            return result;
        }

        var previous = Settings;
        Settings = result.Settings;
        Engine.RetainAliases(Settings.Cycles);
        foreach (var cycleId in new List<string>(States.Keys))
        {
            if (previous.FindCycle(cycleId)?.HasSameItems(Settings.FindCycle(cycleId)) != true)
            {
                States.Remove(cycleId);
            }
        }

        DiagnosticsLog.Enabled = Settings.DiagnosticsLog;
        LogLoad(result, Describe(result.Outcome));
        return result;
    }

    /// <summary>
    /// A status-bar message: <paramref name="head"/>, then the first settings problem and the first key that could
    /// not be bound, if any (the rest are in the diagnostics log).
    /// </summary>
    public static string Summarize(string head, SettingsLoadResult? load, IReadOnlyList<string> keyFailures)
    {
        var message = head;
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

    private static void LogLoad(SettingsLoadResult result, string outcome)
    {
        DiagnosticsLog.Write("Settings", SettingsStore.FilePath, outcome, $"problems={result.Problems.Count}");
        foreach (var problem in result.Problems)
        {
            DiagnosticsLog.Write("SettingsProblem", problem);
        }
    }

    private static string Describe(SettingsLoadOutcome outcome) =>
        outcome == SettingsLoadOutcome.CreatedDefaults ? "created defaults" : "loaded";

    private static string More(int count) =>
        count > 1 ? $" (+{count - 1} more in the diagnostics log)" : string.Empty;
}
