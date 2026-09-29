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
    /// <summary>The settings in use (the defaults until <see cref="LoadSettings"/> runs).</summary>
    public static ToolkitSettings Settings { get; private set; } = ToolkitSettings.Defaults();

    /// <summary>The cycle engine; replaced whenever the settings are (re)loaded.</summary>
    public static CycleEngine Engine { get; private set; } = new CycleEngine();

    /// <summary>Each cycle's state from its previous press, by cycle id.</summary>
    public static Dictionary<string, CycleState> States { get; } = new Dictionary<string, CycleState>(StringComparer.Ordinal);

    /// <summary>
    /// Loads the settings file (writing the defaults on first run), resets the engine and cycle states, applies the
    /// diagnostics setting, and logs any problems. Never throws.
    /// </summary>
    public static SettingsLoadResult LoadSettings()
    {
        var result = SettingsStore.Load();
        Settings = result.Settings;
        Engine = new CycleEngine();
        States.Clear();

        DiagnosticsLog.Enabled = Settings.DiagnosticsLog;
        DiagnosticsLog.Write(
            "Settings",
            SettingsStore.FilePath,
            result.UsedDefaults ? "using defaults" : "loaded",
            $"problems={result.Problems.Count}");
        foreach (var problem in result.Problems)
        {
            DiagnosticsLog.Write("SettingsProblem", problem);
        }

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
            message += $". Settings problem{(load.UsedDefaults ? ", using defaults" : string.Empty)}: " +
                load.Problems[0] + More(load.Problems.Count);
        }

        if (keyFailures.Count > 0)
        {
            message += ". Could not register " + keyFailures[0] + More(keyFailures.Count);
        }

        return message;
    }

    private static string More(int count) =>
        count > 1 ? $" (+{count - 1} more in the diagnostics log)" : string.Empty;
}
