using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelModelingToolkit.Core.Settings;

/// <summary>The outcome of loading settings: the settings to use and any problems found.</summary>
public sealed class SettingsLoadResult
{
    /// <summary>Creates a result.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="settings"/> or <paramref name="problems"/> is null.</exception>
    public SettingsLoadResult(ToolkitSettings settings, IEnumerable<string> problems, SettingsLoadOutcome outcome)
    {
        Settings = settings ?? throw new ArgumentNullException(nameof(settings));
        Problems = (problems ?? throw new ArgumentNullException(nameof(problems))).ToArray();
        Outcome = outcome;
    }

    /// <summary>
    /// The settings to use: the loaded ones, or the defaults if there was no file or it was rejected. Never null.
    /// </summary>
    public ToolkitSettings Settings { get; }

    /// <summary>Human-readable problems (empty if the file loaded cleanly).</summary>
    public IReadOnlyList<string> Problems { get; }

    /// <summary>Whether the file was loaded, created with the defaults, or rejected.</summary>
    public SettingsLoadOutcome Outcome { get; }

    /// <summary>A <see cref="SettingsLoadOutcome.Rejected"/> result carrying the defaults and <paramref name="problems"/>.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="problems"/> is null.</exception>
    public static SettingsLoadResult Rejected(IEnumerable<string> problems) =>
        new SettingsLoadResult(ToolkitSettings.Defaults(), problems, SettingsLoadOutcome.Rejected);
}
