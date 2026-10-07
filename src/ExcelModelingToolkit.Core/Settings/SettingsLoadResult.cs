using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelModelingToolkit.Core.Settings;

/// <summary>The outcome of loading settings: the settings to use and any problems found.</summary>
public sealed class SettingsLoadResult
{
    /// <summary>Creates a result.</summary>
    /// <param name="settings">The settings to use.</param>
    /// <param name="problems">The problems found.</param>
    /// <param name="outcome">Whether the file was loaded, created with the defaults, or rejected.</param>
    /// <param name="notes">What was brought up to date when the file was loaded (see <see cref="Notes"/>); null for none.</param>
    /// <exception cref="ArgumentNullException"><paramref name="settings"/> or <paramref name="problems"/> is null.</exception>
    public SettingsLoadResult(
        ToolkitSettings settings,
        IEnumerable<string> problems,
        SettingsLoadOutcome outcome,
        IEnumerable<string>? notes = null)
    {
        Settings = settings ?? throw new ArgumentNullException(nameof(settings));
        Problems = (problems ?? throw new ArgumentNullException(nameof(problems))).ToArray();
        Outcome = outcome;
        Notes = notes?.ToArray() ?? Array.Empty<string>();
    }

    /// <summary>
    /// The settings to use: the loaded ones, or the defaults if there was no file or it was rejected. Never null.
    /// </summary>
    public ToolkitSettings Settings { get; }

    /// <summary>Human-readable problems (empty if the file loaded cleanly).</summary>
    public IReadOnlyList<string> Problems { get; }

    /// <summary>
    /// Informational, not problems: what <see cref="Settings"/> add to a file written by an earlier build (actions and
    /// cycles added since, and default lists in place of provisional placeholders; see
    /// <see cref="ToolkitSettings.FromJson"/>). The file itself is unchanged. Empty if nothing was added.
    /// </summary>
    public IReadOnlyList<string> Notes { get; }

    /// <summary>Whether the file was loaded, created with the defaults, or rejected.</summary>
    public SettingsLoadOutcome Outcome { get; }

    /// <summary>A <see cref="SettingsLoadOutcome.Rejected"/> result carrying the defaults and <paramref name="problems"/>.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="problems"/> is null.</exception>
    public static SettingsLoadResult Rejected(IEnumerable<string> problems) =>
        new SettingsLoadResult(ToolkitSettings.Defaults(), problems, SettingsLoadOutcome.Rejected);
}
