using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelModelingToolkit.Core.Settings;

/// <summary>The outcome of loading settings: the settings to use and any problems found.</summary>
public sealed class SettingsLoadResult
{
    /// <summary>Creates a result.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="settings"/> or <paramref name="problems"/> is null.</exception>
    public SettingsLoadResult(ToolkitSettings settings, IEnumerable<string> problems, bool usedDefaults)
    {
        Settings = settings ?? throw new ArgumentNullException(nameof(settings));
        Problems = (problems ?? throw new ArgumentNullException(nameof(problems))).ToArray();
        UsedDefaults = usedDefaults;
    }

    /// <summary>The settings to use: the loaded ones, or the defaults if the file was rejected.</summary>
    public ToolkitSettings Settings { get; }

    /// <summary>Human-readable problems (empty if the file loaded cleanly).</summary>
    public IReadOnlyList<string> Problems { get; }

    /// <summary>True when the file could not be used and <see cref="Settings"/> are the defaults.</summary>
    public bool UsedDefaults { get; }
}
