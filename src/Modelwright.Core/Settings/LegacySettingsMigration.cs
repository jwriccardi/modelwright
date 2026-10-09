using System;
using System.Collections.Generic;
using System.Linq;

namespace Modelwright.Core.Settings;

/// <summary>
/// The one-time move of the per-user files from the working name's folder (<c>%APPDATA%\ModelingToolkit</c>) to the
/// product's (<c>%APPDATA%\Modelwright</c>), decided from which files each folder holds. Files are only ever copied:
/// the legacy folder is left as it is.
/// </summary>
public static class LegacySettingsMigration
{
    /// <summary>The folder name used under <c>%APPDATA%</c> before the rename to Modelwright (D12).</summary>
    public const string LegacyFolderName = "ModelingToolkit";

    /// <summary>The settings file.</summary>
    public const string SettingsFileName = "settings.json";

    /// <summary>The Trace In window's remembered state.</summary>
    public const string UiStateFileName = "ui-state.json";

    /// <summary>
    /// The files to copy from the legacy folder into the current one, given the file names each holds (empty for a
    /// folder that does not exist). Nothing is copied unless the legacy folder has <see cref="SettingsFileName"/> and
    /// the current folder has neither it nor a backup of it (<c>settings.json.*.bak</c>, which the settings file
    /// restores itself from). Then <see cref="SettingsFileName"/> is copied, and <see cref="UiStateFileName"/> too if
    /// the legacy folder has it and the current one does not. Names compare ignoring case, as Windows file names do.
    /// </summary>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    public static IReadOnlyList<string> FilesToCopy(IEnumerable<string> legacyFiles, IEnumerable<string> currentFiles)
    {
        if (legacyFiles is null)
        {
            throw new ArgumentNullException(nameof(legacyFiles));
        }

        if (currentFiles is null)
        {
            throw new ArgumentNullException(nameof(currentFiles));
        }

        var legacy = new HashSet<string>(legacyFiles, StringComparer.OrdinalIgnoreCase);
        var current = new HashSet<string>(currentFiles, StringComparer.OrdinalIgnoreCase);
        if (!legacy.Contains(SettingsFileName) || current.Contains(SettingsFileName) || current.Any(IsSettingsBackup))
        {
            return Array.Empty<string>();
        }

        var copy = new List<string> { SettingsFileName };
        if (legacy.Contains(UiStateFileName) && !current.Contains(UiStateFileName))
        {
            copy.Add(UiStateFileName);
        }

        return copy;
    }

    private static bool IsSettingsBackup(string name) =>
        name.StartsWith(SettingsFileName + ".", StringComparison.OrdinalIgnoreCase)
        && name.EndsWith(".bak", StringComparison.OrdinalIgnoreCase);
}
