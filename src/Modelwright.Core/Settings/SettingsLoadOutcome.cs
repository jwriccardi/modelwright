namespace Modelwright.Core.Settings;

/// <summary>What happened when settings were loaded.</summary>
public enum SettingsLoadOutcome
{
    /// <summary>The file was read and is valid; <see cref="SettingsLoadResult.Settings"/> are its contents.</summary>
    Loaded,

    /// <summary>
    /// There was no file, so the defaults are used (and were written as a new file, unless
    /// <see cref="SettingsLoadResult.Problems"/> says why not).
    /// </summary>
    CreatedDefaults,

    /// <summary>
    /// The file could not be read or is invalid (see <see cref="SettingsLoadResult.Problems"/>). The file is left
    /// untouched and <see cref="SettingsLoadResult.Settings"/> are the defaults; a caller that already has settings
    /// may keep those instead.
    /// </summary>
    Rejected,
}
