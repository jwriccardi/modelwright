using System;
using System.IO;
using ExcelModelingToolkit.Core.Settings;

namespace ExcelModelingToolkit.AddIn;

/// <summary>
/// The settings file, <c>%APPDATA%\ModelingToolkit\settings.json</c> (UTF-8, human-editable), with its timestamped
/// backups beside it. The file handling (atomic save, backups, restoring a missing file, change detection) is
/// <see cref="SettingsFile"/>; no member throws.
/// </summary>
internal static class SettingsStore
{
    /// <summary>How old a leftover temporary file must be before startup deletes it.</summary>
    private static readonly TimeSpan StaleTemporaryFileAge = TimeSpan.FromDays(1);

    private static readonly SettingsFile File = new SettingsFile(Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        ProductInfo.FolderName,
        "settings.json"));

    /// <summary>Full path of the settings file.</summary>
    public static string FilePath => File.FilePath;

    /// <summary>
    /// Loads the settings file; a missing one is restored from the newest backup, or else created with the defaults.
    /// See <see cref="SettingsFile.Load"/>.
    /// </summary>
    public static SettingsFileLoadResult Load() => File.Load();

    /// <summary>Restores or creates the file if it is missing. Returns null on success, else the reason.</summary>
    public static string? EnsureExists() => File.EnsureExists();

    /// <summary>Saves <paramref name="settings"/> atomically, keeping a backup of the replaced file. See <see cref="SettingsFile.Save"/>.</summary>
    public static SettingsFileSaveResult Save(ToolkitSettings settings) => File.Save(settings);

    /// <summary>True if the file exists and is not the one whose hash is <paramref name="knownHash"/> (or cannot be read).</summary>
    public static bool HasChangedSince(string? knownHash) => File.HasChangedSince(knownHash);

    /// <summary>True if the settings file exists.</summary>
    public static bool Exists() => System.IO.File.Exists(File.FilePath);

    /// <summary>Deletes temporary files a save left behind more than a day ago. Returns how many.</summary>
    public static int DeleteStaleTemporaryFiles() => File.DeleteStaleTemporaryFiles(StaleTemporaryFileAge);
}
