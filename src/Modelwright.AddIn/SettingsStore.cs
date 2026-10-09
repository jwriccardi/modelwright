using System;
using System.IO;
using Modelwright.Core.Settings;

namespace Modelwright.AddIn;

/// <summary>
/// The settings file, <c>%APPDATA%\Modelwright\settings.json</c> (UTF-8, human-editable), with its timestamped
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

    /// <summary>
    /// First run after the rename: copies settings.json (and ui-state.json) from the legacy folder,
    /// <c>%APPDATA%\ModelingToolkit</c>, when this folder has no settings (see <see cref="LegacySettingsMigration"/>).
    /// The legacy folder is never changed. Returns a line for the log, or null if nothing was copied. Never throws.
    /// </summary>
    public static string? MigrateLegacyFolder()
    {
        var legacyFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            LegacySettingsMigration.LegacyFolderName);
        try
        {
            var copy = LegacySettingsMigration.FilesToCopy(FileNames(legacyFolder), FileNames(File.DirectoryPath));
            if (copy.Count == 0)
            {
                return null;
            }

            Directory.CreateDirectory(File.DirectoryPath);
            foreach (var name in copy)
            {
                System.IO.File.Copy(Path.Combine(legacyFolder, name), Path.Combine(File.DirectoryPath, name), overwrite: false);
            }

            return $"copied {string.Join(", ", copy)} from {legacyFolder}";
        }
        catch (Exception ex)
        {
            return $"copy from {legacyFolder} failed: {ex.Message}";
        }
    }

    private static string[] FileNames(string folder) =>
        Directory.Exists(folder) ? Array.ConvertAll(Directory.GetFiles(folder), Path.GetFileName) : Array.Empty<string>();
}
