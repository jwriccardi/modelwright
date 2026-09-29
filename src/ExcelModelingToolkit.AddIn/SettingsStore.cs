using System;
using System.IO;
using System.Text;
using ExcelModelingToolkit.Core.Settings;

namespace ExcelModelingToolkit.AddIn;

/// <summary>The settings file, <c>%APPDATA%\ModelingToolkit\settings.json</c> (UTF-8, human-editable).</summary>
internal static class SettingsStore
{
    private static readonly Encoding Utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    /// <summary>Full path of the settings file.</summary>
    public static string FilePath { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        ProductInfo.FolderName,
        "settings.json");

    /// <summary>
    /// Loads the settings file. If it does not exist, writes the defaults there (so people can find and edit it)
    /// and returns them (<see cref="SettingsLoadOutcome.CreatedDefaults"/>). If it cannot be read, is larger than
    /// <see cref="ToolkitSettings.MaxFileBytes"/>, is not valid UTF-8 or is invalid, returns
    /// <see cref="SettingsLoadOutcome.Rejected"/> with the problems; the file is left untouched so the user can
    /// fix it. Never throws.
    /// </summary>
    public static SettingsLoadResult Load()
    {
        try
        {
            if (!File.Exists(FilePath))
            {
                // If another Excel instance creates the file first, read theirs.
                if (TryCreateDefaults(out var problem) || !File.Exists(FilePath))
                {
                    return new SettingsLoadResult(
                        ToolkitSettings.Defaults(),
                        problem is null ? Array.Empty<string>() : new[] { problem },
                        SettingsLoadOutcome.CreatedDefaults);
                }
            }

            var tooLarge = ToolkitSettings.CheckFileSize(new FileInfo(FilePath).Length);
            if (tooLarge is not null)
            {
                return SettingsLoadResult.Rejected(new[] { tooLarge });
            }

            return ToolkitSettings.FromFileBytes(File.ReadAllBytes(FilePath));
        }
        catch (Exception ex)
        {
            return SettingsLoadResult.Rejected(new[] { $"could not read {FilePath}: {ex.Message}" });
        }
    }

    /// <summary>Writes the defaults if the file does not exist. Returns null on success, else the reason. Never throws.</summary>
    public static string? EnsureExists()
    {
        try
        {
            if (File.Exists(FilePath))
            {
                return null;
            }

            TryCreateDefaults(out var problem);
            return problem;
        }
        catch (Exception ex)
        {
            return ex.Message;
        }
    }

    /// <summary>
    /// Saves <paramref name="settings"/> as the settings file, atomically: the JSON is written to a temporary file in
    /// the same folder, which then replaces the file with <see cref="File.Replace(string, string, string, bool)"/>,
    /// keeping the previous file as <c>settings.json.bak</c>. Where the file system cannot replace (some network
    /// shares), the previous file is copied to the backup, deleted, and the temporary file renamed in its place. A new
    /// file is simply renamed into place. Returns null on success, else the reason; on failure the previous file is
    /// left as it was, or can be recovered from the backup. The temporary file is always removed. Never throws.
    /// </summary>
    public static string? Save(ToolkitSettings settings)
    {
        var tempPath = FilePath + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
            File.WriteAllText(tempPath, settings.ToJson(), Utf8NoBom);
            if (!File.Exists(FilePath))
            {
                File.Move(tempPath, FilePath);
                return null;
            }

            var backupPath = FilePath + ".bak";
            try
            {
                File.Replace(tempPath, FilePath, backupPath, ignoreMetadataErrors: true);
            }
            catch (Exception ex) when (ex is IOException || ex is PlatformNotSupportedException)
            {
                // Replace is unsupported here, or failed part-way (it may have moved the file to the backup already).
                if (File.Exists(FilePath))
                {
                    File.Copy(FilePath, backupPath, overwrite: true);
                    File.Delete(FilePath);
                }

                File.Move(tempPath, FilePath);
            }

            return null;
        }
        catch (Exception ex)
        {
            return $"could not save {FilePath}: {ex.Message}";
        }
        finally
        {
            TryDelete(tempPath);
        }
    }

    /// <summary>
    /// Creates the file with the defaults, atomically and without ever overwriting: the JSON goes to a temporary
    /// file that is then renamed, which fails if the file exists (another Excel instance got there first). Returns
    /// true if this call created the file. <paramref name="problem"/> is null on success or when another instance
    /// created the file, else the reason. The temporary file is always removed. Never throws.
    /// </summary>
    private static bool TryCreateDefaults(out string? problem)
    {
        problem = null;
        var tempPath = FilePath + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
            File.WriteAllText(tempPath, ToolkitSettings.Defaults().ToJson(), Utf8NoBom);
            File.Move(tempPath, FilePath);
            return true;
        }
        catch (Exception ex)
        {
            if (!File.Exists(FilePath))
            {
                problem = $"could not create {FilePath}: {ex.Message}";
            }

            return false;
        }
        finally
        {
            TryDelete(tempPath);
        }
    }

    private static void TryDelete(string path)
    {
        try
        {
            File.Delete(path); // No error if it does not exist.
        }
        catch (Exception)
        {
            // A leftover temporary file is harmless.
        }
    }
}
