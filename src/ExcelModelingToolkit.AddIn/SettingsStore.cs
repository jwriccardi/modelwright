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
    /// and returns them. If it cannot be read or is invalid, returns the defaults with the problems; an invalid
    /// file is left untouched so the user can fix it. Never throws.
    /// </summary>
    public static SettingsLoadResult Load()
    {
        try
        {
            if (!File.Exists(FilePath))
            {
                var problem = TryWriteDefaults();
                return new SettingsLoadResult(
                    ToolkitSettings.Defaults(),
                    problem is null ? Array.Empty<string>() : new[] { problem },
                    usedDefaults: problem is not null);
            }

            return ToolkitSettings.FromJson(File.ReadAllText(FilePath, Encoding.UTF8));
        }
        catch (Exception ex)
        {
            return new SettingsLoadResult(
                ToolkitSettings.Defaults(),
                new[] { $"could not read {FilePath}: {ex.Message}" },
                usedDefaults: true);
        }
    }

    /// <summary>Writes the defaults if the file does not exist. Returns null on success, else the reason. Never throws.</summary>
    public static string? EnsureExists()
    {
        try
        {
            return File.Exists(FilePath) ? null : TryWriteDefaults();
        }
        catch (Exception ex)
        {
            return ex.Message;
        }
    }

    private static string? TryWriteDefaults()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
            File.WriteAllText(FilePath, ToolkitSettings.Defaults().ToJson(), Utf8NoBom);
            return null;
        }
        catch (Exception ex)
        {
            return $"could not create {FilePath}: {ex.Message}";
        }
    }
}
