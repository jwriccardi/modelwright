namespace ExcelModelingToolkit.Core.Settings;

/// <summary>What <see cref="SettingsFile.Save"/> did.</summary>
public sealed class SettingsFileSaveResult
{
    internal SettingsFileSaveResult(string? problem, string? hash, string? backupPath, string? restoredFrom)
    {
        Problem = problem;
        Hash = hash;
        BackupPath = backupPath;
        RestoredFrom = restoredFrom;
    }

    /// <summary>Why the settings were not saved, or null if they were.</summary>
    public string? Problem { get; }

    /// <summary>True if the settings were saved.</summary>
    public bool Succeeded => Problem is null;

    /// <summary>The SHA-256 (<see cref="SettingsFile.ComputeHash"/>) of the bytes written, or null if the save failed.</summary>
    public string? Hash { get; }

    /// <summary>The backup made of the replaced file, or null if none was needed (see <see cref="SettingsFile.Save"/>).</summary>
    public string? BackupPath { get; }

    /// <summary>
    /// After a failed save that left no settings file, the backup copied back in its place; otherwise null.
    /// </summary>
    public string? RestoredFrom { get; }
}
