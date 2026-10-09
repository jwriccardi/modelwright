namespace Modelwright.Core.Settings;

/// <summary>What <see cref="SettingsFile.Load"/> found: the settings, the file's hash and any backup it restored.</summary>
public sealed class SettingsFileLoadResult
{
    internal SettingsFileLoadResult(SettingsLoadResult result, string? hash, string? restoredFrom)
    {
        Result = result;
        Hash = hash;
        RestoredFrom = restoredFrom;
    }

    /// <summary>The settings, outcome and problems, as for any settings source.</summary>
    public SettingsLoadResult Result { get; }

    /// <summary>
    /// The SHA-256 (<see cref="SettingsFile.ComputeHash"/>) of the bytes read, or of the defaults written when the
    /// file was created; null if nothing was read or written (the file could not be read, was too large, or could
    /// not be created).
    /// </summary>
    public string? Hash { get; }

    /// <summary>The backup copied back as the settings file because the file was missing, or null.</summary>
    public string? RestoredFrom { get; }
}
