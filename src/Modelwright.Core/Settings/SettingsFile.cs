using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace Modelwright.Core.Settings;

/// <summary>
/// The settings file on disk (<c>settings.json</c>, UTF-8, human-editable) with its backups: loading, atomic
/// saving, and a SHA-256 of the contents so a caller can tell whether the file changed since it loaded or saved it.
/// No member throws for a file system problem; each reports it instead.
/// </summary>
/// <remarks>
/// <para>
/// <b>Backups.</b> A save that replaces the file first keeps the replaced contents as
/// <c>settings.json.YYYYMMDD-HHMMSS.bak</c> (UTC time; <c>-2</c>, <c>-3</c>... is added to a name already taken),
/// unless the newest backup already holds exactly those contents. At most <see cref="MaxBackups"/> are kept; the
/// oldest are deleted.
/// </para>
/// <para>
/// <b>A missing file</b> is restored from the newest backup (by <see cref="Load"/>, <see cref="EnsureExists"/>, and
/// a save that failed part-way). Only when there is no backup are the defaults written.
/// </para>
/// </remarks>
public sealed class SettingsFile
{
    /// <summary>The most backups kept.</summary>
    public const int MaxBackups = 3;

    private const string StampFormat = "yyyyMMdd-HHmmss";

    private static readonly Encoding Utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    private readonly Regex _backupName;
    private readonly Regex _temporaryName;

    /// <summary>Creates the file object for <paramref name="path"/> (the file need not exist).</summary>
    /// <exception cref="ArgumentNullException"><paramref name="path"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="path"/> does not name a file.</exception>
    public SettingsFile(string path)
    {
        if (path is null)
        {
            throw new ArgumentNullException(nameof(path));
        }

        FilePath = Path.GetFullPath(path);
        FileName = Path.GetFileName(FilePath);
        var directory = Path.GetDirectoryName(FilePath);
        if (FileName.Length == 0 || directory is null)
        {
            throw new ArgumentException($"'{path}' does not name a file.", nameof(path));
        }

        DirectoryPath = directory;
        var name = Regex.Escape(FileName);
        _backupName = new Regex("^" + name + @"\.(\d{8}-\d{6})(?:-(\d{1,9}))?\.bak$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        _temporaryName = new Regex("^" + name + @"\.[0-9a-f]{32}\.tmp$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    }

    /// <summary>The full path of the settings file.</summary>
    public string FilePath { get; }

    /// <summary>The settings file's name, e.g. <c>settings.json</c>, as problems name it.</summary>
    public string FileName { get; }

    /// <summary>The folder that holds the file, its backups and its temporary files.</summary>
    public string DirectoryPath { get; }

    /// <summary>The current UTC time, for backup names and the age of temporary files. Replaceable for tests.</summary>
    internal Func<DateTime> UtcNow { get; set; } = () => DateTime.UtcNow;

    /// <summary><see cref="File.Replace(string, string, string)"/>. Replaceable for tests.</summary>
    internal Action<string, string, string?> ReplaceFile { get; set; } =
        (source, destination, backup) => File.Replace(source, destination, backup, ignoreMetadataErrors: true);

    /// <summary>How long a save waits before trying a failed replace once more.</summary>
    internal TimeSpan RetryDelay { get; set; } = TimeSpan.FromMilliseconds(200);

    /// <summary>The SHA-256 of <paramref name="bytes"/> as 64 lower-case hex digits.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="bytes"/> is null.</exception>
    public static string ComputeHash(byte[] bytes)
    {
        if (bytes is null)
        {
            throw new ArgumentNullException(nameof(bytes));
        }

        using var sha = SHA256.Create();
        return ToHex(sha.ComputeHash(bytes));
    }

    /// <summary>
    /// Loads the file. If it is missing, restores the newest backup and loads that
    /// (<see cref="SettingsFileLoadResult.RestoredFrom"/>); with no backup, writes the defaults (so people can find
    /// and edit the file) and returns them (<see cref="SettingsLoadOutcome.CreatedDefaults"/>). A file that cannot
    /// be read, is larger than <see cref="ToolkitSettings.MaxFileBytes"/>, is not valid UTF-8 or is invalid gives
    /// <see cref="SettingsLoadOutcome.Rejected"/> with the problems; it is left untouched so the user can fix it.
    /// </summary>
    public SettingsFileLoadResult Load()
    {
        string? restoredFrom = null;
        try
        {
            if (!File.Exists(FilePath))
            {
                restoredFrom = RestoreNewestBackup();
                if (restoredFrom is null && !File.Exists(FilePath))
                {
                    // If another Excel instance creates the file first, read theirs.
                    if (TryCreateDefaults(out var problem, out var hash) || !File.Exists(FilePath))
                    {
                        return new SettingsFileLoadResult(
                            new SettingsLoadResult(
                                ToolkitSettings.Defaults(),
                                problem is null ? Array.Empty<string>() : new[] { problem },
                                SettingsLoadOutcome.CreatedDefaults),
                            hash,
                            null);
                    }
                }
            }

            var tooLarge = ToolkitSettings.CheckFileSize(new FileInfo(FilePath).Length, FileName);
            if (tooLarge is not null)
            {
                return new SettingsFileLoadResult(SettingsLoadResult.Rejected(new[] { tooLarge }), null, restoredFrom);
            }

            var bytes = File.ReadAllBytes(FilePath);
            return new SettingsFileLoadResult(ToolkitSettings.FromFileBytes(bytes, FileName), ComputeHash(bytes), restoredFrom);
        }
        catch (Exception ex)
        {
            return new SettingsFileLoadResult(
                SettingsLoadResult.Rejected(new[] { $"could not read {FilePath}: {ex.Message}" }),
                null,
                restoredFrom);
        }
    }

    /// <summary>
    /// Makes sure the file exists: if it is missing, restores the newest backup, or with none writes the defaults.
    /// Returns null on success, else the reason.
    /// </summary>
    public string? EnsureExists()
    {
        try
        {
            if (File.Exists(FilePath) || RestoreNewestBackup() is not null)
            {
                return null;
            }

            TryCreateDefaults(out var problem, out _);
            return problem;
        }
        catch (Exception ex)
        {
            return $"could not create {FilePath}: {ex.Message}";
        }
    }

    /// <summary>
    /// True if the file exists and its contents are not the ones whose hash is <paramref name="knownHash"/> (the
    /// hash from the last load or save; null matches nothing), or it cannot be read to tell. False if it is missing.
    /// </summary>
    public bool HasChangedSince(string? knownHash)
    {
        try
        {
            return File.Exists(FilePath) && !string.Equals(HashOfFile(FilePath), knownHash, StringComparison.Ordinal);
        }
        catch (Exception)
        {
            return true;
        }
    }

    /// <summary>
    /// Saves <paramref name="settings"/> atomically: the JSON is written to a temporary file in the same folder,
    /// which then replaces the file (<see cref="File.Replace(string, string, string)"/>, tried once more after a
    /// short pause if the file is busy), keeping the replaced contents as a backup unless the newest backup already
    /// holds them (see the remarks on <see cref="SettingsFile"/>). Where the file system cannot replace (some
    /// network shares), the file is copied to the backup, deleted, and the temporary file renamed in its place. A
    /// new file is simply renamed into place. On failure the file is left as it was; if the failure left no file,
    /// the newest backup is copied back (<see cref="SettingsFileSaveResult.RestoredFrom"/>). The temporary file is
    /// always removed.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="settings"/> is null.</exception>
    public SettingsFileSaveResult Save(ToolkitSettings settings)
    {
        if (settings is null)
        {
            throw new ArgumentNullException(nameof(settings));
        }

        var bytes = Utf8NoBom.GetBytes(settings.ToJson());
        var tempPath = NewTemporaryPath();
        try
        {
            Directory.CreateDirectory(DirectoryPath);
            File.WriteAllBytes(tempPath, bytes);
            if (!File.Exists(FilePath))
            {
                File.Move(tempPath, FilePath);
                return new SettingsFileSaveResult(null, ComputeHash(bytes), null, null);
            }

            var backupPath = NeedsBackup() ? NewBackupPath() : null;
            Replace(tempPath, backupPath);
            if (backupPath is not null)
            {
                PruneBackups();
            }

            return new SettingsFileSaveResult(null, ComputeHash(bytes), backupPath, null);
        }
        catch (Exception ex)
        {
            var problem = $"could not save {FilePath}: {ex.Message}";
            var restoredFrom = File.Exists(FilePath) ? null : RestoreNewestBackup();
            if (restoredFrom is not null)
            {
                problem += $" The previous settings were restored from {Path.GetFileName(restoredFrom)}.";
            }

            return new SettingsFileSaveResult(problem, null, null, restoredFrom);
        }
        finally
        {
            TryDelete(tempPath);
        }
    }

    /// <summary>The backups (see the remarks on <see cref="SettingsFile"/>), newest first. Empty if there are none or the folder cannot be read.</summary>
    public IReadOnlyList<string> GetBackups()
    {
        try
        {
            if (!Directory.Exists(DirectoryPath))
            {
                return Array.Empty<string>();
            }

            var backups = new List<(string Stamp, int Number, string Path)>();
            foreach (var path in Directory.GetFiles(DirectoryPath, FileName + ".*.bak"))
            {
                var match = _backupName.Match(Path.GetFileName(path));
                if (match.Success)
                {
                    var number = match.Groups[2].Success ? int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture) : 1;
                    backups.Add((match.Groups[1].Value, number, path));
                }
            }

            return backups
                .OrderByDescending(b => b.Stamp, StringComparer.Ordinal)
                .ThenByDescending(b => b.Number)
                .Select(b => b.Path)
                .ToArray();
        }
        catch (Exception)
        {
            return Array.Empty<string>();
        }
    }

    /// <summary>
    /// Deletes this file's temporary files (<c>settings.json.&lt;guid&gt;.tmp</c>, left behind if Excel stopped
    /// during a save) last written more than <paramref name="olderThan"/> ago; newer ones may belong to a save in
    /// progress in another Excel instance. Returns how many were deleted.
    /// </summary>
    public int DeleteStaleTemporaryFiles(TimeSpan olderThan)
    {
        try
        {
            if (!Directory.Exists(DirectoryPath))
            {
                return 0;
            }

            var cutoff = UtcNow() - olderThan;
            var deleted = 0;
            foreach (var path in Directory.GetFiles(DirectoryPath, FileName + ".*.tmp"))
            {
                if (_temporaryName.IsMatch(Path.GetFileName(path)) && File.GetLastWriteTimeUtc(path) < cutoff && TryDelete(path))
                {
                    deleted++;
                }
            }

            return deleted;
        }
        catch (Exception)
        {
            return 0;
        }
    }

    private static string HashOfFile(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var sha = SHA256.Create();
        return ToHex(sha.ComputeHash(stream));
    }

    private static string ToHex(byte[] hash)
    {
        var text = new StringBuilder(hash.Length * 2);
        foreach (var b in hash)
        {
            text.Append(b.ToString("x2", CultureInfo.InvariantCulture));
        }

        return text.ToString();
    }

    private static bool TryDelete(string path)
    {
        try
        {
            File.Delete(path); // No error if it does not exist.
            return true;
        }
        catch (Exception)
        {
            return false; // A leftover temporary file or extra backup is harmless.
        }
    }

    /// <summary>Replaces the file with <paramref name="tempPath"/>, keeping it as <paramref name="backupPath"/> if not null.</summary>
    private void Replace(string tempPath, string? backupPath)
    {
        try
        {
            try
            {
                ReplaceFile(tempPath, FilePath, backupPath);
            }
            catch (IOException)
            {
                // Often a moment's lock (an editor, a virus scanner or a sync client reading the file): once more.
                Thread.Sleep(RetryDelay);
                ReplaceFile(tempPath, FilePath, backupPath);
            }
        }
        catch (Exception ex) when (ex is IOException || ex is PlatformNotSupportedException)
        {
            // Replace is unsupported here, or failed part-way (it may have moved the file to the backup already).
            if (File.Exists(FilePath))
            {
                if (backupPath is not null)
                {
                    File.Copy(FilePath, backupPath, overwrite: true);
                }

                File.Delete(FilePath);
            }

            File.Move(tempPath, FilePath);
        }
    }

    /// <summary>False if the newest backup holds exactly the file's current contents.</summary>
    private bool NeedsBackup()
    {
        var newest = GetBackups().FirstOrDefault();
        if (newest is null)
        {
            return true;
        }

        try
        {
            return !string.Equals(HashOfFile(FilePath), HashOfFile(newest), StringComparison.Ordinal);
        }
        catch (Exception)
        {
            return true;
        }
    }

    private string NewBackupPath()
    {
        var stamp = UtcNow().ToString(StampFormat, CultureInfo.InvariantCulture);
        var path = Path.Combine(DirectoryPath, $"{FileName}.{stamp}.bak");
        for (var number = 2; File.Exists(path); number++)
        {
            path = Path.Combine(DirectoryPath, $"{FileName}.{stamp}-{number.ToString(CultureInfo.InvariantCulture)}.bak");
        }

        return path;
    }

    private void PruneBackups()
    {
        foreach (var old in GetBackups().Skip(MaxBackups))
        {
            TryDelete(old);
        }
    }

    /// <summary>
    /// Copies the newest backup to the (missing) file, never overwriting. Returns the backup's path, or null if
    /// there is none or the copy failed (e.g. another Excel instance created the file first).
    /// </summary>
    private string? RestoreNewestBackup()
    {
        var newest = GetBackups().FirstOrDefault();
        if (newest is null)
        {
            return null;
        }

        try
        {
            File.Copy(newest, FilePath, overwrite: false);
            return newest;
        }
        catch (Exception)
        {
            return null;
        }
    }

    /// <summary>
    /// Creates the file with the defaults, atomically and without ever overwriting: the JSON goes to a temporary
    /// file that is then renamed, which fails if the file exists (another Excel instance got there first). Returns
    /// true if this call created the file, with the hash of what it wrote. <paramref name="problem"/> is null on
    /// success or when another instance created the file, else the reason. The temporary file is always removed.
    /// </summary>
    private bool TryCreateDefaults(out string? problem, out string? hash)
    {
        problem = null;
        hash = null;
        var tempPath = NewTemporaryPath();
        try
        {
            var bytes = Utf8NoBom.GetBytes(ToolkitSettings.Defaults().ToJson());
            Directory.CreateDirectory(DirectoryPath);
            File.WriteAllBytes(tempPath, bytes);
            File.Move(tempPath, FilePath);
            hash = ComputeHash(bytes);
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

    private string NewTemporaryPath() => FilePath + "." + Guid.NewGuid().ToString("N") + ".tmp";
}
