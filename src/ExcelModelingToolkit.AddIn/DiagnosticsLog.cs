using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;

namespace ExcelModelingToolkit.AddIn;

/// <summary>
/// Tab-separated diagnostics log, <c>%LOCALAPPDATA%\ModelingToolkit\log.txt</c>, one line per event:
/// UTC timestamp, event, details. When the file passes 1 MB it is renamed to <c>log.1.txt</c> (one backup).
/// Written only while <see cref="Enabled"/> (the <c>diagnosticsLog</c> setting). Never throws.
/// </summary>
/// <remarks>
/// Several Excel instances can share the file: each line is appended under a session-wide named mutex, with the
/// file opened for append and shared, so lines do not interleave and rolling never races an append. If the mutex
/// is busy for longer than <see cref="LockTimeout"/>, the line is dropped rather than stall a command.
/// </remarks>
internal static class DiagnosticsLog
{
    private const long MaxBytes = 1024 * 1024;

    /// <summary>Named mutex that serializes writers across Excel instances in this logon session.</summary>
    private const string MutexName = @"Local\ModelingToolkit.log";

    private static readonly TimeSpan LockTimeout = TimeSpan.FromMilliseconds(200);

    private static readonly Encoding Utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    private static readonly string LogDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        ProductInfo.FolderName);

    private static readonly string FilePath = Path.Combine(LogDirectory, "log.txt");
    private static readonly string BackupPath = Path.Combine(LogDirectory, "log.1.txt");

    /// <summary>The folder that holds the log files (it may not exist yet).</summary>
    public static string FolderPath => LogDirectory;

    /// <summary>True to write; false makes <see cref="Write"/> a no-op.</summary>
    public static bool Enabled { get; set; } = true;

    /// <summary>
    /// Appends a line: timestamp, <paramref name="eventName"/>, then each of <paramref name="fields"/>, tab-separated.
    /// Tabs and line breaks inside values are replaced with spaces. Never throws.
    /// </summary>
    public static void Write(string eventName, params string[] fields)
    {
        if (!Enabled)
        {
            return;
        }

        try
        {
            var line = new StringBuilder(128)
                .Append(DateTime.UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'", CultureInfo.InvariantCulture))
                .Append('\t').Append(Clean(eventName));
            foreach (var field in fields)
            {
                line.Append('\t').Append(Clean(field));
            }

            line.Append("\r\n");
            var bytes = Utf8NoBom.GetBytes(line.ToString());

            Directory.CreateDirectory(LogDirectory);
            using var mutex = new Mutex(initiallyOwned: false, MutexName);
            if (!Acquire(mutex))
            {
                return;
            }

            try
            {
                RollIfFull();
                using var stream = new FileStream(
                    FilePath,
                    FileMode.Append,
                    FileAccess.Write,
                    FileShare.ReadWrite | FileShare.Delete);
                stream.Write(bytes, 0, bytes.Length);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }
        catch (Exception)
        {
            // Diagnostics must never break a command.
        }
    }

    /// <summary>True if this thread now owns <paramref name="mutex"/>; false if another writer held it too long.</summary>
    private static bool Acquire(Mutex mutex)
    {
        try
        {
            return mutex.WaitOne(LockTimeout);
        }
        catch (AbandonedMutexException)
        {
            // A writer died holding the lock; this thread owns it now, and the file is still usable.
            return true;
        }
    }

    /// <summary>Renames a full log to the backup. Call while holding the mutex, so the length checked is current.</summary>
    private static void RollIfFull()
    {
        var file = new FileInfo(FilePath);
        if (!file.Exists || file.Length <= MaxBytes)
        {
            return;
        }

        if (File.Exists(BackupPath))
        {
            File.Delete(BackupPath);
        }

        File.Move(FilePath, BackupPath);
    }

    private static string Clean(string? value) =>
        (value ?? string.Empty).Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ');
}
