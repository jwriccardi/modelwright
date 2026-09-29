using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace ExcelModelingToolkit.AddIn;

/// <summary>
/// Tab-separated diagnostics log, <c>%LOCALAPPDATA%\ModelingToolkit\log.txt</c>, one line per event:
/// UTC timestamp, event, details. When the file passes 1 MB it is renamed to <c>log.1.txt</c> (one backup).
/// Written only while <see cref="Enabled"/> (the <c>diagnosticsLog</c> setting). Never throws.
/// </summary>
internal static class DiagnosticsLog
{
    private const long MaxBytes = 1024 * 1024;

    private static readonly Encoding Utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    private static readonly string LogDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        ProductInfo.FolderName);

    private static readonly string FilePath = Path.Combine(LogDirectory, "log.txt");
    private static readonly string BackupPath = Path.Combine(LogDirectory, "log.1.txt");

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

            Directory.CreateDirectory(LogDirectory);
            RollIfFull();
            File.AppendAllText(FilePath, line.ToString(), Utf8NoBom);
        }
        catch (Exception)
        {
            // Diagnostics must never break a command.
        }
    }

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
