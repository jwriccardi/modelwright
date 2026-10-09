using System;
using System.IO;
using System.Text;
using Modelwright.Core.Trace;

namespace Modelwright.AddIn;

/// <summary>
/// The Trace In window's remembered position, size and wrap setting, and whether the Macabacus notice was shown,
/// <c>%APPDATA%\Modelwright\ui-state.json</c> (<see cref="TraceUiState"/>). Best effort, and kept apart from settings.json and its overwrite protection: a
/// missing, unreadable or damaged file gives the defaults, and a failed save is only logged. No member throws.
/// </summary>
internal static class UiStateStore
{
    private static readonly Encoding Utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    /// <summary>Full path of the file.</summary>
    public static string FilePath { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        ProductInfo.FolderName,
        "ui-state.json");

    /// <summary>The saved state, or the defaults.</summary>
    public static TraceUiState Load()
    {
        try
        {
            return File.Exists(FilePath) ? TraceUiState.FromJson(File.ReadAllText(FilePath, Utf8NoBom)) : new TraceUiState();
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("UiState", "load failed", ex.Message);
            return new TraceUiState();
        }
    }

    /// <summary>
    /// Saves the Trace In window's <paramref name="state"/>, keeping the file's Macabacus notice flag (which
    /// <paramref name="state"/>, loaded when the window opened, may predate).
    /// </summary>
    public static void Save(TraceUiState state) => Write(state.WithMacabacusNoticeShown(Load().MacabacusNoticeShown));

    /// <summary>Remembers that the one-time Macabacus notice was shown (<see cref="MacabacusCheck"/>).</summary>
    public static void MarkMacabacusNoticeShown() => Write(Load().WithMacabacusNoticeShown(true));

    /// <summary>Writes <paramref name="state"/> through a temporary file, so a failed write leaves the old file whole.</summary>
    private static void Write(TraceUiState state)
    {
        var temporary = FilePath + ".tmp";
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
            File.WriteAllText(temporary, state.ToJson(), Utf8NoBom);
            if (File.Exists(FilePath))
            {
                File.Copy(temporary, FilePath, overwrite: true);
                File.Delete(temporary);
            }
            else
            {
                File.Move(temporary, FilePath);
            }
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("UiState", "save failed", ex.Message);
            try
            {
                File.Delete(temporary);
            }
            catch (Exception)
            {
                // Nothing more to do; the next save overwrites it.
            }
        }
    }
}
