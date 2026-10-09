using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows.Forms;
using ExcelDna.Integration;
using Modelwright.Core.Settings;

namespace Modelwright.AddIn;

/// <summary>Excel macros. Each runs in macro context and never throws into Excel.</summary>
public static class Commands
{
    /// <summary>General Number cycle.</summary>
    [ExcelCommand(Name = "MwNumberCycle")]
    public static void MwNumberCycle() => RunCycle(ActionIds.NumberCycle);

    /// <summary>Date cycle.</summary>
    [ExcelCommand(Name = "MwDateCycle")]
    public static void MwDateCycle() => RunCycle(ActionIds.DateCycle);

    /// <summary>Local Currency cycle.</summary>
    [ExcelCommand(Name = "MwCurrencyCycle")]
    public static void MwCurrencyCycle() => RunCycle(ActionIds.CurrencyCycle);

    /// <summary>Percent cycle.</summary>
    [ExcelCommand(Name = "MwPercentCycle")]
    public static void MwPercentCycle() => RunCycle(ActionIds.PercentCycle);

    /// <summary>Multiple cycle.</summary>
    [ExcelCommand(Name = "MwMultipleCycle")]
    public static void MwMultipleCycle() => RunCycle(ActionIds.MultipleCycle);

    /// <summary>Binary cycle.</summary>
    [ExcelCommand(Name = "MwBinaryCycle")]
    public static void MwBinaryCycle() => RunCycle(ActionIds.BinaryCycle);

    /// <summary>Ratio cycle.</summary>
    [ExcelCommand(Name = "MwRatioCycle")]
    public static void MwRatioCycle() => RunCycle(ActionIds.RatioCycle);

    /// <summary>Font Color cycle.</summary>
    [ExcelCommand(Name = "MwFontColorCycle")]
    public static void MwFontColorCycle() => RunCycle(ActionIds.FontColorCycle);

    /// <summary>Fill Color cycle.</summary>
    [ExcelCommand(Name = "MwFillColorCycle")]
    public static void MwFillColorCycle() => RunCycle(ActionIds.FillColorCycle);

    /// <summary>Blue-Black toggle.</summary>
    [ExcelCommand(Name = "MwBlueBlackToggle")]
    public static void MwBlueBlackToggle() => RunCycle(ActionIds.BlueBlackToggle);

    /// <summary>Trace In: the precedents of the active cell (docs/PLAN.md section 4.5).</summary>
    [ExcelCommand(Name = "MwTraceIn")]
    public static void MwTraceIn() => TraceCommand.TraceIn(KeyBindings.KeyFor(ActionIds.TraceIn));

    /// <summary>Last Audited Cell: back to the cell Trace In was last opened on.</summary>
    [ExcelCommand(Name = "MwLastAuditedCell")]
    public static void MwLastAuditedCell() => TraceCommand.LastAuditedCell(KeyBindings.KeyFor(ActionIds.LastAuditedCell));

    /// <summary>Shows the product name, version, commit, build date and add-in path.</summary>
    [ExcelCommand(Name = "MwAbout")]
    public static void MwAbout()
    {
        try
        {
            var xllPath = ExcelDnaUtil.XllPath;
            // The add-in path (long, and user-specific) goes in the dialog only, not the status bar.
            StatusBar.Show(
                $"{ProductInfo.Name} {ProductInfo.Version} (commit {ProductInfo.Commit}, built {ProductInfo.BuildDate})");
            MessageBox.Show(
                $"{ProductInfo.Name}\n\n" +
                $"Version: {ProductInfo.Version}\n" +
                $"Commit: {ProductInfo.Commit}\n" +
                $"Built: {ProductInfo.BuildDate}\n" +
                $"Add-in: {xllPath}\n" +
                $"Settings: {SettingsStore.FilePath}",
                "About " + ProductInfo.Name,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            StatusBar.Show($"{ProductInfo.Name}: About failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Re-binds every shortcut in the keymap, taking back keys another add-in has bound since (like Macabacus's
    /// Override button).
    /// </summary>
    [ExcelCommand(Name = "MwReregisterKeys")]
    public static void MwReregisterKeys()
    {
        try
        {
            var failures = KeyBindings.Apply(Session.Settings.Keymap);
            StatusBar.Show(Session.Summarize($"{ProductInfo.Name}: {KeyBindings.Count} shortcuts registered", null, failures));
        }
        catch (Exception ex)
        {
            StatusBar.Show($"{ProductInfo.Name}: re-registering shortcuts failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Reloads settings.json and re-binds the shortcuts. If the file is rejected, keeps the previous settings and
    /// shortcuts and says why.
    /// </summary>
    [ExcelCommand(Name = "MwReloadSettings")]
    public static void MwReloadSettings()
    {
        try
        {
            var load = Session.Reload(out var keptPrevious);
            if (keptPrevious)
            {
                var count = load.Result.Problems.Count;
                var first = count > 0 ? load.Result.Problems[0] : "see the diagnostics log";
                StatusBar.Show(
                    $"{ProductInfo.Name}: settings.json has {count} problem{(count == 1 ? string.Empty : "s")}; " +
                    $"still using the previous settings: {first}");
                return;
            }

            var failures = KeyBindings.Apply(Session.Settings.Keymap);
            StatusBar.Show(Session.Summarize(
                $"{ProductInfo.Name}: settings reloaded, {KeyBindings.Count} shortcuts registered", load, failures));
        }
        catch (Exception ex)
        {
            StatusBar.Show($"{ProductInfo.Name}: reloading settings failed: {ex.Message}");
        }
    }

    /// <summary>Opens settings.json in the default editor for .json files (Notepad if there is none), creating it first if needed.</summary>
    [ExcelCommand(Name = "MwOpenSettings")]
    public static void MwOpenSettings()
    {
        var problem = OpenSettingsFile();
        StatusBar.Show(problem is null
            ? $"{ProductInfo.Name}: opened settings.json. Save your changes, then click Reload settings."
            : $"{ProductInfo.Name}: {problem}");
    }

    /// <summary>
    /// Opens the settings dialog on the current settings. OK saves settings.json and applies it at once, exactly as
    /// Reload settings would (<see cref="Session.ApplySaved"/>), then re-binds the shortcuts; Cancel changes nothing.
    /// Runs as a macro (the ribbon queues it, see <see cref="ShowSettings"/>), so the dialog's number format preview
    /// can call Excel.
    /// </summary>
    [ExcelCommand(Name = "MwSettings")]
    public static void MwSettings() => ShowSettings(Environment.TickCount);

    /// <summary>
    /// <see cref="MwSettings"/>, requested at <paramref name="requestedAt"/> (<see cref="Environment.TickCount"/>):
    /// a request made before the last settings dialog closed (a second click queued while it was open) is ignored,
    /// and one made while a dialog is open brings that dialog to the front.
    /// </summary>
    internal static void ShowSettings(int requestedAt)
    {
        try
        {
            var saved = SettingsDialog.Edit(Session.Settings, requestedAt);
            if (saved is null)
            {
                return;
            }

            Session.ApplySaved(saved.Value.Settings, saved.Value.Save);
            var failures = KeyBindings.Apply(Session.Settings.Keymap);
            StatusBar.Show(Session.Summarize(
                $"{ProductInfo.Name}: settings saved, {KeyBindings.Count} shortcuts registered", null, failures));
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("SettingsDialogFailed", ex.ToString());
            StatusBar.Show($"{ProductInfo.Name}: the settings dialog failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Opens settings.json in the default editor for .json files (Notepad if there is none), creating it first if
    /// needed. Returns null on success, else the reason. Never throws.
    /// </summary>
    internal static string? OpenSettingsFile()
    {
        try
        {
            var problem = SettingsStore.EnsureExists();
            if (problem is not null)
            {
                return problem;
            }

            try
            {
                Process.Start(new ProcessStartInfo(SettingsStore.FilePath) { UseShellExecute = true })?.Dispose();
            }
            catch (Win32Exception)
            {
                // No program is associated with .json files.
                Process.Start(new ProcessStartInfo("notepad.exe", "\"" + SettingsStore.FilePath + "\"") { UseShellExecute = true })?.Dispose();
            }

            return null;
        }
        catch (Exception ex)
        {
            return $"could not open {SettingsStore.FilePath}: {ex.Message}";
        }
    }

    /// <summary>Runs a cycle from the ribbon (queued as a macro by <see cref="ToolkitRibbon"/>).</summary>
    internal static void RunCycleFromRibbon(string actionId) => CycleCommand.Run(actionId, "ribbon");

    private static void RunCycle(string actionId) => CycleCommand.Run(actionId, KeyBindings.KeyFor(actionId));
}
