using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelModelingToolkit.Core.Settings;

namespace ExcelModelingToolkit.AddIn;

/// <summary>Excel macros. Each runs in macro context and never throws into Excel.</summary>
public static class Commands
{
    /// <summary>General Number cycle.</summary>
    [ExcelCommand(Name = "EmtNumberCycle")]
    public static void EmtNumberCycle() => RunCycle(ActionIds.NumberCycle);

    /// <summary>Date cycle.</summary>
    [ExcelCommand(Name = "EmtDateCycle")]
    public static void EmtDateCycle() => RunCycle(ActionIds.DateCycle);

    /// <summary>Local Currency cycle.</summary>
    [ExcelCommand(Name = "EmtCurrencyCycle")]
    public static void EmtCurrencyCycle() => RunCycle(ActionIds.CurrencyCycle);

    /// <summary>Percent cycle.</summary>
    [ExcelCommand(Name = "EmtPercentCycle")]
    public static void EmtPercentCycle() => RunCycle(ActionIds.PercentCycle);

    /// <summary>Multiple cycle.</summary>
    [ExcelCommand(Name = "EmtMultipleCycle")]
    public static void EmtMultipleCycle() => RunCycle(ActionIds.MultipleCycle);

    /// <summary>Binary cycle.</summary>
    [ExcelCommand(Name = "EmtBinaryCycle")]
    public static void EmtBinaryCycle() => RunCycle(ActionIds.BinaryCycle);

    /// <summary>Ratio cycle.</summary>
    [ExcelCommand(Name = "EmtRatioCycle")]
    public static void EmtRatioCycle() => RunCycle(ActionIds.RatioCycle);

    /// <summary>Font Color cycle.</summary>
    [ExcelCommand(Name = "EmtFontColorCycle")]
    public static void EmtFontColorCycle() => RunCycle(ActionIds.FontColorCycle);

    /// <summary>Fill Color cycle.</summary>
    [ExcelCommand(Name = "EmtFillColorCycle")]
    public static void EmtFillColorCycle() => RunCycle(ActionIds.FillColorCycle);

    /// <summary>Blue-Black toggle.</summary>
    [ExcelCommand(Name = "EmtBlueBlackToggle")]
    public static void EmtBlueBlackToggle() => RunCycle(ActionIds.BlueBlackToggle);

    /// <summary>Trace In: the precedents of the active cell (docs/PLAN.md section 4.5).</summary>
    [ExcelCommand(Name = "EmtTraceIn")]
    public static void EmtTraceIn() => TraceCommand.TraceIn(KeyBindings.KeyFor(ActionIds.TraceIn));

    /// <summary>Last Audited Cell: back to the cell Trace In was last opened on.</summary>
    [ExcelCommand(Name = "EmtLastAuditedCell")]
    public static void EmtLastAuditedCell() => TraceCommand.LastAuditedCell(KeyBindings.KeyFor(ActionIds.LastAuditedCell));

    /// <summary>Shows the product name, version, commit, build date and add-in path.</summary>
    [ExcelCommand(Name = "EmtAbout")]
    public static void EmtAbout()
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
    [ExcelCommand(Name = "EmtReregisterKeys")]
    public static void EmtReregisterKeys()
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
    [ExcelCommand(Name = "EmtReloadSettings")]
    public static void EmtReloadSettings()
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
    [ExcelCommand(Name = "EmtOpenSettings")]
    public static void EmtOpenSettings()
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
    [ExcelCommand(Name = "EmtSettings")]
    public static void EmtSettings() => ShowSettings(Environment.TickCount);

    /// <summary>
    /// <see cref="EmtSettings"/>, requested at <paramref name="requestedAt"/> (<see cref="Environment.TickCount"/>):
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
