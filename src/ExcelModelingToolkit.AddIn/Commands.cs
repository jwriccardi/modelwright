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

    /// <summary>Font Color cycle.</summary>
    [ExcelCommand(Name = "EmtFontColorCycle")]
    public static void EmtFontColorCycle() => RunCycle(ActionIds.FontColorCycle);

    /// <summary>Fill Color cycle.</summary>
    [ExcelCommand(Name = "EmtFillColorCycle")]
    public static void EmtFillColorCycle() => RunCycle(ActionIds.FillColorCycle);

    /// <summary>Blue-Black toggle.</summary>
    [ExcelCommand(Name = "EmtBlueBlackToggle")]
    public static void EmtBlueBlackToggle() => RunCycle(ActionIds.BlueBlackToggle);

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

    /// <summary>Reloads settings.json, resets the cycles and re-binds the shortcuts.</summary>
    [ExcelCommand(Name = "EmtReloadSettings")]
    public static void EmtReloadSettings()
    {
        try
        {
            var load = Session.LoadSettings();
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
        try
        {
            var problem = SettingsStore.EnsureExists();
            if (problem is not null)
            {
                StatusBar.Show($"{ProductInfo.Name}: {problem}");
                return;
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

            StatusBar.Show($"{ProductInfo.Name}: opened settings.json. Save your changes, then click Reload settings.");
        }
        catch (Exception ex)
        {
            StatusBar.Show($"{ProductInfo.Name}: could not open {SettingsStore.FilePath}: {ex.Message}");
        }
    }

    /// <summary>Runs a cycle from the ribbon (queued as a macro by <see cref="ToolkitRibbon"/>).</summary>
    internal static void RunCycleFromRibbon(string actionId) => CycleCommand.Run(actionId, "ribbon");

    private static void RunCycle(string actionId) => CycleCommand.Run(actionId, KeyBindings.KeyFor(actionId));
}
