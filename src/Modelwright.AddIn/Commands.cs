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
    /// Override button). With the shortcuts switched off, binds nothing and says so.
    /// </summary>
    [ExcelCommand(Name = "MwReregisterKeys")]
    public static void MwReregisterKeys()
    {
        try
        {
            if (!Session.Settings.UseKeyboardShortcuts)
            {
                StatusBar.Show($"{ProductInfo.Name}: shortcuts are off; switch them on with {ProductInfo.Name} > Shortcuts.");
                return;
            }

            var failures = KeyBindings.Apply(Session.Settings);
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

            var failures = KeyBindings.Apply(Session.Settings);
            StatusBar.Show(Session.Summarize(
                $"{ProductInfo.Name}: settings reloaded, {KeyBindings.Describe(Session.Settings)}", load, failures));
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

            var wasOn = Session.Settings.UseKeyboardShortcuts;
            Session.ApplySaved(saved.Value.Settings, saved.Value.Save, "dialog");
            var failures = KeyBindings.Apply(Session.Settings);
            StatusBar.Show(wasOn && !Session.Settings.UseKeyboardShortcuts
                ? OffMessage
                : Session.Summarize($"{ProductInfo.Name}: settings saved, {KeyBindings.Describe(Session.Settings)}", null, failures));
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("SettingsDialogFailed", ex.ToString());
            StatusBar.Show($"{ProductInfo.Name}: the settings dialog failed: {ex.Message}");
        }
    }

    /// <summary>The status-bar message for the shortcuts switched off while Excel runs.</summary>
    internal static string OffMessage =>
        $"{ProductInfo.Name}'s shortcuts are off. Restart Excel to give them back to Macabacus (until then they do " +
        "Excel's usual thing). The ribbon still works.";

    /// <summary>
    /// Switches the keyboard shortcuts on or off (<see cref="ToolkitSettings.UseKeyboardShortcuts"/>) from the ribbon's
    /// Shortcuts button or the Macabacus notice (<paramref name="source"/>, for the log): saves settings.json the way
    /// the settings dialog does (atomically, with a backup, asking first if the file was changed outside the add-in
    /// or is not in use), then applies the switch at once (<see cref="KeyBindings.Apply(ToolkitSettings)"/>) and says
    /// so on the status bar. Returns true if the setting is now <paramref name="on"/>. Call in macro context. Never
    /// throws.
    /// </summary>
    internal static bool SetKeyboardShortcuts(bool on, string source)
    {
        // Excel keeps one binding per key and Macabacus binds its keys only when it loads: a key we release goes to
        // Excel's default, not back to Macabacus (measured 2026-10-09), hence the restart in the off message.
        try
        {
            if (Session.Settings.UseKeyboardShortcuts == on)
            {
                return true;
            }

            var settings = Session.Settings.WithUseKeyboardShortcuts(on);
            if (SettingsStore.Exists() &&
                (Session.SourceState == SettingsLoadOutcome.Rejected || SettingsStore.HasChangedSince(Session.SourceHash)))
            {
                var answer = MessageBox.Show(
                    SettingsDialog.ExcelOwner(),
                    $"settings.json was changed outside {ProductInfo.Name} (or has problems and isn't in use). Switching " +
                    "the shortcuts saves the settings in use now over it; the current file will be kept as a backup.\n\n" +
                    "Switch anyway? Choose No to leave everything as it is; you can click Reload settings to use the " +
                    "file's changes first.",
                    ProductInfo.Name,
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2);
                if (answer != DialogResult.Yes)
                {
                    StatusBar.Show($"{ProductInfo.Name}: shortcuts not switched; settings.json was not changed.");
                    return false;
                }
            }

            var save = SettingsStore.Save(settings);
            if (!save.Succeeded)
            {
                StatusBar.Show($"{ProductInfo.Name}: shortcuts not switched: the settings were not saved: {save.Problem}");
                return false;
            }

            Session.ApplySaved(settings, save, source);
            var failures = KeyBindings.Apply(Session.Settings);
            StatusBar.Show(on
                ? Session.Summarize($"{ProductInfo.Name} shortcuts on", null, failures)
                : OffMessage);
            return true;
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("ShortcutsSwitchFailed", source, ex.ToString());
            StatusBar.Show($"{ProductInfo.Name}: switching the shortcuts failed: {ex.Message}");
            return false;
        }
        finally
        {
            // The toggle shows the setting in use, also when the switch was refused or failed.
            ToolkitRibbon.RefreshShortcuts();
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
