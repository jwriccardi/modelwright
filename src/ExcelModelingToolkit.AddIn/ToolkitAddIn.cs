using System;
using ExcelDna.Integration;

namespace ExcelModelingToolkit.AddIn;

/// <summary>
/// Add-in lifetime: loads the settings and registers their keyboard shortcuts on load, and restores Excel's
/// defaults for those keys on unload.
/// </summary>
public sealed class ToolkitAddIn : IExcelAddIn
{
    /// <inheritdoc />
    public void AutoOpen()
    {
        try
        {
            var load = Session.Initialize();
            DiagnosticsLog.Write("AutoOpen", ProductInfo.Version);
            var failures = KeyBindings.Apply(Session.Settings.Keymap);
            var message = Session.Summarize($"{ProductInfo.Name} {ProductInfo.Version} loaded", load, failures);

            // Self-check: every action must have a command. A missing one is skipped (not bound), never fatal.
            var missing = KeyBindings.ActionsWithoutCommand();
            if (missing.Count > 0)
            {
                DiagnosticsLog.Write("SelfCheckFailed", "no command for: " + string.Join(", ", missing));
                message += ". Internal problem: no command for " + string.Join(", ", missing);
            }

            StatusBar.Show(message);
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("AutoOpenFailed", ex.ToString());
            StatusBar.Show($"{ProductInfo.Name} {ProductInfo.Version}: failed to start: {ex.Message}");
        }
    }

    /// <inheritdoc />
    public void AutoClose()
    {
        KeyBindings.Clear();
        DiagnosticsLog.Write("AutoClose");
        StatusBar.Shutdown();
    }
}
