using System;
using ExcelDna.Integration;

namespace ExcelModelingToolkit.AddIn;

/// <summary>
/// Add-in lifetime: loads the settings, registers their keyboard shortcuts, installs the Ctrl+Z / Ctrl+Y hook and
/// connects the Excel events that invalidate undo history on load; disconnects them, removes the hook, empties our
/// undo stacks and restores Excel's defaults for our keys on unload.
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
            if (!UndoKeyHook.Install())
            {
                message += ". Ctrl+Z / Ctrl+Y cannot undo our formatting; use the Undo formatting button";
            }

            // Without the events, undo history is still checked cell by cell before each restore.
            ExcelEvents.Connect();

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
        ExcelEvents.Disconnect();
        UndoKeyHook.Uninstall();
        Session.Undo.Clear();
        KeyBindings.Clear();
        DiagnosticsLog.Write("AutoClose");
        StatusBar.Shutdown();
    }
}
