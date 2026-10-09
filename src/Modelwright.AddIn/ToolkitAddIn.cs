using System;
using ExcelDna.Integration;

namespace Modelwright.AddIn;

/// <summary>
/// Add-in lifetime: loads the settings, registers their keyboard shortcuts (unless they are switched off), installs the
/// Ctrl+Z / Ctrl+Y hook, connects the Excel events that invalidate undo history and schedules the Macabacus check
/// (<see cref="MacabacusCheck"/>) on load; closes any Trace In window (removing its key hook),
/// disconnects the events, removes the hook, empties our undo stacks and restores Excel's defaults for our keys on
/// unload.
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
            var failures = KeyBindings.Apply(Session.Settings);
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

            // Trace In's one-time start-up costs, paid shortly after Excel has loaded us rather than at the first trace.
            TraceSession.ScheduleWarmUp();

            // Macabacus binds the same keys: look for it once Excel has loaded the other add-ins.
            MacabacusCheck.Schedule();
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
        TraceSession.CancelWarmUp();
        MacabacusCheck.Cancel();
        TraceSession.Current?.Abort("add-in closing");
        TraceWindow.DestroySpare();
        TraceKeyHook.Uninstall();
        ExcelEvents.Disconnect();
        UndoKeyHook.Uninstall();
        Session.Undo.Clear();
        KeyBindings.Clear();
        DiagnosticsLog.Write("AutoClose");
        StatusBar.Shutdown();
    }
}
