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
        var load = Session.LoadSettings();
        DiagnosticsLog.Write("AutoOpen", ProductInfo.Version);
        var failures = KeyBindings.Apply(Session.Settings.Keymap);
        StatusBar.Show(Session.Summarize($"{ProductInfo.Name} {ProductInfo.Version} loaded", load, failures));
    }

    /// <inheritdoc />
    public void AutoClose()
    {
        KeyBindings.Clear();
        DiagnosticsLog.Write("AutoClose");
        StatusBar.Shutdown();
    }
}
