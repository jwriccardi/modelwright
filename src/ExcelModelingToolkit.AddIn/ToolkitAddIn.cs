using System;
using ExcelDna.Integration;
using ExcelModelingToolkit.Core.Keys;

namespace ExcelModelingToolkit.AddIn;

/// <summary>Add-in lifetime: registers keyboard shortcuts on load and restores Excel's defaults on unload.</summary>
public sealed class ToolkitAddIn : IExcelAddIn
{
    /// <summary>Placeholder shortcut (not a Macabacus key) that proves key registration works.</summary>
    internal static readonly KeyChord AboutKey = KeyChord.Parse("Ctrl+Alt+Shift+F12");

    /// <inheritdoc />
    public void AutoOpen()
    {
        var message = $"{ProductInfo.Name} {ProductInfo.Version} loaded";
        var error = RegisterKey(AboutKey, nameof(Commands.EmtAbout));
        if (error != null)
        {
            message += $". Could not register {AboutKey}: {error}";
        }

        StatusBar.Show(message);
    }

    /// <inheritdoc />
    public void AutoClose()
    {
        try
        {
            // xlcOnKey with no macro restores Excel's default behavior for the key.
            XlCall.Excel(XlCall.xlcOnKey, AboutKey.ToOnKeyString());
        }
        catch (XlCallException)
        {
            // Excel is shutting down or the key was never registered; nothing to restore.
        }

        StatusBar.Shutdown();
    }

    /// <summary>Binds <paramref name="chord"/> to <paramref name="macro"/>. Returns null on success, else the reason.</summary>
    private static string? RegisterKey(KeyChord chord, string macro)
    {
        try
        {
            // The C API can "succeed" yet return FALSE or an error value for a rejected key string.
            var result = XlCall.Excel(XlCall.xlcOnKey, chord.ToOnKeyString(), macro);
            if (result is ExcelError || (result is bool ok && !ok))
            {
                return $"xlcOnKey returned {result}";
            }

            return null;
        }
        catch (Exception ex)
        {
            return ex.Message;
        }
    }
}
