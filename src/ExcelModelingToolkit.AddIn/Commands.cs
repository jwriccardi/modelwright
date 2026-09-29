using System;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace ExcelModelingToolkit.AddIn;

/// <summary>Excel macros. Each runs in macro context and never throws into Excel.</summary>
public static class Commands
{
    /// <summary>Shows the product name, version, commit, build date and add-in path.</summary>
    [ExcelCommand(Name = "EmtAbout")]
    public static void EmtAbout()
    {
        try
        {
            var xllPath = ExcelDnaUtil.XllPath;
            StatusBar.Show(
                $"{ProductInfo.Name} {ProductInfo.Version} (commit {ProductInfo.Commit}, built {ProductInfo.BuildDate}) - {xllPath}");
            MessageBox.Show(
                $"{ProductInfo.Name}\n\n" +
                $"Version: {ProductInfo.Version}\n" +
                $"Commit: {ProductInfo.Commit}\n" +
                $"Built: {ProductInfo.BuildDate}\n" +
                $"Add-in: {xllPath}",
                "About " + ProductInfo.Name,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            StatusBar.Show($"{ProductInfo.Name}: About failed: {ex.Message}");
        }
    }
}
