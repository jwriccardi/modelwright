using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelModelingToolkit.Core.Settings;

namespace ExcelModelingToolkit.AddIn;

/// <summary>The "Modeling Toolkit" ribbon tab. Callbacks queue their work as a macro so the C API is available.</summary>
[ComVisible(true)]
public class ToolkitRibbon : ExcelRibbon
{
    /// <summary>Format group buttons: action id (also the button tag), label, screentip.</summary>
    private static readonly string[][] FormatButtons =
    {
        new[] { ActionIds.NumberCycle, "Number", "Cycle number formats" },
        new[] { ActionIds.DateCycle, "Date", "Cycle date formats" },
        new[] { ActionIds.CurrencyCycle, "Currency", "Cycle currency formats" },
        new[] { ActionIds.PercentCycle, "Percent", "Cycle percent formats" },
        new[] { ActionIds.MultipleCycle, "Multiple", "Cycle multiple formats" },
        new[] { ActionIds.BinaryCycle, "Binary", "Cycle binary formats (Yes/No, On/Off...)" },
        new[] { ActionIds.RatioCycle, "Ratio", "Cycle ratio and fraction formats" },
        new[] { ActionIds.FontColorCycle, "Font Color", "Cycle font colors" },
        new[] { ActionIds.FillColorCycle, "Fill Color", "Cycle fill colors" },
        new[] { ActionIds.BlueBlackToggle, "Blue/Black", "Toggle the font between blue and black" },
    };

    /// <inheritdoc />
    public override string GetCustomUI(string RibbonID)
    {
        var name = SecurityElement.Escape(ProductInfo.Name);
        var format = new StringBuilder();
        foreach (var button in FormatButtons)
        {
            format.Append($@"
          <button id='emt{button[0]}' tag='{button[0]}' label='{SecurityElement.Escape(button[1])}' screentip='{SecurityElement.Escape(button[2])}' onAction='OnCycle' />");
        }

        return $@"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='emtTab' label='{name}'>
        <group id='emtFormatGroup' label='Format'>{format}
        </group>
        <group id='emtToolkitGroup' label='Toolkit'>
          <button id='emtAbout' label='About' screentip='About {name}' onAction='OnAbout' />
          <button id='emtReregisterKeys' label='Re-register shortcuts' screentip='Take back every shortcut from other add-ins' onAction='OnReregisterKeys' />
          <button id='emtOpenSettings' label='Open settings file' screentip='Edit settings.json' onAction='OnOpenSettings' />
          <button id='emtReloadSettings' label='Reload settings' screentip='Apply your changes to settings.json' onAction='OnReloadSettings' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
    }

    /// <summary>Ribbon callback for the Format buttons; the button's tag is the cycle's action id.</summary>
    public void OnCycle(IRibbonControl control)
    {
        var actionId = control.Tag;
        ExcelAsyncUtil.QueueAsMacro(() => Commands.RunCycleFromRibbon(actionId));
    }

    /// <summary>Ribbon callback for the About button.</summary>
    public void OnAbout(IRibbonControl control) => ExcelAsyncUtil.QueueAsMacro(Commands.EmtAbout);

    /// <summary>Ribbon callback for Re-register shortcuts.</summary>
    public void OnReregisterKeys(IRibbonControl control) => ExcelAsyncUtil.QueueAsMacro(Commands.EmtReregisterKeys);

    /// <summary>Ribbon callback for Open settings file.</summary>
    public void OnOpenSettings(IRibbonControl control) => ExcelAsyncUtil.QueueAsMacro(Commands.EmtOpenSettings);

    /// <summary>Ribbon callback for Reload settings.</summary>
    public void OnReloadSettings(IRibbonControl control) => ExcelAsyncUtil.QueueAsMacro(Commands.EmtReloadSettings);
}
