using System.Runtime.InteropServices;
using System.Security;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace ExcelModelingToolkit.AddIn;

/// <summary>The "Modeling Toolkit" ribbon tab. Callbacks queue their work as a macro so the C API is available.</summary>
[ComVisible(true)]
public class ToolkitRibbon : ExcelRibbon
{
    /// <inheritdoc />
    public override string GetCustomUI(string RibbonID)
    {
        var name = SecurityElement.Escape(ProductInfo.Name);
        return $@"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='emtTab' label='{name}'>
        <group id='emtHelpGroup' label='Help'>
          <button id='emtAbout' label='About' screentip='About {name}' onAction='OnAbout' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
    }

    /// <summary>Ribbon callback for the About button.</summary>
    public void OnAbout(IRibbonControl control) => ExcelAsyncUtil.QueueAsMacro(Commands.EmtAbout);
}
