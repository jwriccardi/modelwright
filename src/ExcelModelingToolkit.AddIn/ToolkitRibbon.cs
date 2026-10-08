using System;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelModelingToolkit.Core.Settings;
using ExcelModelingToolkit.Core.Undo;

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
        <group id='emtAuditGroup' label='Audit'>
          <button id='emtTraceIn' label='Trace In' screentip='Trace the precedents of the active cell' supertip='Default key: Ctrl+Shift+[. Opens the precedents of the active cell. Up and Down go to each precedent (on other sheets and in other workbooks, opening closed ones read-only); Right expands, Left goes back up; Enter closes and stays, Esc closes and returns to the audited cell. F2 edits the selected reference in its formula (Point mode: the arrows pick its replacement; Enter commits), with the window open.' onAction='OnTraceIn' />
          <button id='emtLastAuditedCell' label='Last Audited Cell' screentip='Go back to the last audited cell' supertip='Default key: Ctrl+Shift+\. Goes back to the cell Trace In was last opened on. Press again to go further back (up to 20 audits).' onAction='OnLastAuditedCell' />
        </group>
        <group id='emtToolkitGroup' label='Toolkit'>
          <button id='emtSettings' label='Settings…' screentip='Edit cycles, shortcuts and options' supertip='Edit, reorder and preview the cycles, change shortcuts, and import, export or reset your settings.' onAction='OnSettings' />
          <button id='emtAbout' label='About' screentip='About {name}' onAction='OnAbout' />
          <button id='emtReregisterKeys' label='Re-register shortcuts' screentip='Take back every shortcut from other add-ins' onAction='OnReregisterKeys' />
          <button id='emtOpenSettings' label='Open settings file' screentip='Edit settings.json' onAction='OnOpenSettings' />
          <button id='emtReloadSettings' label='Reload settings' screentip='Apply your changes to settings.json' onAction='OnReloadSettings' />
          <button id='emtUndoFormatting' label='Undo formatting' screentip='Undo our last formatting change' supertip='Undoes the last cycle even when Ctrl+Z would go to Excel (Excel has newer changes to undo).' onAction='OnUndoFormatting' />
          <button id='emtRedoFormatting' label='Redo formatting' screentip='Redo our last undone formatting change' supertip='Redoes the last undone formatting change. Ctrl+Y drops this redo history once you change something in Excel after an undo.' onAction='OnRedoFormatting' />
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

    /// <summary>Ribbon callback for Trace In.</summary>
    public void OnTraceIn(IRibbonControl control) => ExcelAsyncUtil.QueueAsMacro(() => TraceCommand.TraceIn("ribbon"));

    /// <summary>Ribbon callback for Last Audited Cell.</summary>
    public void OnLastAuditedCell(IRibbonControl control) =>
        ExcelAsyncUtil.QueueAsMacro(() => TraceCommand.LastAuditedCell("ribbon"));

    /// <summary>
    /// Ribbon callback for Settings: the dialog runs as a macro, so its preview can call Excel. The click time lets
    /// a second click, queued while the dialog was open, be ignored once it has closed.
    /// </summary>
    public void OnSettings(IRibbonControl control)
    {
        var clickedAt = Environment.TickCount;
        ExcelAsyncUtil.QueueAsMacro(() => Commands.ShowSettings(clickedAt));
    }

    /// <summary>Ribbon callback for the About button.</summary>
    public void OnAbout(IRibbonControl control) => ExcelAsyncUtil.QueueAsMacro(Commands.EmtAbout);

    /// <summary>Ribbon callback for Re-register shortcuts.</summary>
    public void OnReregisterKeys(IRibbonControl control) => ExcelAsyncUtil.QueueAsMacro(Commands.EmtReregisterKeys);

    /// <summary>Ribbon callback for Open settings file.</summary>
    public void OnOpenSettings(IRibbonControl control) => ExcelAsyncUtil.QueueAsMacro(Commands.EmtOpenSettings);

    /// <summary>Ribbon callback for Reload settings.</summary>
    public void OnReloadSettings(IRibbonControl control) => ExcelAsyncUtil.QueueAsMacro(Commands.EmtReloadSettings);

    /// <summary>Ribbon callback for Undo formatting: always our stack, whatever Excel's own undo holds.</summary>
    public void OnUndoFormatting(IRibbonControl control) =>
        ExcelAsyncUtil.QueueAsMacro(() => UndoCommand.Run(UndoKey.Undo, "ribbon"));

    /// <summary>Ribbon callback for Redo formatting: always acts on our redo stack (<see cref="UndoCommand"/>).</summary>
    public void OnRedoFormatting(IRibbonControl control) =>
        ExcelAsyncUtil.QueueAsMacro(() => UndoCommand.Run(UndoKey.Redo, "ribbon"));
}
