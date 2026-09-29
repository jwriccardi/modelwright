using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace EmtSpike
{
    /// <summary>"EMT Spike" ribbon tab. Every button queues its work as a macro so COM/C API calls run in macro context.</summary>
    [ComVisible(true)]
    public class SpikeRibbon : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID) => @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='emtSpikeTab' label='EMT Spike'>
        <group id='gK2' label='K2 Undo'>
          <button id='k2_1' label='1 COM Bold (control)' onAction='OnAction' />
          <button id='k2_2' label='2 ExecuteMso Bold' onAction='OnAction' />
          <button id='k2_3' label='3 ExecuteMso Bold (deferred)' onAction='OnAction' />
          <button id='k2_4' label='4 ExecuteMso PercentStyle' onAction='OnAction' />
          <button id='k2_5' label='5 Copy + PasteSpecial formats' onAction='OnAction' />
          <button id='k2_6' label='6 Copy + ExecuteMso PasteFormatting' onAction='OnAction' />
          <button id='k2_7' label='7 xlcFormatNumber (C API)' onAction='OnAction' />
          <button id='k2_10' label='10 ExecuteMso Bold (ribbon, no macro)' onAction='OnAction' />
          <button id='k2_0' label='Log undo snapshot' onAction='OnAction' />
        </group>
        <group id='gK4' label='K4 Trace'>
          <button id='fixtures' label='Create K4 fixtures' onAction='OnAction' />
          <button id='traceWpf' label='Trace (WPF)' onAction='OnAction' />
          <button id='traceWpfReact' label='Trace (WPF, re-activate)' onAction='OnAction' />
          <button id='traceWinForms' label='Trace (WinForms)' onAction='OnAction' />
          <button id='traceWinFormsReact' label='Trace (WinForms, re-activate)' onAction='OnAction' />
          <button id='traceHook' label='Trace (hook)' onAction='OnAction' />
        </group>
        <group id='gMisc' label='Spike'>
          <button id='override' label='Override (re-register keys)' onAction='OnAction' />
          <button id='openLog' label='Open log folder' onAction='OnAction' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        public void OnAction(IRibbonControl control)
        {
            string id = control.Id;
            try
            {
                if (id == "k2_10") { K2b.RibbonBold(); return; } // K2b: deliberately NOT in macro context
                ExcelAsyncUtil.QueueAsMacro(() => Commands.Ribbon(id));
            }
            catch (Exception ex)
            {
                Log.Error("Ribbon.OnAction " + id, ex);
            }
        }
    }
}
