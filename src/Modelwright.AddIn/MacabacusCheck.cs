using System;
using System.Drawing;
using System.Windows.Forms;
using ExcelDna.Integration;
using Modelwright.Core.Settings;

namespace Modelwright.AddIn;

/// <summary>
/// Running alongside Macabacus, which binds the same keyboard shortcuts. Shortly after the add-in has loaded (on a
/// WinForms timer, outside macro context, so startup never waits for it) it looks for Macabacus among Excel's add-ins
/// and does what <see cref="Coexistence.Decide"/> says: binds our keys again, so ours are the most recent bindings
/// (those win in Excel) over the ones Macabacus made while it loaded, and the first time, shows a notice that offers
/// to switch our shortcuts off instead. Logs one <c>Coexist</c> line. Never throws.
/// </summary>
/// <remarks>
/// The notice is never shown with the environment variable <see cref="Coexistence.NoNoticesVariable"/> set to
/// <c>1</c>, nor once ui-state.json says it was shown (<see cref="UiStateStore.MarkMacabacusNoticeShown"/>), which is
/// how the Excel smoke tests keep it away.
/// </remarks>
internal static class MacabacusCheck
{
    /// <summary>How long after the add-in loads to look: Excel loads the COM add-ins (Macabacus) around us.</summary>
    private const int DelayMilliseconds = 3000;

    private static Timer? _timer;

    /// <summary>Has the check run <see cref="DelayMilliseconds"/> from now, from Excel's message loop. Never throws.</summary>
    public static void Schedule()
    {
        try
        {
            Cancel();
            var timer = new Timer { Interval = DelayMilliseconds };
            timer.Tick += (sender, e) =>
            {
                Cancel();
                Run();
            };
            _timer = timer;
            timer.Start();
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("Coexist", "not scheduled: " + ex.Message);
        }
    }

    /// <summary>Stops a check that has not run yet (the add-in is unloading). Never throws.</summary>
    public static void Cancel()
    {
        try
        {
            _timer?.Stop();
            _timer?.Dispose();
        }
        catch (Exception)
        {
            // A timer that will not stop only runs a check that changes nothing once the add-in has gone.
        }

        _timer = null;
    }

    // On the timer tick: detection through COM, then the actions queued as a macro (binding keys needs the C API).
    private static void Run()
    {
        try
        {
            var detected = Detect(out var found);
            var noticeShown = UiStateStore.Load().MacabacusNoticeShown;
            var suppressed = Coexistence.NoticesSuppressed(Environment.GetEnvironmentVariable(Coexistence.NoNoticesVariable));
            var shortcutsOn = Session.Settings.UseKeyboardShortcuts;
            var actions = Coexistence.Decide(detected, shortcutsOn, noticeShown, suppressed);
            var notice =
                !detected ? "not needed" :
                !shortcutsOn ? "not needed: shortcuts off" :
                noticeShown ? "shown before" :
                suppressed ? "suppressed (" + Coexistence.NoNoticesVariable + "=1)" :
                "shown";
            if (actions == CoexistenceActions.None)
            {
                Log(detected, reregistered: false, found, notice);
                return;
            }

            ExcelAsyncUtil.QueueAsMacro(() => Act(actions, found, notice));
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("Coexist", "failed: " + ex.Message);
        }
    }

    private static void Act(CoexistenceActions actions, string found, string notice)
    {
        try
        {
            // The shortcuts may have been switched off since the check was decided (the ribbon, say).
            var shortcutsOn = Session.Settings.UseKeyboardShortcuts;
            var reregistered = false;
            if ((actions & CoexistenceActions.Reregister) != 0 && shortcutsOn)
            {
                KeyBindings.Apply(Session.Settings);
                reregistered = true;
            }

            var show = (actions & CoexistenceActions.ShowNotice) != 0 && shortcutsOn;
            Log(detected: true, reregistered, found, show || !shortcutsOn ? notice : "not needed: shortcuts off");
            if (!show)
            {
                return;
            }

            // Remembered first, so a notice that fails to show is not tried again at every start.
            UiStateStore.MarkMacabacusNoticeShown();
            var keepMacabacus = Notice.Ask();
            DiagnosticsLog.Write("CoexistNotice", keepMacabacus ? "choice=keep macabacus" : "choice=use modelwright");
            if (keepMacabacus)
            {
                Commands.SetKeyboardShortcuts(false, "notice");
            }
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("Coexist", "failed: " + ex.Message);
        }
    }

    private static void Log(bool detected, bool reregistered, string found, string notice) =>
        DiagnosticsLog.Write(
            "Coexist",
            "macabacus=" + (detected ? "true" : "false"),
            "reregistered=" + (reregistered ? "true" : "false"),
            "found=" + found,
            "notice=" + notice);

    /// <summary>
    /// True if Macabacus's COM add-in, which binds its keys, is loaded (<see cref="Coexistence.IsMacabacusComAddIn"/>);
    /// Macabacus.xlam binds none, so it is not looked for. <paramref name="found"/> names it, for the log. An add-in
    /// Excel will not describe is skipped.
    /// </summary>
    private static bool Detect(out string found)
    {
        dynamic app = ExcelDnaUtil.Application;
        dynamic comAddIns = app.COMAddIns;
        int comCount = comAddIns.Count;
        for (var i = 1; i <= comCount; i++)
        {
            try
            {
                dynamic item = comAddIns.Item(i);
                string? progId = item.ProgId;
                bool connected = item.Connect;
                if (Coexistence.IsMacabacusComAddIn(progId, connected))
                {
                    found = "COM add-in " + progId;
                    return true;
                }
            }
            catch (Exception)
            {
                // An add-in Excel cannot describe (a broken registration) is not Macabacus as far as we can tell.
            }
        }

        found = "none";
        return false;
    }

    /// <summary>The one-time notice: a small modal dialog owned by Excel's window.</summary>
    private sealed class Notice : Form
    {
        private Notice()
        {
            Text = ProductInfo.Name + " and Macabacus";
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96F, 96F);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ShowIcon = false;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;

            var text = new Label
            {
                Text =
                    "Macabacus is also loaded. Both add-ins use the same keyboard shortcuts, so only one can answer " +
                    $"each key. {ProductInfo.Name}'s shortcuts are on and now answer Ctrl+Shift+1, Ctrl+', " +
                    "Ctrl+Shift+[ and the others; Macabacus's commands stay on its ribbon. You can switch them off " +
                    $"at any time with {ProductInfo.Name} \u203A Shortcuts (the ribbon keeps working); Macabacus gets " +
                    "the keys back the next time Excel starts, and until then they do Excel's usual thing.",
                AutoSize = true,
                MaximumSize = new Size(420, 0),
                UseMnemonic = false,
                Margin = new Padding(0, 0, 0, 12),
            };
            var use = new Button { Text = $"Use {ProductInfo.Name}'s shortcuts", AutoSize = true, DialogResult = DialogResult.OK, UseMnemonic = false };
            var keep = new Button { Text = "Keep Macabacus's shortcuts", AutoSize = true, DialogResult = DialogResult.No, UseMnemonic = false };

            // Enter and Escape both mean the first button: our shortcuts stay on.
            AcceptButton = use;
            CancelButton = use;

            var buttons = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Anchor = AnchorStyles.Right, Margin = new Padding(0) };
            buttons.Controls.Add(use);
            buttons.Controls.Add(keep);

            var layout = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 1, RowCount = 2, Padding = new Padding(12), Dock = DockStyle.Fill };
            layout.Controls.Add(text, 0, 0);
            layout.Controls.Add(buttons, 0, 1);
            Controls.Add(layout);
            ActiveControl = use;
        }

        /// <summary>
        /// Shows the notice, modal and owned by Excel's window. True if "Keep Macabacus's shortcuts" was chosen;
        /// false for the first button, Enter, Escape or closing it. Call on Excel's main thread.
        /// </summary>
        public static bool Ask()
        {
            // Showing a form would otherwise install a WinForms SynchronizationContext on Excel's main thread.
            WindowsFormsSynchronizationContext.AutoInstall = false;
            Application.EnableVisualStyles();
            var previousDpiContext = SettingsDialog.DpiContext.EnterSystemAware();
            try
            {
                using var notice = new Notice();
                return notice.ShowDialog(SettingsDialog.ExcelOwner()) == DialogResult.No;
            }
            finally
            {
                SettingsDialog.DpiContext.Restore(previousDpiContext);
            }
        }
    }
}
