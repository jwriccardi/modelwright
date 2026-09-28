using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace EmtSpike
{
    /// <summary>K4 variant B: modeless WinForms form owned by Excel's HWND (via NativeWindow).</summary>
    internal sealed class TraceForm : Form
    {
        private readonly Trace.Session _s;
        private readonly ListBox _list = new ListBox { Dock = DockStyle.Fill, Font = new Font("Consolas", 9f), IntegralHeight = false };
        private readonly CheckBox _react = new CheckBox { Dock = DockStyle.Top, Text = "Re-activate after Goto" };
        private NativeWindow _owner;
        private bool _suppress;
        private long _keyTs;

        public TraceForm(Trace.Session s)
        {
            _s = s;
            Text = "EMT Spike trace (WinForms)";
            ClientSize = new Size(640, 300);
            StartPosition = FormStartPosition.Manual;
            Location = new Point(220, 220);
            ShowInTaskbar = false;
            KeyPreview = true;

            _react.Checked = s.Reactivate;
            _react.CheckedChanged += (o, e) => _s.Reactivate = _react.Checked;
            var hint = new Label { Dock = DockStyle.Bottom, Height = 20, Text = "Up/Down: go to ref   Enter: close   Esc: back to root + close" };

            Controls.Add(_list);
            Controls.Add(hint);
            Controls.Add(_react);
            foreach (var item in s.Items) _list.Items.Add(item);
            _list.SelectedIndexChanged += OnSelectedIndexChanged;
        }

        public void ShowOwned(IntPtr excelHwnd)
        {
            _owner = new NativeWindow();
            _owner.AssignHandle(excelHwnd);
            Show(_owner);
            _suppress = true;
            _list.SelectedIndex = 0;
            _suppress = false;
            _list.Focus();

            _s.OurHwnd = Handle;
            _s.FrameworkFocus = () => ContainsFocus;
            _s.ReactivateAction = () => { Activate(); _list.Focus(); };
            Log.Write("k4open", new { ui = _s.Ui.ToString(), hwnd = Handle.ToInt64(), owner = excelHwnd.ToInt64(), state = _s.FocusState() });
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Up:
                case Keys.Down:
                    _keyTs = Stopwatch.GetTimestamp();
                    break;
                case Keys.Enter:
                    e.Handled = true;
                    Log.Write("k4close", new { ui = _s.Ui.ToString(), mode = "enter-keep", index = _list.SelectedIndex });
                    Close();
                    break;
                case Keys.Escape:
                    e.Handled = true;
                    Log.Write("k4close", new { ui = _s.Ui.ToString(), mode = "esc-back" });
                    _s.GoBack();
                    Close();
                    break;
            }
            base.OnKeyDown(e);
        }

        private void OnSelectedIndexChanged(object sender, EventArgs e)
        {
            if (_suppress) return;
            long ts = _keyTs;
            _keyTs = 0;
            _s.Navigate(_list.SelectedIndex, ts);
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _owner?.ReleaseHandle();
            _owner = null;
            base.OnFormClosed(e);
        }
    }
}
