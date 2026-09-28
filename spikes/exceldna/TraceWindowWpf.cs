using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using ExcelDna.Integration;

namespace EmtSpike
{
    /// <summary>
    /// K4 variant A (focused WPF window, ElementHost modeless keyboard interop) and
    /// variant C (non-activating WPF window + thread keyboard hook, Excel keeps focus).
    /// </summary>
    internal sealed class TraceWindowWpf : Window
    {
        private readonly Trace.Session _s;
        private readonly bool _noActivate;
        private readonly ListBox _list = new ListBox { FontFamily = new FontFamily("Consolas"), FontSize = 12 };
        private readonly CheckBox _react = new CheckBox { Content = "Re-activate after Goto", Margin = new Thickness(4) };
        private bool _suppress;
        private long _keyTs;

        private TraceWindowWpf(Trace.Session s, bool noActivate)
        {
            _s = s;
            _noActivate = noActivate;
            Title = noActivate ? "EMT Spike trace (WPF, non-activating + hook)" : "EMT Spike trace (WPF)";
            Width = 640;
            Height = 320;
            ShowInTaskbar = false;
            ShowActivated = !noActivate;
            WindowStartupLocation = WindowStartupLocation.Manual;
            Left = 200;
            Top = 200;

            _react.IsChecked = s.Reactivate;
            _react.Checked += (o, e) => _s.Reactivate = true;
            _react.Unchecked += (o, e) => _s.Reactivate = false;
            if (noActivate) _react.Visibility = Visibility.Collapsed; // focus stays in Excel by design

            var hint = new TextBlock
            {
                Margin = new Thickness(4),
                Text = noActivate
                    ? "Up/Down: go to ref (hook)   Enter: close   Esc: back to root + close   F2: edit in Excel (keys pass until Enter/Esc)"
                    : "Up/Down: go to ref   Enter: close   Esc: back to root + close",
            };
            var panel = new DockPanel();
            DockPanel.SetDock(_react, Dock.Top);
            DockPanel.SetDock(hint, Dock.Bottom);
            panel.Children.Add(_react);
            panel.Children.Add(hint);
            panel.Children.Add(_list);
            Content = panel;

            _list.ItemsSource = s.Items;
            _list.SelectionChanged += OnSelectionChanged;
            PreviewKeyDown += OnPreviewKeyDown;
            Loaded += OnLoaded;
            Closed += OnClosed;
        }

        public static TraceWindowWpf ShowNew(Trace.Session s, IntPtr excelHwnd, bool noActivate)
        {
            var w = new TraceWindowWpf(s, noActivate);
            var helper = new WindowInteropHelper(w) { Owner = excelHwnd };
            IntPtr hwnd = helper.EnsureHandle();
            if (noActivate)
            {
                long ex = Native.GetWindowLongPtr(hwnd, Native.GWL_EXSTYLE).ToInt64();
                Native.SetWindowLongPtr(hwnd, Native.GWL_EXSTYLE, new IntPtr(ex | Native.WS_EX_NOACTIVATE));
            }
            else
            {
                System.Windows.Forms.Integration.ElementHost.EnableModelessKeyboardInterop(w);
            }

            s.OurHwnd = hwnd;
            s.FrameworkFocus = () => w.IsKeyboardFocusWithin;
            s.ReactivateAction = w.ReactivateNow;
            w.Show();
            if (noActivate) KeyHook.Install(w);
            Log.Write("k4open", new { ui = s.Ui.ToString(), noActivate, hwnd = hwnd.ToInt64(), owner = excelHwnd.ToInt64(), state = s.FocusState() });
            return w;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            _suppress = true;
            _list.SelectedIndex = 0;
            _suppress = false;
            if (!_noActivate) FocusSelected();
            Log.Write("k4loaded", new { ui = _s.Ui.ToString(), state = _s.FocusState() });
        }

        private void FocusSelected()
        {
            _list.UpdateLayout();
            if (_list.ItemContainerGenerator.ContainerFromIndex(Math.Max(0, _list.SelectedIndex)) is ListBoxItem item)
                Keyboard.Focus(item);
            else
                Keyboard.Focus(_list);
        }

        internal void ReactivateNow()
        {
            Activate();
            FocusSelected();
        }

        private void OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Up:
                case Key.Down:
                    _keyTs = Stopwatch.GetTimestamp();
                    break;
                case Key.Enter:
                    e.Handled = true;
                    CloseKeep();
                    break;
                case Key.Escape:
                    e.Handled = true;
                    CloseBack();
                    break;
            }
        }

        private void OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_suppress) return;
            long ts = _keyTs;
            _keyTs = 0;
            _s.Navigate(_list.SelectedIndex, ts);
        }

        /// <summary>Variant C: key routed from the keyboard hook (already posted to the dispatcher).</summary>
        internal void HookKey(int vk, long ts)
        {
            try
            {
                switch (vk)
                {
                    case 0x26: // VK_UP
                    case 0x28: // VK_DOWN
                        int i = Math.Max(0, Math.Min(_list.Items.Count - 1, _list.SelectedIndex + (vk == 0x26 ? -1 : 1)));
                        if (i == _list.SelectedIndex) return;
                        _keyTs = ts;
                        _list.SelectedIndex = i; // -> OnSelectionChanged -> Navigate
                        _list.ScrollIntoView(_list.SelectedItem);
                        break;
                    case 0x0D: CloseKeep(); break;
                    case 0x1B: CloseBack(); break;
                    default: Log.Write("hookKeyIgnored", new { vk, note = "Left/Right expand/collapse not implemented in spike" }); break;
                }
            }
            catch (Exception ex) { Log.Error("TraceWindowWpf.HookKey", ex); }
        }

        /// <summary>Variant C: after F2-edit ends with Enter, re-read the root formula.</summary>
        internal void QueueRebuild()
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    _s.Build();
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        _suppress = true;
                        _list.ItemsSource = _s.Items;
                        _list.SelectedIndex = 0;
                        _suppress = false;
                    }));
                    Log.Write("k4rebuild", new { formula = _s.Formula, refs = _s.Items.Count - 1 });
                }
                catch (Exception ex) { Log.Error("QueueRebuild", ex); }
            });
        }

        private void CloseKeep()
        {
            Log.Write("k4close", new { ui = _s.Ui.ToString(), mode = "enter-keep", index = _list.SelectedIndex });
            Close();
        }

        private void CloseBack()
        {
            Log.Write("k4close", new { ui = _s.Ui.ToString(), mode = "esc-back" });
            _s.GoBack();
            Close();
        }

        private void OnClosed(object sender, EventArgs e)
        {
            if (_noActivate && KeyHook.Target == this) KeyHook.Uninstall("window closed");
        }
    }
}
