using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Markup;
using System.Windows.Media;
using ExcelModelingToolkit.Core.Trace;

namespace ExcelModelingToolkit.AddIn;

/// <summary>
/// The Trace In window (variant C of spike K4): a WPF window that never takes the focus (<c>WS_EX_NOACTIVATE</c>),
/// owned by Excel's active workbook window, so Excel keeps the keyboard and F2 edits the active cell with the
/// window open. Its keys come from <see cref="TraceKeyHook"/>; WPF's own keyboard handling is not used (it is
/// unreliable on Excel's thread). The mouse works: a click selects a row, a double-click or a click on the
/// expander expands or collapses it.
/// </summary>
/// <remarks>
/// Only displays what <see cref="TraceSession"/> gives it and reports clicks back; all positions are in screen
/// pixels through Win32, so they agree with the monitor work areas whatever the DPI. Main thread only.
/// </remarks>
internal sealed class TraceWindow : Window
{
    // The window's content. Loaded at run time (no markup compilation); buttons and the list never take the focus.
    private const string ContentXaml = @"
<DockPanel xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
           xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'>
  <DockPanel.Resources>
    <Style TargetType='Button'>
      <Setter Property='Focusable' Value='False'/>
      <Setter Property='MinWidth' Value='72'/>
      <Setter Property='Margin' Value='6,0,0,0'/>
      <Setter Property='Padding' Value='8,2'/>
    </Style>
    <Style TargetType='ToggleButton'>
      <Setter Property='Focusable' Value='False'/>
    </Style>
    <Style TargetType='GridViewColumnHeader'>
      <Setter Property='Focusable' Value='False'/>
      <Setter Property='HorizontalContentAlignment' Value='Left'/>
      <Setter Property='Padding' Value='4,1'/>
    </Style>
  </DockPanel.Resources>
  <Border DockPanel.Dock='Top' BorderBrush='{x:Static SystemColors.ControlDarkBrush}' BorderThickness='0,0,0,1' Padding='6,4'>
    <DockPanel>
      <ToggleButton x:Name='WrapToggle' DockPanel.Dock='Right' VerticalAlignment='Top' Margin='6,0,0,0' Padding='6,0'
                    Content='&#x21B5;' ToolTip='Wrap formula text'/>
      <StackPanel>
        <ScrollViewer x:Name='FormulaScroll' Focusable='False' MaxHeight='96'
                      HorizontalScrollBarVisibility='Auto' VerticalScrollBarVisibility='Disabled'>
          <TextBlock x:Name='FormulaText' FontFamily='Consolas' FontSize='13' TextWrapping='NoWrap'/>
        </ScrollViewer>
        <TextBlock x:Name='NoteText' Margin='0,2,0,0' FontStyle='Italic' TextWrapping='Wrap' Visibility='Collapsed'
                   Foreground='{x:Static SystemColors.GrayTextBrush}'/>
      </StackPanel>
    </DockPanel>
  </Border>
  <DockPanel DockPanel.Dock='Bottom' Margin='6'>
    <Button x:Name='GearButton' DockPanel.Dock='Left' Margin='0' MinWidth='28' Padding='4,0' Content='&#x2699;' ToolTip='Options'>
      <Button.ContextMenu>
        <ContextMenu>
          <MenuItem x:Name='WrapMenu' Header='Wrap formula text' IsCheckable='True'/>
          <MenuItem x:Name='EvaluateMenu' Header='Evaluate functions &amp; groups' InputGestureText='Ctrl+E' IsEnabled='False'
                    ToolTip='Coming in a later version' ToolTipService.ShowOnDisabled='True'/>
        </ContextMenu>
      </Button.ContextMenu>
    </Button>
    <Button x:Name='CancelButton' DockPanel.Dock='Right' Content='Cancel' ToolTip='Close and return to the audited cell (Esc)'/>
    <Button x:Name='OkButton' DockPanel.Dock='Right' Content='OK' ToolTip='Close and stay on the current cell (Enter)'/>
    <TextBlock x:Name='StatusText' Margin='8,0' VerticalAlignment='Center' TextTrimming='CharacterEllipsis'
               Foreground='{x:Static SystemColors.GrayTextBrush}'/>
  </DockPanel>
  <ListView x:Name='Tree' Focusable='False' BorderThickness='0' SelectionMode='Single'
            VirtualizingPanel.IsVirtualizing='True' VirtualizingPanel.VirtualizationMode='Recycling'>
    <ListView.ItemContainerStyle>
      <Style TargetType='ListViewItem'>
        <Setter Property='Focusable' Value='False'/>
        <Setter Property='Template'>
          <Setter.Value>
            <ControlTemplate TargetType='ListViewItem'>
              <Border x:Name='Row' Background='Transparent' Padding='0,1'>
                <GridViewRowPresenter Columns='{TemplateBinding GridView.ColumnCollection}' Content='{TemplateBinding Content}'/>
              </Border>
              <ControlTemplate.Triggers>
                <Trigger Property='IsSelected' Value='True'>
                  <Setter TargetName='Row' Property='Background' Value='{x:Static SystemColors.HighlightBrush}'/>
                  <Setter Property='Foreground' Value='{x:Static SystemColors.HighlightTextBrush}'/>
                </Trigger>
              </ControlTemplate.Triggers>
            </ControlTemplate>
          </Setter.Value>
        </Setter>
      </Style>
    </ListView.ItemContainerStyle>
    <ListView.View>
      <GridView>
        <GridViewColumn Header='Precedents' Width='300'>
          <GridViewColumn.CellTemplate>
            <DataTemplate>
              <StackPanel Orientation='Horizontal'>
                <FrameworkElement Width='{Binding IndentWidth}'/>
                <TextBlock Text='{Binding Glyph}' Width='14' TextAlignment='Center'/>
                <TextBlock Text='{Binding Icon}' Width='22' TextAlignment='Center' ToolTip='{Binding KindName}'/>
                <TextBlock Text='{Binding Label}' FontStyle='{Binding LabelStyle}'/>
                <TextBlock Text='{Binding Badge}' Margin='6,0,0,0' FontStyle='Italic' Opacity='0.7'/>
              </StackPanel>
            </DataTemplate>
          </GridViewColumn.CellTemplate>
        </GridViewColumn>
        <GridViewColumn Header='Argument' Width='90' DisplayMemberBinding='{Binding Argument}'/>
        <GridViewColumn Header='Value' Width='170' DisplayMemberBinding='{Binding Value}'/>
      </GridView>
    </ListView.View>
  </ListView>
</DockPanel>";

    private const int GwlExStyle = -20;
    private const int GwlpHwndParent = -8;
    private const long WsExNoActivate = 0x08000000L;
    private const uint SwpNoSize = 0x0001;
    private const uint SwpNoMove = 0x0002;
    private const uint SwpNoZOrder = 0x0004;
    private const uint SwpNoActivate = 0x0010;
    private const uint SwpShowWindow = 0x0040;
    private const uint MonitorDefaultToNearest = 2;

    // GridViewRowPresenter's margin before each cell's content.
    private const double CellMargin = 6;

    private readonly TextBlock _formulaText;
    private readonly ScrollViewer _formulaScroll;
    private readonly TextBlock _noteText;
    private readonly TextBlock _statusText;
    private readonly ToggleButton _wrapToggle;
    private readonly MenuItem _wrapMenu;
    private readonly ListView _tree;
    private IntPtr _hwnd;
    private IntPtr _owner;
    private bool _closingByCode;
    private bool _syncingWrap;
    private IntPtr _focusBeforeMenu;
    private string _formula = string.Empty;
    private IReadOnlyList<FormulaSegment> _segments = new FormulaSegment[0];

    public TraceWindow()
    {
        Title = "Trace In";
        ShowInTaskbar = false;
        ShowActivated = false;
        WindowStyle = WindowStyle.ToolWindow;
        ResizeMode = ResizeMode.CanResize;
        WindowStartupLocation = WindowStartupLocation.Manual;
        Background = SystemColors.WindowBrush;
        FontSize = 12;

        var content = (FrameworkElement)XamlReader.Parse(ContentXaml);
        Content = content;
        _formulaText = (TextBlock)content.FindName("FormulaText");
        _formulaScroll = (ScrollViewer)content.FindName("FormulaScroll");
        _noteText = (TextBlock)content.FindName("NoteText");
        _statusText = (TextBlock)content.FindName("StatusText");
        _wrapToggle = (ToggleButton)content.FindName("WrapToggle");
        _wrapMenu = (MenuItem)content.FindName("WrapMenu");
        _tree = (ListView)content.FindName("Tree");
        var gear = (Button)content.FindName("GearButton");
        var ok = (Button)content.FindName("OkButton");
        var cancel = (Button)content.FindName("CancelButton");

        _wrapToggle.Checked += (s, e) => Safe(() => SetWrap(true, byUser: true));
        _wrapToggle.Unchecked += (s, e) => Safe(() => SetWrap(false, byUser: true));
        _wrapMenu.Checked += (s, e) => Safe(() => SetWrap(true, byUser: true));
        _wrapMenu.Unchecked += (s, e) => Safe(() => SetWrap(false, byUser: true));
        // A WPF menu takes the keyboard focus while it is open (the only control here that does). Remember where the
        // focus was (Excel's grid: this window never has it) and give it back when the menu closes, so F2 and typing
        // go to Excel again. Right-click opening is off, so the menu only opens here.
        ContextMenuService.SetIsEnabled(gear, false);
        gear.Click += (s, e) => Safe(() =>
        {
            _focusBeforeMenu = NativeMethods.GetFocus();
            gear.ContextMenu.PlacementTarget = gear;
            gear.ContextMenu.Placement = PlacementMode.Top;
            gear.ContextMenu.IsOpen = true;
        });
        gear.ContextMenu.Closed += (s, e) => Safe(() =>
            Dispatcher.BeginInvoke(new Action(() => Safe(RestoreExcelFocus))));
        ok.Click += (s, e) => Safe(() => OkClicked?.Invoke());
        cancel.Click += (s, e) => Safe(() => CancelClicked?.Invoke());
        _tree.PreviewMouseLeftButtonDown += (s, e) => Safe(() => OnRowMouseDown(e));
        LocationChanged += (s, e) => Safe(RememberBounds);
        SizeChanged += (s, e) => Safe(RememberBounds);
        Closing += (s, e) => Safe(() =>
        {
            // The title bar's close button is Cancel; the session closes the window itself once it has gone back.
            if (!_closingByCode)
            {
                e.Cancel = true;
                CancelClicked?.Invoke();
            }
        });
    }

    /// <summary>A row was clicked: its index, whether it was a double-click, and whether on the expander.</summary>
    public event Action<int, bool, bool>? RowClicked;

    /// <summary>OK was clicked.</summary>
    public event Action? OkClicked;

    /// <summary>Cancel or the title bar's close button was clicked.</summary>
    public event Action? CancelClicked;

    /// <summary>The wrap setting was changed with the mouse.</summary>
    public event Action<bool>? WrapChanged;

    /// <summary>The window's handle (zero until shown).</summary>
    public IntPtr Handle => _hwnd;

    /// <summary>The Excel window that owns this one.</summary>
    public IntPtr OwnerHandle => _owner;

    /// <summary>The window's last known bounds in screen pixels (kept up to date as it moves), or null before it is shown.</summary>
    public WindowRect? LastBounds { get; private set; }

    /// <summary>
    /// Shows the window without activating it, owned by <paramref name="owner"/> (Excel's active window), at
    /// <paramref name="remembered"/> brought back on screen if needed, or at the default place.
    /// </summary>
    public void ShowFor(IntPtr owner, WindowRect? remembered, bool wrap)
    {
        _owner = owner;
        var helper = new WindowInteropHelper(this) { Owner = owner };
        _hwnd = helper.EnsureHandle();
        var style = NativeMethods.GetWindowLongPtr(_hwnd, GwlExStyle).ToInt64();
        NativeMethods.SetWindowLongPtr(_hwnd, GwlExStyle, new IntPtr(style | WsExNoActivate));
        SetWrap(wrap, byUser: false);

        var workArea = WorkAreaOf(owner);
        var scale = Scale();
        var bounds = remembered is WindowRect saved
            ? TraceWindowGeometry.EnsureVisible(saved, AllWorkAreas(), workArea, scale)
            : TraceWindowGeometry.DefaultBounds(WindowBoundsOf(owner) ?? workArea, workArea, scale);
        Place(bounds);
        Show();
        Place(bounds);
    }

    /// <summary>Re-owns the window to another Excel window (after a Goto into another workbook) and raises it, without activating it.</summary>
    public void ReOwn(IntPtr owner)
    {
        if (owner == IntPtr.Zero || owner == _owner || _hwnd == IntPtr.Zero)
        {
            return;
        }

        NativeMethods.SetWindowLongPtr(_hwnd, GwlpHwndParent, owner);
        NativeMethods.SetWindowPos(_hwnd, IntPtr.Zero, 0, 0, 0, 0, SwpNoMove | SwpNoSize | SwpNoActivate | SwpShowWindow);
        _owner = owner;
    }

    /// <summary>Applies a move, snap or resize command (<see cref="TraceKeys.IsWindowCommand"/>). Returns the new bounds.</summary>
    public WindowRect ApplyWindowCommand(TraceKeyCommand command)
    {
        var current = WindowBoundsOf(_hwnd) ?? LastBounds ?? new WindowRect(0, 0, 0, 0);
        var bounds = TraceWindowGeometry.Apply(command, current, WorkAreaOf(_hwnd), AllWorkAreas(), Scale());
        if (bounds != current)
        {
            Place(bounds);
        }

        return bounds;
    }

    /// <summary>Shows the audited formula with its references colored, and a note under it (or none).</summary>
    public void SetFormula(string formula, IReadOnlyList<FormulaSegment> segments, string? note)
    {
        _formula = formula;
        _segments = segments;
        RenderFormula();
        _noteText.Text = note ?? string.Empty;
        _noteText.Visibility = string.IsNullOrEmpty(note) ? Visibility.Collapsed : Visibility.Visible;
    }

    /// <summary>Replaces the rows and selects one.</summary>
    public void SetRows(IReadOnlyList<TraceRow> rows, int selected)
    {
        _tree.ItemsSource = rows;
        Select(selected);
    }

    /// <summary>Selects a row and scrolls it into view.</summary>
    public void Select(int index)
    {
        if (index < 0 || index >= _tree.Items.Count)
        {
            return;
        }

        _tree.SelectedIndex = index;
        _tree.ScrollIntoView(_tree.Items[index]);
    }

    /// <summary>The text in the footer.</summary>
    public void SetStatus(string text) => _statusText.Text = text;

    /// <summary>Closes the window (the session's own close: no Cancel is raised).</summary>
    public void CloseWindow()
    {
        _closingByCode = true;
        Close();
    }

    private void OnRowMouseDown(MouseButtonEventArgs e)
    {
        if (ItemsControl.ContainerFromElement(_tree, e.OriginalSource as DependencyObject) is not ListViewItem item)
        {
            return; // the header, a scroll bar: WPF handles it
        }

        // Handled here so WPF does not try to give the list the focus (that would activate the window).
        e.Handled = true;
        var index = _tree.ItemContainerGenerator.IndexFromContainer(item);
        if (index < 0 || item.Content is not TraceRow row)
        {
            return;
        }

        var x = e.GetPosition(item).X;
        // GridViewRowPresenter indents each cell by its 6-pixel margin; the expander is the first 14 pixels after the indent.
        var expanderLeft = CellMargin + row.IndentWidth;
        var onExpander = row.Glyph.Length > 0 && x >= expanderLeft - 2 && x < expanderLeft + 16;
        RowClicked?.Invoke(index, e.ClickCount >= 2, onExpander);
    }

    private void SetWrap(bool wrap, bool byUser)
    {
        if (_syncingWrap)
        {
            return;
        }

        _syncingWrap = true;
        try
        {
            var changed = _wrapToggle.IsChecked != wrap || _wrapMenu.IsChecked != wrap;
            _wrapToggle.IsChecked = wrap;
            _wrapMenu.IsChecked = wrap;
            _formulaText.TextWrapping = wrap ? TextWrapping.Wrap : TextWrapping.NoWrap;
            _formulaScroll.HorizontalScrollBarVisibility = wrap ? ScrollBarVisibility.Disabled : ScrollBarVisibility.Auto;
            _formulaScroll.VerticalScrollBarVisibility = wrap ? ScrollBarVisibility.Auto : ScrollBarVisibility.Disabled;
            RenderFormula();
            if (changed && byUser)
            {
                WrapChanged?.Invoke(wrap);
            }
        }
        finally
        {
            _syncingWrap = false;
        }
    }

    // Unwrapped, line breaks (Alt+Enter) show as spaces so the formula stays on one line.
    private void RenderFormula()
    {
        var wrap = _formulaText.TextWrapping == TextWrapping.Wrap;
        _formulaText.Inlines.Clear();
        var segments = _segments.Count > 0 ? _segments : new[] { new FormulaSegment(_formula, 0, -1) };
        foreach (var segment in segments)
        {
            var text = wrap ? segment.Text.Replace("\r\n", "\n") : segment.Text.Replace("\r\n", " ").Replace('\n', ' ').Replace('\r', ' ');
            var run = new Run(text);
            if (segment.IsReference)
            {
                var rgb = FormulaColoring.Palette[segment.ColorIndex];
                run.Foreground = new SolidColorBrush(Color.FromRgb((byte)(rgb >> 16), (byte)(rgb >> 8), (byte)rgb));
            }

            _formulaText.Inlines.Add(run);
        }
    }

    // After the gear menu closes: Excel's window to the front again and the keyboard back on the grid it was on.
    private void RestoreExcelFocus()
    {
        if (_owner != IntPtr.Zero)
        {
            NativeMethods.SetForegroundWindow(_owner);
        }

        if (_focusBeforeMenu != IntPtr.Zero && NativeMethods.IsWindow(_focusBeforeMenu))
        {
            NativeMethods.SetFocus(_focusBeforeMenu);
        }

        _focusBeforeMenu = IntPtr.Zero;
    }

    private void RememberBounds()
    {
        if (_hwnd != IntPtr.Zero && WindowBoundsOf(_hwnd) is WindowRect bounds && bounds.Width > 0 && bounds.Height > 0)
        {
            LastBounds = bounds;
        }
    }

    private void Place(WindowRect bounds)
    {
        NativeMethods.SetWindowPos(_hwnd, IntPtr.Zero, bounds.Left, bounds.Top, bounds.Width, bounds.Height, SwpNoZOrder | SwpNoActivate);
        LastBounds = bounds;
    }

    // The window's DPI over 96 (1 on Windows before 10 1607, where the call does not exist).
    private double Scale()
    {
        try
        {
            var dpi = _hwnd == IntPtr.Zero ? 0 : NativeMethods.GetDpiForWindow(_hwnd);
            return dpi > 0 ? dpi / 96.0 : 1.0;
        }
        catch (EntryPointNotFoundException)
        {
            return 1.0;
        }
    }

    private static WindowRect? WindowBoundsOf(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero || !NativeMethods.GetWindowRect(hwnd, out var rect))
        {
            return null;
        }

        return new WindowRect(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top);
    }

    private static WindowRect WorkAreaOf(IntPtr hwnd)
    {
        var monitor = NativeMethods.MonitorFromWindow(hwnd, MonitorDefaultToNearest);
        var info = new NativeMethods.MonitorInfo { Size = Marshal.SizeOf(typeof(NativeMethods.MonitorInfo)) };
        if (monitor != IntPtr.Zero && NativeMethods.GetMonitorInfo(monitor, ref info))
        {
            return ToRect(info.WorkArea);
        }

        var screen = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea;
        return new WindowRect(screen.Left, screen.Top, screen.Width, screen.Height);
    }

    private static IReadOnlyList<WindowRect> AllWorkAreas()
    {
        var areas = new List<WindowRect>();
        foreach (var screen in System.Windows.Forms.Screen.AllScreens)
        {
            var area = screen.WorkingArea;
            areas.Add(new WindowRect(area.Left, area.Top, area.Width, area.Height));
        }

        return areas;
    }

    private static WindowRect ToRect(NativeMethods.Rect rect) =>
        new WindowRect(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top);

    // WPF event handlers run on Excel's thread: an exception escaping one would reach Excel.
    private static void Safe(Action action)
    {
        try
        {
            action();
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("TraceWindowError", ex.ToString());
        }
    }

    private static class NativeMethods
    {
        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")]
        public static extern IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")]
        public static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint flags);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(IntPtr hWnd, out Rect rect);

        [DllImport("user32.dll")]
        public static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint flags);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetMonitorInfo(IntPtr monitor, ref MonitorInfo info);

        [DllImport("user32.dll")]
        public static extern uint GetDpiForWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetFocus();

        [DllImport("user32.dll")]
        public static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindow(IntPtr hWnd);

        /// <summary>RECT.</summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct Rect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        /// <summary>MONITORINFO.</summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct MonitorInfo
        {
            public int Size;
            public Rect Monitor;
            public Rect WorkArea;
            public uint Flags;
        }
    }
}

/// <summary>A row of the Trace In list: what <see cref="TraceWindow"/> binds to. Public, because WPF binds to public types.</summary>
public sealed class TraceRow
{
    /// <summary>Creates a row.</summary>
    public TraceRow(int depth, string glyph, string icon, string kindName, string label, bool dim, string badge, string argument, string value)
    {
        IndentWidth = depth * 16.0;
        Glyph = glyph;
        Icon = icon;
        KindName = kindName;
        Label = label;
        LabelStyle = dim ? FontStyles.Italic : FontStyles.Normal;
        Badge = badge;
        Argument = argument;
        Value = value;
    }

    /// <summary>Left indent of the Precedents column.</summary>
    public double IndentWidth { get; }

    /// <summary>The expander: ▸ (collapsed), ▾ (expanded) or empty.</summary>
    public string Glyph { get; }

    /// <summary>The kind's symbol.</summary>
    public string Icon { get; }

    /// <summary>The kind's name (the icon's tooltip).</summary>
    public string KindName { get; }

    /// <summary>The Precedents column.</summary>
    public string Label { get; }

    /// <summary>Italic for marker rows ("more cells", "truncated") and errors.</summary>
    public FontStyle LabelStyle { get; }

    /// <summary>Badges after the label: ↻ for a cycle, the hidden note.</summary>
    public string Badge { get; }

    /// <summary>The Argument column (empty in classic mode).</summary>
    public string Argument { get; }

    /// <summary>The Value column.</summary>
    public string Value { get; }
}
