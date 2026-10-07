using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelModelingToolkit.Core.Keys;
using ExcelModelingToolkit.Core.Trace;

namespace ExcelModelingToolkit.AddIn;

/// <summary>
/// The Trace In keyboard (variant C of spike K4): a thread keyboard hook (<c>WH_KEYBOARD</c>) on Excel's main
/// thread, installed only while the Trace In window is open. For each press it asks <see cref="TraceKeys"/> whether
/// the key is a Trace In command and whether Excel can hand it over now; if so it swallows the key (with its
/// key-up, and its auto-repeats unless the command repeats) and runs the command after the hook returns; every
/// other key passes untouched.
/// </summary>
/// <remarks>
/// <para>
/// <b>Edit mode.</b> Excel's formula-edit state (<see cref="ExcelDnaUtil.IsInFormulaEditMode"/>: Ready, or Enter,
/// Edit or Point) is read at every candidate press, the check <see cref="UndoKeyHook"/> already relies on. So F2
/// passes through and puts Excel in Edit or Point mode with the window open, and from then on the arrows, Enter and
/// Esc go to Excel until the edit ends, however it was started (F2, typing, a double-click, the formula bar) or
/// ended. Keys also pass while the focus is not on a worksheet grid or this window (a dialog, the Name Box), and
/// while the mouse is captured, a menu is open or a window is being moved or sized (no Excel call is made then).
/// </para>
/// <para>
/// <b>Coexistence with <see cref="UndoKeyHook"/>.</b> Both are thread hooks on the same thread. Windows calls the
/// newest first, so this one sees each key first and hands every key it does not take to the next hook
/// (<c>CallNextHookEx</c>), which is how Ctrl+Z and Ctrl+Y still reach the undo hook while the window is open. Trace
/// In never takes those keys.
/// </para>
/// <para>
/// <b>Running commands.</b> As in <see cref="UndoKeyHook"/>, the decision is made in the hook and the work is posted to
/// a hidden window on the main thread. Unlike the undo hook, the work then runs there, from Excel's message loop, not
/// as a queued macro: Trace In writes nothing, and spike K2/K2b showed that macro context, not the message loop, is
/// what loses Excel's undo history (see <see cref="TraceSession"/>).
/// </para>
/// </remarks>
internal static class TraceKeyHook
{
    private const int WhKeyboard = 2;
    private const int HcAction = 0;
    private const int VkShift = 0x10;
    private const int VkControl = 0x11;
    private const int VkMenu = 0x12; // Alt

    private const int GuiInMoveSize = 0x2;
    private const int GuiInMenuMode = 0x4;
    private const int GuiPopupMenuMode = 0x10;

    /// <summary>GetAncestor: the top-level window.</summary>
    private const uint GaRoot = 2;

    /// <summary>Window class of a worksheet grid window, which has the keyboard focus while cells are selected.</summary>
    private const string GridWindowClass = "EXCEL7";

    // The hook holds only a native pointer to the delegate: keep it referenced while installed.
    private static NativeMethods.HookProc? _proc;
    private static IntPtr _hook;
    private static Control? _poster;
    private static uint _thread;
    private static bool _unloadSubscribed;
    private static bool _inProc;

    // The key whose press we took, and its command: its key-up (and, unless the command repeats, its auto-repeats)
    // are swallowed too. Zero when none.
    private static int _swallowedKey;
    private static TraceKeyCommand _swallowedCommand;

    /// <summary>True while the hook is installed.</summary>
    public static bool IsInstalled => _hook != IntPtr.Zero;

    /// <summary>
    /// Installs the hook on the calling thread, which must be Excel's main thread (Trace In opens in a macro).
    /// Returns true on success. Never throws.
    /// </summary>
    public static bool Install()
    {
        if (IsInstalled)
        {
            return true;
        }

        try
        {
            if (!_unloadSubscribed)
            {
                // Excel can unload the add-in's AppDomain without AutoClose (e.g. when it exits): unhook then too.
                AppDomain.CurrentDomain.DomainUnload += (sender, e) => Uninstall();
                _unloadSubscribed = true;
            }

            // As in UndoKeyHook: no WinForms SynchronizationContext on Excel's main thread.
            WindowsFormsSynchronizationContext.AutoInstall = false;
            var poster = new Control();
            poster.CreateControl();
            _ = poster.Handle;
            _poster = poster;

            _thread = NativeMethods.GetCurrentThreadId();
            _proc = Proc;
            _hook = NativeMethods.SetWindowsHookEx(WhKeyboard, _proc, IntPtr.Zero, _thread);
            var error = Marshal.GetLastWin32Error();
            DiagnosticsLog.Write(
                "TraceHook",
                _hook != IntPtr.Zero ? "installed" : "failed",
                "thread=" + _thread.ToString(CultureInfo.InvariantCulture),
                "win32Error=" + error.ToString(CultureInfo.InvariantCulture));
            if (_hook == IntPtr.Zero)
            {
                Uninstall();
                return false;
            }

            return true;
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("TraceHook", "failed", ex.ToString());
            Uninstall();
            return false;
        }
    }

    /// <summary>Removes the hook and its hidden window (the window closed, AutoClose, or the AppDomain unloading). Never throws.</summary>
    public static void Uninstall()
    {
        try
        {
            if (_hook != IntPtr.Zero)
            {
                NativeMethods.UnhookWindowsHookEx(_hook);
                DiagnosticsLog.Write("TraceHook", "uninstalled");
            }

            _poster?.Dispose();
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("TraceHook", "uninstall failed", ex.Message);
        }

        _hook = IntPtr.Zero;
        _proc = null;
        _poster = null;
        _swallowedKey = 0;
    }

    private static IntPtr Proc(int code, IntPtr wParam, IntPtr lParam)
    {
        if (_inProc)
        {
            return NativeMethods.CallNextHookEx(_hook, code, wParam, lParam);
        }

        _inProc = true;
        try
        {
            if (code == HcAction && _poster is not null && TraceSession.Current is TraceSession session)
            {
                var key = wParam.ToInt32();
                var flags = lParam.ToInt64();
                var keyUp = (flags & 0x80000000L) != 0;
                var repeat = (flags & 0x40000000L) != 0;

                if (key == _swallowedKey && keyUp)
                {
                    _swallowedKey = 0;
                    return new IntPtr(1);
                }

                // Never drive a window that is gone or hidden (closed without its Closed event, say): pass the key
                // and close the trace.
                var window = session.WindowHandle;
                if (window == IntPtr.Zero || !NativeMethods.IsWindow(window) || !NativeMethods.IsWindowVisible(window))
                {
                    _swallowedKey = 0;
                    Post(() => TraceSession.Current?.Abort("window gone"));
                    return NativeMethods.CallNextHookEx(_hook, code, wParam, lParam);
                }

                if (!keyUp && repeat && key == _swallowedKey)
                {
                    if (!TraceKeys.Repeats(_swallowedCommand))
                    {
                        return new IntPtr(1);
                    }

                    // The modifiers may have changed while the key was held (Ctrl released during Ctrl+Down).
                    var held = TraceKeys.Command(key, Modifiers());
                    if (held != TraceKeyCommand.None && TraceKeys.Takes(held, Context()) && Dispatch(held))
                    {
                        _swallowedCommand = held;
                        return new IntPtr(1);
                    }

                    _swallowedKey = 0;
                }
                else if (!keyUp)
                {
                    // A fresh press ends any swallowed key whose key-up we never saw (focus moved, say).
                    _swallowedKey = 0;
                    var command = TraceKeys.Command(key, Modifiers());
                    if (command != TraceKeyCommand.None)
                    {
                        var context = Context();
                        if (TraceKeys.Takes(command, context) && Dispatch(command))
                        {
                            _swallowedKey = key;
                            _swallowedCommand = command;
                            return new IntPtr(1);
                        }

                        Log("TraceKey", command.ToString(), "passed to Excel: " + context);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Log("TraceHook", "error", ex.ToString());
        }
        finally
        {
            _inProc = false;
        }

        return NativeMethods.CallNextHookEx(_hook, code, wParam, lParam);
    }

    /// <summary>
    /// Runs the command after the hook returns, from Excel's message loop and not as a macro (see
    /// <see cref="TraceSession"/>: macro context would cost Excel's undo history). If the keyboard is on another
    /// workbook's grid than the one that owns the window (the user switched workbooks), the window is re-owned to it
    /// first, so the keys never drive a window hidden behind another.
    /// </summary>
    private static bool Dispatch(TraceKeyCommand command)
    {
        var pressed = Stopwatch.GetTimestamp();
        var focusRoot = NativeMethods.GetAncestor(NativeMethods.GetFocus(), GaRoot);
        if (!Post(() => TraceSession.Current?.Follow(focusRoot)))
        {
            return false;
        }

        if (TraceKeys.IsWindowCommand(command))
        {
            return Post(() => TraceSession.Current?.ApplyWindowCommand(command));
        }

        var isMove = command == TraceKeyCommand.Up || command == TraceKeyCommand.Down ||
            command == TraceKeyCommand.Left || command == TraceKeyCommand.Right;
        if (!Post(() => TraceSession.Current?.Handle(command, pressed, "key")))
        {
            return false;
        }

        if (isMove)
        {
            TraceSession.Current?.MoveQueued();
        }

        return true;
    }

    /// <summary>Whether Excel can hand keys to Trace In now (see the remarks). Never throws.</summary>
    private static TraceKeyContext Context()
    {
        try
        {
            // Checked before any other call: Excel is in a modal loop of its own.
            var modal = NativeMethods.GetCapture() != IntPtr.Zero;
            if (!modal)
            {
                var info = new NativeMethods.GuiThreadInfo { Size = Marshal.SizeOf(typeof(NativeMethods.GuiThreadInfo)) };
                modal = NativeMethods.GetGUIThreadInfo(_thread, ref info) &&
                    (info.Flags & (GuiInMenuMode | GuiPopupMenuMode | GuiInMoveSize)) != 0;
            }

            if (modal)
            {
                return TraceKeys.Context(modalLoop: true, editing: false, focusOnGridOrTraceWindow: false);
            }

            return TraceKeys.Context(false, ExcelDnaUtil.IsInFormulaEditMode(), FocusIsOnGridOrTraceWindow());
        }
        catch (Exception)
        {
            return TraceKeyContext.Elsewhere;
        }
    }

    private static bool FocusIsOnGridOrTraceWindow()
    {
        var focus = NativeMethods.GetFocus();
        if (focus == IntPtr.Zero)
        {
            return false;
        }

        var window = TraceSession.Current?.WindowHandle ?? IntPtr.Zero;
        if (window != IntPtr.Zero && (focus == window || NativeMethods.IsChild(window, focus)))
        {
            return true;
        }

        var name = new StringBuilder(64);
        return NativeMethods.GetClassName(focus, name, name.Capacity) > 0 &&
            string.Equals(name.ToString(), GridWindowClass, StringComparison.Ordinal);
    }

    private static KeyModifiers Modifiers() =>
        (IsDown(VkControl) ? KeyModifiers.Ctrl : KeyModifiers.None) |
        (IsDown(VkMenu) ? KeyModifiers.Alt : KeyModifiers.None) |
        (IsDown(VkShift) ? KeyModifiers.Shift : KeyModifiers.None);

    /// <summary>Writes a diagnostics log line after the hook returns; nothing is posted when logging is off.</summary>
    private static void Log(string eventName, params string[] fields)
    {
        if (DiagnosticsLog.Enabled)
        {
            Post(() => DiagnosticsLog.Write(eventName, fields));
        }
    }

    private static bool Post(Action work)
    {
        try
        {
            if (_poster is null)
            {
                return false;
            }

            _poster.BeginInvoke(new Action(() =>
            {
                // Runs in WinForms' message dispatch: an exception escaping here would show its error dialog in Excel.
                try
                {
                    work();
                }
                catch (Exception ex)
                {
                    DiagnosticsLog.Write("TraceHook", "posted work failed", ex.ToString());
                }
            }));
            return true;
        }
        catch (Exception)
        {
            // The window is gone (the hook is being removed); nothing can run.
            return false;
        }
    }

    private static bool IsDown(int virtualKey) => (NativeMethods.GetKeyState(virtualKey) & 0x8000) != 0;

    private static class NativeMethods
    {
        public delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        public static extern short GetKeyState(int nVirtKey);

        [DllImport("user32.dll")]
        public static extern IntPtr GetFocus();

        [DllImport("user32.dll")]
        public static extern IntPtr GetCapture();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsChild(IntPtr hWndParent, IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetGUIThreadInfo(uint idThread, ref GuiThreadInfo pgui);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        /// <summary>GUITHREADINFO.</summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct GuiThreadInfo
        {
            public int Size;
            public int Flags;
            public IntPtr Active;
            public IntPtr Focus;
            public IntPtr Capture;
            public IntPtr MenuOwner;
            public IntPtr MoveSize;
            public IntPtr Caret;
            public int CaretLeft;
            public int CaretTop;
            public int CaretRight;
            public int CaretBottom;
        }
    }
}
