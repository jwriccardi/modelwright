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
/// key-up, and its auto-repeats unless the command repeats) and queues the command on the session; every other key
/// passes untouched.
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
/// <b>Running commands.</b> As in <see cref="UndoKeyHook"/>, the decision is made in the hook and the work runs after
/// it returns, from a hidden window on the main thread: the command joins the session's queue
/// (<see cref="TraceSession.Enqueue"/>), with clicks and buttons, and runs there in order, from Excel's message loop,
/// not as a queued macro: Trace In writes nothing, and spike K2/K2b showed that macro context, not the message loop,
/// is what loses Excel's undo history (see <see cref="TraceSession"/>).
/// </para>
/// <para>
/// <b>Closing.</b> When Enter or Esc closes the window, the hook stays until that key is released (at most
/// <see cref="LingerMilliseconds"/>), swallowing only that key's auto-repeats and key-up, so a held Enter does not go
/// on to move Excel's selection; any other key ends that at once and passes.
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

    /// <summary>Window class of an Excel workbook's top-level window (Excel 2013 and later: one per workbook window).</summary>
    private const string WorkbookWindowClass = "XLMAIN";

    /// <summary>How long, at most, the hook outlives the window to swallow the key that closed it.</summary>
    private const int LingerMilliseconds = 1000;

    // The hook holds only a native pointer to the delegate: it stays referenced for the life of the AppDomain (a
    // call can still be on its way in while the hook is being removed).
    private static readonly NativeMethods.HookProc HookCallback = Proc;

    private static IntPtr _hook;
    private static Control? _poster;
    private static uint _thread;
    private static bool _unloadSubscribed;
    private static bool _inProc;

    // The key whose press we took, and its command: its key-up (and, unless the command repeats, its auto-repeats)
    // are swallowed too. Zero when none.
    private static int _swallowedKey;
    private static TraceKeyCommand _swallowedCommand;

    // After the window closed on Enter or Esc: the hook stays (until _lingerUntil, a Stopwatch timestamp) to swallow
    // that key's repeats and key-up; the timer removes it.
    private static bool _lingering;
    private static long _lingerUntil;
    private static Timer? _lingerTimer;

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
            // Still there after a close (swallowing its key): a new trace keeps it.
            StopLingering();
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
            _hook = NativeMethods.SetWindowsHookEx(WhKeyboard, HookCallback, IntPtr.Zero, _thread);
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

    /// <summary>
    /// The window has closed: removes the hook, or, if it closed on Enter or Esc and that key is still down, once the
    /// key is released (at most <see cref="LingerMilliseconds"/> later). Never throws.
    /// </summary>
    public static void Release()
    {
        if (!IsInstalled || _swallowedKey == 0 || TraceKeys.CloseMode(_swallowedCommand) is null)
        {
            Uninstall();
            return;
        }

        try
        {
            StopLingering();
            _lingering = true;
            _lingerUntil = Stopwatch.GetTimestamp() + (LingerMilliseconds * Stopwatch.Frequency / 1000);
            var timer = new Timer { Interval = LingerMilliseconds + 50 };
            timer.Tick += (sender, e) => EndLingering();
            timer.Start();
            _lingerTimer = timer;
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("TraceHook", "linger failed", ex.Message);
            Uninstall();
        }
    }

    /// <summary>Removes the hook and its hidden window (the window closed, AutoClose, or the AppDomain unloading). Never throws.</summary>
    public static void Uninstall()
    {
        try
        {
            StopLingering();
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
        _poster = null;
        _swallowedKey = 0;
    }

    /// <summary>
    /// Runs <paramref name="work"/> on the main thread after the current handler returns, from the hook's hidden
    /// window (outside macro context). False if the hook is not installed. An exception from the work is logged.
    /// Never throws.
    /// </summary>
    internal static bool RunLater(Action work)
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

    /// <summary>
    /// True if <paramref name="window"/> is a workbook's top-level window (class XLMAIN) on Excel's main thread: the
    /// only windows the Trace In window is ever re-owned to (never another process's window or a dialog). Never throws.
    /// </summary>
    internal static bool IsExcelWorkbookWindow(IntPtr window)
    {
        try
        {
            if (window == IntPtr.Zero || !NativeMethods.IsWindow(window) || NativeMethods.GetAncestor(window, GaRoot) != window)
            {
                return false;
            }

            var thread = _thread != 0 ? _thread : NativeMethods.GetCurrentThreadId();
            if (NativeMethods.GetWindowThreadProcessId(window, out _) != thread)
            {
                return false;
            }

            var name = new StringBuilder(16);
            return NativeMethods.GetClassName(window, name, name.Capacity) > 0 &&
                string.Equals(name.ToString(), WorkbookWindowClass, StringComparison.Ordinal);
        }
        catch (Exception)
        {
            return false;
        }
    }

    private static void StopLingering()
    {
        _lingering = false;
        if (_lingerTimer is not null)
        {
            _lingerTimer.Stop();
            _lingerTimer.Dispose();
            _lingerTimer = null;
        }
    }

    // The timer: the key that closed the window was released, or not within the limit. Removes the hook unless a new
    // trace has opened since.
    private static void EndLingering()
    {
        try
        {
            StopLingering();
            if (TraceSession.Current is null)
            {
                Uninstall();
            }
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("TraceHook", "linger end failed", ex.Message);
        }
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
            if (code == HcAction && _lingering)
            {
                var key = wParam.ToInt32();
                var flags = lParam.ToInt64();
                var keyUp = (flags & 0x80000000L) != 0;
                var repeat = (flags & 0x40000000L) != 0;
                if (key == _swallowedKey && (keyUp || repeat) && Stopwatch.GetTimestamp() < _lingerUntil)
                {
                    if (keyUp)
                    {
                        _lingering = false;
                        _swallowedKey = 0;
                    }

                    return new IntPtr(1);
                }

                // Any other key, or too late: stop swallowing and pass it (the timer removes the hook).
                _lingering = false;
                _swallowedKey = 0;
            }

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

                // Never drive a window that is gone (closed without its Closed event, say): pass the key and close the
                // trace.
                var window = session.WindowHandle;
                if (window == IntPtr.Zero || !NativeMethods.IsWindow(window))
                {
                    _swallowedKey = 0;
                    RunLater(() => TraceSession.Current?.Abort("window gone"));
                    return NativeMethods.CallNextHookEx(_hook, code, wParam, lParam);
                }

                // Hidden because the workbook window that owns it is minimized: a command key is taken only if the
                // keyboard is on another workbook window it can be re-owned to (Dispatch does that first).
                var hiddenAndStuck = !NativeMethods.IsWindowVisible(window) && !CanFollowFocus(session);

                if (!keyUp && repeat && key == _swallowedKey)
                {
                    if (!TraceKeys.Repeats(_swallowedCommand))
                    {
                        return new IntPtr(1);
                    }

                    // The modifiers may have changed while the key was held (Ctrl released during Ctrl+Down).
                    var held = TraceKeys.Command(key, Modifiers());
                    if (!hiddenAndStuck && held != TraceKeyCommand.None && TraceKeys.Takes(held, Context()) && Dispatch(session, held))
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
                        if (hiddenAndStuck)
                        {
                            Log("TraceKey", command.ToString(), "passed to Excel: the window is hidden (its workbook window is minimized)");
                            return NativeMethods.CallNextHookEx(_hook, code, wParam, lParam);
                        }

                        var context = Context();
                        if (TraceKeys.Takes(command, context) && Dispatch(session, command))
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
    /// Queues the command on the session, to run after the hook returns, from Excel's message loop and not as a macro
    /// (see <see cref="TraceSession"/>: macro context would cost Excel's undo history). If the keyboard is on another
    /// workbook's window than the one that owns the Trace In window (the user switched workbooks), the window is
    /// re-owned to it first, so the keys never drive a window hidden behind another.
    /// </summary>
    private static bool Dispatch(TraceSession session, TraceKeyCommand command)
    {
        var pressed = Stopwatch.GetTimestamp();
        var focusRoot = FocusRoot();
        if (focusRoot != IntPtr.Zero && focusRoot != session.OwnerHandle && !session.Enqueue(() => session.Follow(focusRoot)))
        {
            return false;
        }

        if (TraceKeys.IsWindowCommand(command))
        {
            return session.Enqueue(() => session.ApplyWindowCommand(command));
        }

        if (!session.Enqueue(() => session.Handle(command, pressed, "key")))
        {
            return false;
        }

        if (command == TraceKeyCommand.Up || command == TraceKeyCommand.Down ||
            command == TraceKeyCommand.Left || command == TraceKeyCommand.Right)
        {
            session.MoveQueued();
        }

        return true;
    }

    // The workbook window that has the keyboard focus, or zero if the focus is not in one (see IsExcelWorkbookWindow).
    private static IntPtr FocusRoot()
    {
        var root = NativeMethods.GetAncestor(NativeMethods.GetFocus(), GaRoot);
        return IsExcelWorkbookWindow(root) ? root : IntPtr.Zero;
    }

    // True if the keyboard is on a workbook window other than the one that owns the (hidden) Trace In window.
    private static bool CanFollowFocus(TraceSession session)
    {
        var root = FocusRoot();
        return root != IntPtr.Zero && root != session.OwnerHandle;
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
            RunLater(() => DiagnosticsLog.Write(eventName, fields));
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
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

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
