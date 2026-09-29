using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelModelingToolkit.Core.Undo;

namespace ExcelModelingToolkit.AddIn;

/// <summary>
/// Ctrl+Z and Ctrl+Y for our formatting undo (docs/PLAN.md section 4.4): a thread keyboard hook
/// (<c>WH_KEYBOARD</c>) on Excel's main thread, the design proven in spikes K2b and K4. For each press it asks
/// <see cref="UndoManager.Decide"/> whether the key is ours or Excel's. If it is Excel's, the key passes through
/// untouched; if ours, it is swallowed (with its repeats and key-up) and <see cref="UndoCommand"/> runs after the
/// hook returns.
/// </summary>
/// <remarks>
/// <para>
/// Only the keyboard is intercepted. Excel's own Undo and Redo buttons (Quick Access Toolbar, and the Undo list)
/// are not in v1: they always act on Excel's native history. The ribbon's Undo formatting and Redo formatting
/// buttons act on ours.
/// </para>
/// <para>
/// Excel's state is queried inside the hook (the decision must be made before the key is passed on), with
/// <c>CommandBars.GetEnabledMso("Undo")</c> or <c>("Redo")</c>. The state is unknown, so the key passes, while a
/// cell is being edited, when the keyboard focus is not on a worksheet grid (a dialog, the formula bar, the VBA
/// editor, a task pane), or if the query fails. All other work, including the log lines, is posted to a hidden
/// window on the main thread so the hook returns at once.
/// </para>
/// </remarks>
internal static class UndoKeyHook
{
    private const int WhKeyboard = 2;
    private const int HcAction = 0;
    private const int VkShift = 0x10;
    private const int VkControl = 0x11;
    private const int VkMenu = 0x12; // Alt
    private const int VkY = 0x59;
    private const int VkZ = 0x5A;

    /// <summary>Window class of a worksheet grid window, which has the keyboard focus while cells are selected.</summary>
    private const string GridWindowClass = "EXCEL7";

    // The hook holds only a native pointer to the delegate: keep it referenced while installed.
    private static NativeMethods.HookProc? _proc;
    private static IntPtr _hook;
    private static Control? _poster;

    // The key whose press we swallowed: its auto-repeats and key-up are swallowed too. Zero when none.
    private static int _swallowedKey;

    /// <summary>
    /// Installs the hook on the calling thread, which must be Excel's main thread (call from AutoOpen). Returns
    /// true on success. Never throws.
    /// </summary>
    public static bool Install()
    {
        Uninstall();
        try
        {
            // The hidden window must be created on the main thread: BeginInvoke then runs work there.
            var poster = new Control();
            poster.CreateControl();
            _ = poster.Handle;
            _poster = poster;

            var thread = NativeMethods.GetCurrentThreadId();
            _proc = Proc;
            _hook = NativeMethods.SetWindowsHookEx(WhKeyboard, _proc, IntPtr.Zero, thread);
            var error = Marshal.GetLastWin32Error();
            DiagnosticsLog.Write(
                "UndoHook",
                _hook != IntPtr.Zero ? "installed" : "failed",
                "thread=" + thread.ToString(CultureInfo.InvariantCulture),
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
            DiagnosticsLog.Write("UndoHook", "failed", ex.ToString());
            Uninstall();
            return false;
        }
    }

    /// <summary>Removes the hook and the hidden window (AutoClose). Never throws.</summary>
    public static void Uninstall()
    {
        try
        {
            if (_hook != IntPtr.Zero)
            {
                NativeMethods.UnhookWindowsHookEx(_hook);
                DiagnosticsLog.Write("UndoHook", "uninstalled");
            }

            _poster?.Dispose();
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("UndoHook", "uninstall failed", ex.Message);
        }

        _hook = IntPtr.Zero;
        _proc = null;
        _poster = null;
        _swallowedKey = 0;
    }

    private static IntPtr Proc(int code, IntPtr wParam, IntPtr lParam)
    {
        try
        {
            if (code == HcAction && _poster is not null)
            {
                var key = wParam.ToInt32();
                var flags = lParam.ToInt64();
                var keyUp = (flags & 0x80000000L) != 0;
                var repeat = (flags & 0x40000000L) != 0;

                if (key == _swallowedKey && (keyUp || repeat))
                {
                    if (keyUp)
                    {
                        _swallowedKey = 0;
                    }

                    return new IntPtr(1);
                }

                if ((key == VkZ || key == VkY) && !keyUp && !repeat &&
                    IsDown(VkControl) && !IsDown(VkMenu) && !IsDown(VkShift) &&
                    Handle(key == VkZ ? UndoKey.Undo : UndoKey.Redo))
                {
                    _swallowedKey = key;
                    return new IntPtr(1);
                }
            }
        }
        catch (Exception ex)
        {
            Post(() => DiagnosticsLog.Write("UndoHook", "error", ex.ToString()));
        }

        return NativeMethods.CallNextHookEx(_hook, code, wParam, lParam);
    }

    /// <summary>Decides one press and logs it. True if it is ours: the restore has been posted and the key must be swallowed.</summary>
    private static bool Handle(UndoKey key)
    {
        var count = Session.Undo.Count(key);
        bool? native = null;
        var context = "context=stack empty, not queried";
        if (count > 0)
        {
            native = NativeAvailable(key, out context);
        }

        var decision = UndoManager.Decide(key, count > 0, native);
        var fields = new[]
        {
            key == UndoKey.Undo ? "Z" : "Y",
            "stack=" + count.ToString(CultureInfo.InvariantCulture),
            "native=" + (native is bool enabled ? (enabled ? "true" : "false") : "unknown"),
            "decision=" + decision,
            context,
        };
        Post(() => DiagnosticsLog.Write("UndoKey", fields));
        if (decision != UndoDecision.HandleOurs)
        {
            return false;
        }

        // Leave the hook first, then run as a macro: the restore needs the C API for the status bar. Macro context
        // costs nothing here, because the restore's COM writes clear Excel's history anyway (spike K2c).
        return Post(() => ExcelAsyncUtil.QueueAsMacro(() => UndoCommand.Run(key, "key")));
    }

    /// <summary>
    /// Whether Excel's own Undo (or Redo) is enabled, or null when unknown: a cell is being edited, the focus is
    /// not on a worksheet grid, or the query failed. <paramref name="context"/> says which, for the log.
    /// </summary>
    private static bool? NativeAvailable(UndoKey key, out string context)
    {
        try
        {
            if (ExcelDnaUtil.IsInFormulaEditMode())
            {
                context = "context=editing a cell";
                return null;
            }

            var focus = FocusClass();
            if (!string.Equals(focus, GridWindowClass, StringComparison.Ordinal))
            {
                context = "context=focus " + focus;
                return null;
            }

            dynamic app = ExcelDnaUtil.Application;
            object enabled = app.CommandBars.GetEnabledMso(key == UndoKey.Undo ? "Undo" : "Redo");
            context = "context=grid";
            return enabled is bool value ? value : (bool?)null;
        }
        catch (Exception ex)
        {
            context = "context=query failed: " + ex.Message;
            return null;
        }
    }

    /// <summary>The window class of this thread's keyboard focus window, or <c>(none)</c>.</summary>
    private static string FocusClass()
    {
        var focus = NativeMethods.GetFocus();
        if (focus == IntPtr.Zero)
        {
            return "(none)";
        }

        var name = new StringBuilder(64);
        return NativeMethods.GetClassName(focus, name, name.Capacity) > 0 ? name.ToString() : "(unknown)";
    }

    /// <summary>Runs <paramref name="work"/> on the main thread after the hook returns. False if it could not be posted.</summary>
    private static bool Post(Action work)
    {
        try
        {
            if (_poster is null)
            {
                return false;
            }

            _poster.BeginInvoke(work);
            return true;
        }
        catch (Exception)
        {
            // The window is gone (the add-in is closing); nothing can run.
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

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
    }
}
