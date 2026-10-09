using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;
using Modelwright.Core.Undo;

namespace Modelwright.AddIn;

/// <summary>
/// Ctrl+Z and Ctrl+Y for our formatting undo (docs/PLAN.md section 4.4): a thread keyboard hook
/// (<c>WH_KEYBOARD</c>) on Excel's main thread, the design proven in spikes K2b and K4. For each press it asks
/// <see cref="UndoManager.Decide"/> whether the key is ours or Excel's. If it is Excel's, the key passes through
/// untouched (after dropping our redo stack if Excel's history shows it is stale); if ours, it is swallowed (with
/// its repeats and key-up) and <see cref="UndoCommand"/> runs after the hook returns.
/// </summary>
/// <remarks>
/// <para>
/// Only the keyboard is intercepted. Excel's own Undo and Redo buttons (Quick Access Toolbar, and the Undo list)
/// are not in v1: they always act on Excel's native history. The ribbon's Undo formatting and Redo formatting
/// buttons act on ours.
/// </para>
/// <para>
/// Holding Ctrl+Z (or Ctrl+Y) down undoes (or redoes) one step when the key is ours: the auto-repeats are
/// swallowed with the press. Release and press again for the next step.
/// </para>
/// <para>
/// Excel's state is queried inside the hook (the decision must be made before the key is passed on), with
/// <c>CommandBars.GetEnabledMso("Undo")</c> and <c>("Redo")</c>. The state is unknown, so the key passes, while the
/// mouse is captured, a menu is open or a window is being moved or sized (no COM call is made then), while a cell
/// is being edited, when the keyboard focus is not on a worksheet grid (a dialog, the formula bar, the VBA editor,
/// a task pane), or if the query fails. A press that arrives while the hook is already running
/// (the query can pump messages) passes untouched. All other work, including the log lines, is posted to a hidden
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

    /// <summary>GUITHREADINFO flags: a window is being moved or sized, or a menu (menu bar or popup) is active.</summary>
    private const int GuiInMoveSize = 0x2;
    private const int GuiInMenuMode = 0x4;
    private const int GuiPopupMenuMode = 0x10;

    /// <summary>Window class of a worksheet grid window, which has the keyboard focus while cells are selected.</summary>
    private const string GridWindowClass = "EXCEL7";

    // The hook holds only a native pointer to the delegate: it stays referenced for the life of the AppDomain (a
    // call can still be on its way in while the hook is being removed).
    private static readonly NativeMethods.HookProc HookCallback = Proc;

    private static IntPtr _hook;
    private static Control? _poster;
    private static uint _thread;
    private static bool _unloadSubscribed;

    // True while Proc runs: a key that arrives meanwhile (a COM call can pump messages) passes untouched.
    private static bool _inProc;

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
            if (!_unloadSubscribed)
            {
                // Excel can unload the add-in's AppDomain without AutoClose (e.g. when it exits): unhook then too.
                AppDomain.CurrentDomain.DomainUnload += (sender, e) => Uninstall();
                _unloadSubscribed = true;
            }

            // Creating a Control would otherwise install a WinForms SynchronizationContext on Excel's main thread,
            // changing how every awaiting add-in in this AppDomain resumes.
            WindowsFormsSynchronizationContext.AutoInstall = false;

            // The hidden window must be created on the main thread: BeginInvoke then runs work there.
            var poster = new Control();
            poster.CreateControl();
            _ = poster.Handle;
            _poster = poster;

            _thread = NativeMethods.GetCurrentThreadId();
            _hook = NativeMethods.SetWindowsHookEx(WhKeyboard, HookCallback, IntPtr.Zero, _thread);
            var error = Marshal.GetLastWin32Error();
            DiagnosticsLog.Write(
                "UndoHook",
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
            DiagnosticsLog.Write("UndoHook", "failed", ex.ToString());
            Uninstall();
            return false;
        }
    }

    /// <summary>Removes the hook and the hidden window (AutoClose, or the AppDomain unloading). Never throws.</summary>
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

                if (!keyUp && !repeat)
                {
                    // A fresh press ends any swallowed key whose key-up we never saw (focus moved, say).
                    _swallowedKey = 0;
                    if ((key == VkZ || key == VkY) &&
                        IsDown(VkControl) && !IsDown(VkMenu) && !IsDown(VkShift) &&
                        Handle(key == VkZ ? UndoKey.Undo : UndoKey.Redo))
                    {
                        _swallowedKey = key;
                        return new IntPtr(1);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Log("UndoHook", "error", ex.ToString());
        }
        finally
        {
            _inProc = false;
        }

        return NativeMethods.CallNextHookEx(_hook, code, wParam, lParam);
    }

    /// <summary>
    /// Decides one press and logs it. Drops our redo stack if Excel's history shows it is stale. True if the key is
    /// ours: the restore has been posted and the key must be swallowed.
    /// </summary>
    private static bool Handle(UndoKey key)
    {
        var undo = Session.Undo.UndoCount;
        var redo = Session.Undo.RedoCount;
        bool? nativeUndo = null;
        bool? nativeRedo = null;
        bool? nativeRepeat = null;
        string context;
        if (key == UndoKey.Undo ? undo + redo > 0 : redo > 0)
        {
            (nativeUndo, nativeRedo, nativeRepeat) = NativeState(out context);
        }
        else
        {
            context = "context=stacks empty, not queried";
        }

        var decision = UndoManager.Decide(key, undo, redo, nativeUndo, nativeRedo);
        var dropped = decision == UndoDecision.PassToExcelAndClearRedo ? Session.Undo.ClearRedo() : 0;
        Log(
            "UndoKey",
            key == UndoKey.Undo ? "Z" : "Y",
            "undo=" + undo.ToString(CultureInfo.InvariantCulture),
            "redo=" + redo.ToString(CultureInfo.InvariantCulture),
            "nativeUndo=" + NativeUndoState.Describe(nativeUndo),
            "nativeRedo=" + NativeUndoState.Describe(nativeRedo),
            "nativeRepeat=" + NativeUndoState.Describe(nativeRepeat),
            "decision=" + decision,
            "droppedRedo=" + dropped.ToString(CultureInfo.InvariantCulture),
            context);
        if (decision != UndoDecision.HandleOurs)
        {
            return false;
        }

        // Leave the hook first, then run as a macro: the restore needs the C API for the status bar. Macro context
        // costs nothing here, because the restore's COM writes clear Excel's history anyway (spike K2c).
        return Post(() => ExcelAsyncUtil.QueueAsMacro(() => UndoCommand.Run(key, "key")));
    }

    /// <summary>
    /// Excel's own Undo and Redo states (and Repeat's, when logging), each null when unknown: the mouse is captured,
    /// a menu is open or a window is being moved or sized, a cell is being edited, the focus is not on a worksheet
    /// grid, or the query failed. <paramref name="context"/> says which, for the log.
    /// </summary>
    private static (bool? Undo, bool? Redo, bool? Repeat) NativeState(out string context)
    {
        try
        {
            // Checked before any COM call: Excel is in a modal loop of its own.
            if (NativeMethods.GetCapture() != IntPtr.Zero)
            {
                context = "context=mouse captured";
                return (null, null, null);
            }

            var info = new NativeMethods.GuiThreadInfo { Size = Marshal.SizeOf(typeof(NativeMethods.GuiThreadInfo)) };
            if (NativeMethods.GetGUIThreadInfo(_thread, ref info) &&
                (info.Flags & (GuiInMenuMode | GuiPopupMenuMode | GuiInMoveSize)) != 0)
            {
                context = "context=menu or move/size, flags=0x" + info.Flags.ToString("x", CultureInfo.InvariantCulture);
                return (null, null, null);
            }

            if (ExcelDnaUtil.IsInFormulaEditMode())
            {
                context = "context=editing a cell";
                return (null, null, null);
            }

            var focus = FocusClass();
            if (!string.Equals(focus, GridWindowClass, StringComparison.Ordinal))
            {
                context = "context=focus " + focus;
                return (null, null, null);
            }

            context = "context=grid";
            return NativeUndoState.Query(withRepeat: DiagnosticsLog.Enabled);
        }
        catch (Exception ex)
        {
            context = "context=query failed: " + ex.Message;
            return (null, null, null);
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

    /// <summary>Writes a diagnostics log line after the hook returns; nothing is posted when logging is off.</summary>
    private static void Log(string eventName, params string[] fields)
    {
        if (DiagnosticsLog.Enabled)
        {
            Post(() => DiagnosticsLog.Write(eventName, fields));
        }
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

        [DllImport("user32.dll")]
        public static extern IntPtr GetCapture();

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
