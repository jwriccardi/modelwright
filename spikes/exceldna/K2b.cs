using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace EmtSpike
{
    /// <summary>
    /// K2b: run built-in formatting commands OUTSIDE macro context, to test whether Excel then records them
    /// like a user action (full undo history kept even when the workbook has volatile formulas).
    /// A thread-local keyboard hook catches two spike-only keys and posts the work to a hidden window on
    /// Excel's main thread, so it runs from Excel's ordinary message loop rather than from an OnKey macro.
    /// </summary>
    internal static class K2b
    {
        private const int WH_KEYBOARD = 2, HC_ACTION = 0;
        private const int VK_SHIFT = 0x10, VK_CONTROL = 0x11, VK_MENU = 0x12, VK_X = 0x58, VK_Z = 0x5A;

        private static Native.HookProc _proc; // keep referenced: the hook holds only a native pointer
        private static IntPtr _hook;
        private static Control _poster;

        public static void Install()
        {
            try
            {
                _poster = new Control();
                _poster.CreateControl();
                IntPtr _ = _poster.Handle; // force the window to exist on this (Excel's main) thread
                _proc = Proc;
                _hook = Native.SetWindowsHookEx(WH_KEYBOARD, _proc, IntPtr.Zero, Native.GetCurrentThreadId());
                Log.Write("k2bHook", new { action = "install", ok = _hook != IntPtr.Zero, win32Error = Marshal.GetLastWin32Error() });
            }
            catch (Exception ex) { Log.Error("K2b.Install", ex); }
        }

        public static void Uninstall()
        {
            try
            {
                if (_hook != IntPtr.Zero) Native.UnhookWindowsHookEx(_hook);
                _hook = IntPtr.Zero;
                _poster?.Dispose();
                _poster = null;
                Log.Write("k2bHook", new { action = "uninstall" });
            }
            catch (Exception ex) { Log.Error("K2b.Uninstall", ex); }
        }

        /// <summary>Ribbon path: called straight from the ribbon callback, no QueueAsMacro.</summary>
        public static void RibbonBold() =>
            Commands.Run("K2.10 ribbon-direct ExecuteMso Bold", "ribbon-direct", () => K2.MsoBold("K2.10"));

        private static IntPtr Proc(int code, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (code == HC_ACTION && _poster != null)
                {
                    int vk = wParam.ToInt32();
                    if ((vk == VK_Z || vk == VK_X) && Down(VK_CONTROL) && Down(VK_MENU) && Down(VK_SHIFT))
                    {
                        long flags = lParam.ToInt64();
                        bool keyUp = (flags & 0x80000000L) != 0;
                        bool repeat = (flags & 0x40000000L) != 0;
                        if (!keyUp && !repeat)
                        {
                            Action work = vk == VK_Z
                                ? (Action)(() => Commands.Run("K2.8 hook ExecuteMso Bold (no macro)", "hook:^%+z", () => K2.MsoBold("K2.8")))
                                : () => Commands.Run("K2.9 hook Copy+PasteFormatting (no macro)", "hook:^%+x", () => K2.CopyPasteMso("K2.9"));
                            _poster.BeginInvoke(work); // runs after the hook returns, from Excel's message loop
                        }
                        return new IntPtr(1); // swallow down, repeat and up
                    }
                }
            }
            catch (Exception ex) { Log.Error("K2b.Proc", ex); }
            return Native.CallNextHookEx(_hook, code, wParam, lParam);
        }

        private static bool Down(int vk) => (Native.GetKeyState(vk) & 0x8000) != 0;
    }
}
