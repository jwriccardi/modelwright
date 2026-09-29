using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace EmtSpike
{
    /// <summary>K4: throwaway precedent extraction + a modeless navigator window (WPF, WinForms, or non-activating WPF + keyboard hook).</summary>
    internal static class Trace
    {
        public enum Ui { Wpf, WinForms, Hook }

        internal sealed class Item
        {
            public string Label, Address, Value, Error;
            public object Range;
            public override string ToString() =>
                $"{Label,-36} {Address}  = {Value}" + (Error != null ? "   [" + Error + "]" : "");
        }

        // Optional prefix ('[Book]Sheet'! | [Book]Sheet! | Sheet!) + A1 cell or range. Not a real parser.
        private static readonly Regex RefRx = new Regex(
            @"(?<![\w.$!:'\]])(?:(?:'(?<q>(?:[^']|'')+)'|(?<u>(?:\[[^\]]+\])?[A-Za-z_][\w.]*))!)?" +
            @"(?<a>\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?)(?![\w(!])");
        private static readonly Regex BookRx = new Regex(@"^(?:.*\\)?\[(?<b>[^\]]+)\](?<s>.+)$");

        private static dynamic _lastRoot;

        public static object Open(Ui ui, string key, bool reactivate)
        {
            var sw = Stopwatch.StartNew();
            dynamic app = Xl.App;
            dynamic root = app.ActiveCell;
            if (!(bool)root.HasFormula)
            {
                string where = Address((object)root);
                Xl.Status($"EMT spike: {where} has no formula - select EMT_Fixture_A.xlsx Sheet1!A1 and try again");
                return new { ui = ui.ToString(), key, refused = "no formula", where };
            }
            _lastRoot = root;
            var session = new Session(ui, root, reactivate);
            session.Build();

            var owner = new IntPtr(Convert.ToInt64(app.ActiveWindow.Hwnd));
            session.OwnerHwnd = owner;
            switch (ui)
            {
                case Ui.Wpf: TraceWindowWpf.ShowNew(session, owner, noActivate: false); break;
                case Ui.Hook: TraceWindowWpf.ShowNew(session, owner, noActivate: true); break;
                default: new TraceForm(session).ShowOwned(owner); break;
            }
            return new { ui = ui.ToString(), key, session.Formula, refs = session.Items.Count - 1, openMs = Xl.Ms(sw), ourHwnd = session.OurHwnd.ToInt64() };
        }

        public static object GotoLastAudited()
        {
            if (_lastRoot == null) return "no audited cell yet";
            Goto(_lastRoot);
            return Address(_lastRoot);
        }

        /// <summary>Activates the range's workbook window if needed, then Application.Goto. Throws on failure (e.g. hidden sheet).</summary>
        internal static void Goto(dynamic rng)
        {
            dynamic app = Xl.App;
            dynamic wb = rng.Worksheet.Parent;
            if (Convert.ToString(wb.Name) != Convert.ToString(app.ActiveWorkbook.Name))
                wb.Activate();
            app.Goto(rng);
        }

        internal static string Address(object rng) =>
            Convert.ToString(Xl.Get(rng, "Address", true, true, 1, true)); // RowAbs, ColAbs, xlA1, External

        private static Item MakeItem(string label, dynamic rng)
        {
            var item = new Item { Label = label, Range = rng };
            item.Address = Address((object)rng);
            item.Value = Convert.ToString(rng.Cells[1, 1].Value2);
            return item;
        }

        private static Item Resolve(Match m, dynamic root)
        {
            string addr = m.Groups["a"].Value, book = null, sheet = null;
            string prefix = m.Groups["q"].Success ? m.Groups["q"].Value.Replace("''", "'")
                          : m.Groups["u"].Success ? m.Groups["u"].Value : null;
            if (prefix != null)
            {
                Match bm = BookRx.Match(prefix);
                if (bm.Success) { book = bm.Groups["b"].Value; sheet = bm.Groups["s"].Value; }
                else sheet = prefix;
            }
            try
            {
                dynamic app = Xl.App;
                dynamic ws = book != null ? app.Workbooks[book].Worksheets[sheet]
                           : sheet != null ? root.Worksheet.Parent.Worksheets[sheet]
                           : root.Worksheet;
                return MakeItem(m.Value, ws.Range[addr]);
            }
            catch (Exception ex)
            {
                return new Item { Label = m.Value, Address = $"book={book} sheet={sheet} addr={addr}", Error = ex.Message };
            }
        }

        /// <summary>State shared by all window variants. All members are used on Excel's main thread.</summary>
        internal sealed class Session
        {
            public readonly Ui Ui;
            public readonly dynamic Root;
            public bool Reactivate;
            public string Formula;
            public List<Item> Items = new List<Item>();
            public IntPtr OurHwnd;
            public Func<bool> FrameworkFocus;   // WPF IsKeyboardFocusWithin / WinForms ContainsFocus
            public Action ReactivateAction;
            public IntPtr OwnerHwnd;
            public bool NoActivate;

            public Session(Ui ui, dynamic root, bool reactivate) { Ui = ui; Root = root; Reactivate = reactivate; }

            /// <summary>(Re)reads the root formula and resolves its references. Call in macro context.</summary>
            public void Build()
            {
                Formula = Convert.ToString(Root.Formula);
                var items = new List<Item> { MakeItem("(root)", Root) };
                foreach (Match m in RefRx.Matches(Formula ?? ""))
                    items.Add(Resolve(m, Root));
                Items = items;
            }

            /// <summary>UI event -> queue Goto as a macro, then log focus state. keyTs = Stopwatch timestamp of the keypress (0 = unknown).</summary>
            public void Navigate(int index, long keyTs)
            {
                if (index < 0 || index >= Items.Count) return;
                Item item = Items[index];
                long t0 = keyTs != 0 ? keyTs : Stopwatch.GetTimestamp();
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    string gotoError = null, activeBook = null;
                    try
                    {
                        if (item.Range == null) throw new InvalidOperationException("unresolved reference");
                        Goto((dynamic)item.Range);
                        activeBook = Convert.ToString(Xl.App.ActiveWorkbook.Name);
                    }
                    catch (Exception ex) { gotoError = (ex.InnerException ?? ex).Message; }
                    double elapsedMs = Math.Round((Stopwatch.GetTimestamp() - t0) * 1000.0 / Stopwatch.Frequency, 3);

                    object afterGoto = FocusState(), afterReactivate = null, reowned = null;
                    bool windowChanged = false;
                    try
                    {
                        var active = new IntPtr(Convert.ToInt64(Xl.App.ActiveWindow.Hwnd));
                        if (active != OwnerHwnd && active != IntPtr.Zero)
                        {
                            windowChanged = true;
                            Native.SetWindowLongPtr(OurHwnd, Native.GWLP_HWNDPARENT, active);
                            uint flags = Native.SWP_NOMOVE | Native.SWP_NOSIZE | Native.SWP_SHOWWINDOW | (NoActivate ? Native.SWP_NOACTIVATE : 0u);
                            Native.SetWindowPos(OurHwnd, Native.HWND_TOP, 0, 0, 0, 0, flags);
                            reowned = new { from = OwnerHwnd.ToInt64(), to = active.ToInt64() };
                            OwnerHwnd = active;
                        }
                    }
                    catch (Exception ex) { reowned = "reown failed: " + ex.Message; }

                    // Focused variants must take focus back after a workbook switch; the hook variant must not.
                    if ((Reactivate || (windowChanged && !NoActivate)) && ReactivateAction != null)
                    {
                        try { ReactivateAction(); afterReactivate = FocusState(); }
                        catch (Exception ex) { afterReactivate = "reactivate failed: " + ex.Message; }
                    }
                    Log.Write("k4nav", new
                    {
                        ui = Ui.ToString(), index, target = item.Address, gotoError, elapsedMs, activeBook,
                        afterGoto, reactivate = Reactivate, windowChanged, reowned, afterReactivate,
                    });

                    // Excel may take focus back after the macro returns; check again a bit later.
                    Task.Delay(250).ContinueWith(_ => ExcelAsyncUtil.QueueAsMacro(() =>
                        Log.Write("k4navLate", new { ui = Ui.ToString(), index, state = FocusState() })));
                });
            }

            /// <summary>Esc: go back to the root cell.</summary>
            public void GoBack()
            {
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try { Goto(Root); Log.Write("k4back", new { ui = Ui.ToString(), ok = true }); }
                    catch (Exception ex) { Log.Write("k4back", new { ui = Ui.ToString(), ok = false, error = ex.Message }); }
                });
            }

            public object FocusState()
            {
                IntPtr fg = Native.GetForegroundWindow(), focus = Native.GetFocus();
                bool? framework = null;
                try { framework = FrameworkFocus?.Invoke(); } catch { }
                return new
                {
                    foregroundIsOurs = fg == OurHwnd,
                    win32FocusIsOurs = focus != IntPtr.Zero && (focus == OurHwnd || Native.IsChild(OurHwnd, focus)),
                    frameworkFocus = framework,
                    foreground = fg.ToInt64(),
                    focus = focus.ToInt64(),
                    ours = OurHwnd.ToInt64(),
                };
            }
        }
    }

    /// <summary>
    /// K4 variant C: thread-local WH_KEYBOARD hook on Excel's main thread while the non-activating trace window is open.
    /// Swallows unmodified arrows/Enter/Esc and routes them to the window; F2 lets keys pass until Enter/Esc end the edit.
    /// </summary>
    internal static class KeyHook
    {
        private const int WH_KEYBOARD = 2, HC_ACTION = 0;
        private const int VK_RETURN = 0x0D, VK_SHIFT = 0x10, VK_CONTROL = 0x11, VK_MENU = 0x12, VK_ESCAPE = 0x1B;
        private const int VK_LEFT = 0x25, VK_UP = 0x26, VK_RIGHT = 0x27, VK_DOWN = 0x28, VK_F2 = 0x71;

        private static Native.HookProc _proc; // keep referenced: the hook holds only a native pointer
        private static IntPtr _hook;
        private static TraceWindowWpf _target;
        private static bool _editing;

        public static TraceWindowWpf Target => _target;

        public static void Install(TraceWindowWpf target)
        {
            Uninstall("reinstall");
            _target = target;
            _editing = false;
            _proc = Proc;
            _hook = Native.SetWindowsHookEx(WH_KEYBOARD, _proc, IntPtr.Zero, Native.GetCurrentThreadId());
            Log.Write("hook", new { action = "install", ok = _hook != IntPtr.Zero, win32Error = Marshal.GetLastWin32Error(), threadId = Native.GetCurrentThreadId() });
        }

        public static void Uninstall(string reason)
        {
            if (_hook != IntPtr.Zero)
            {
                bool ok = Native.UnhookWindowsHookEx(_hook);
                Log.Write("hook", new { action = "uninstall", reason, ok });
            }
            _hook = IntPtr.Zero;
            _target = null;
        }

        private static IntPtr Proc(int code, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (code == HC_ACTION && _target != null)
                {
                    int vk = wParam.ToInt32();
                    bool keyUp = (lParam.ToInt64() & 0x80000000L) != 0;
                    bool modifier = Down(VK_CONTROL) || Down(VK_SHIFT) || Down(VK_MENU);

                    if (vk == VK_F2)
                    {
                        if (!keyUp) { _editing = true; Log.Write("hookKey", new { vk, action = "F2 -> editing, pass through" }); }
                        return Next(code, wParam, lParam);
                    }
                    if (_editing)
                    {
                        if (!keyUp && (vk == VK_RETURN || vk == VK_ESCAPE))
                        {
                            _editing = false;
                            Log.Write("hookKey", new { vk, action = "edit ended, pass through" });
                            if (vk == VK_RETURN) _target.QueueRebuild();
                        }
                        return Next(code, wParam, lParam);
                    }

                    bool ours = vk == VK_UP || vk == VK_DOWN || vk == VK_LEFT || vk == VK_RIGHT || vk == VK_RETURN || vk == VK_ESCAPE;
                    if (ours && !modifier)
                    {
                        if (!keyUp)
                        {
                            long ts = Stopwatch.GetTimestamp();
                            Log.Write("hookKey", new { vk, action = "swallowed", foreground = Native.GetForegroundWindow().ToInt64(), focus = Native.GetFocus().ToInt64() });
                            TraceWindowWpf target = _target;
                            // Do not touch UI or unhook inside the hook callback; post to the dispatcher instead.
                            target.Dispatcher.BeginInvoke(new Action(() => target.HookKey(vk, ts)));
                        }
                        return new IntPtr(1); // swallow down and up
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("KeyHook.Proc", ex);
            }
            return Next(code, wParam, lParam);
        }

        private static bool Down(int vk) => (Native.GetKeyState(vk) & 0x8000) != 0;
        private static IntPtr Next(int code, IntPtr wParam, IntPtr lParam) => Native.CallNextHookEx(_hook, code, wParam, lParam);
    }

    internal static class Native
    {
        public delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)] public static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);
        [DllImport("user32.dll", SetLastError = true)] public static extern bool UnhookWindowsHookEx(IntPtr hhk);
        [DllImport("user32.dll")] public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] public static extern short GetKeyState(int nVirtKey);
        [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] public static extern IntPtr GetFocus();
        [DllImport("user32.dll")] public static extern bool IsChild(IntPtr hWndParent, IntPtr hWnd);
        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")] public static extern IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")] public static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint flags);

        public const int GWL_EXSTYLE = -20, GWLP_HWNDPARENT = -8;
        public const uint SWP_NOSIZE = 0x0001, SWP_NOMOVE = 0x0002, SWP_NOACTIVATE = 0x0010, SWP_SHOWWINDOW = 0x0040;
        public static readonly IntPtr HWND_TOP = IntPtr.Zero;
        public const long WS_EX_NOACTIVATE = 0x08000000L;
    }
}
