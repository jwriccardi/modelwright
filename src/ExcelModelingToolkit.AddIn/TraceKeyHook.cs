using System;
using System.Collections.Generic;
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
/// Edit or Point) is read at every candidate press, the check <see cref="UndoKeyHook"/> already relies on. So while a
/// cell is being edited the arrows, Enter, Esc and F2 go to Excel until the edit ends, however it was started (F2,
/// typing, a double-click, the formula bar) or ended. Keys also pass while the focus is not on a worksheet grid or
/// this window (a dialog, the Name Box), and while the mouse is captured, a menu is open or a window is being moved
/// or sized (no Excel call is made then).
/// </para>
/// <para>
/// <b>F2.</b> Outside an edit, F2 is a Trace In command (<see cref="TraceKeyCommand.EditReference"/>): the session
/// sends Excel the keys that edit the selected reference (<see cref="Send"/>, with <c>SendInput</c>). Those keys are
/// partly Trace In keys and arrive before Excel is editing, so while they are in flight the hook passes them
/// untouched, counting them off (<see cref="SyntheticKeyTracker"/>) until the last has arrived, a press that is not
/// one of them arrives, the window closes, or <see cref="SynthesisTimeoutMilliseconds"/> pass. They can include the
/// Go To step (<see cref="ReferenceEdit"/>): for a target in another workbook Ctrl+Tab (to the target's window), then
/// F5, characters typed into Excel's Go To dialog (sent as Unicode characters, seen here as <c>VK_PACKET</c>: the
/// dialog runs on Excel's thread, so its keys pass through this hook too) and Enter. They are sent in batches that end
/// after Ctrl+Tab and after F5 (<see cref="ReferenceEdit.Batches"/>). The batch after Ctrl+Tab (F5) goes only once the
/// target's window is the foreground window and Excel's thread has a keyboard focus (polled every
/// <see cref="SwitchPollMilliseconds"/>, at most <see cref="SwitchWaitMilliseconds"/>); if Ctrl+Tab brought another
/// workbook window to the front, Ctrl+Tab is sent again, at most once per visible workbook window. The batch after F5,
/// the text typed into Go To, goes only once Excel's Go To dialog is the foreground window and has the keyboard focus
/// (polled every <see cref="GoToPollMilliseconds"/>, at most <see cref="GoToWaitMilliseconds"/>). If a wait runs out,
/// the rest is not sent (Excel stays in Point mode, where Esc cancels). If the time runs out first, keys already sent
/// still reach Excel: no Trace In command is taken while Excel is editing or a dialog has the focus. After such an edit,
/// the Enter, Tab or Esc that ends it is passed to Excel and also tells the session to check the formula. Where F5, the
/// first and last character of each Go To text and Enter landed, and which window each Ctrl+Tab brought to the front
/// (the foreground window and the focused window's class as the hook saw them) is logged as one
/// <c>TraceSynthKeys</c> line.
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

    /// <summary>
    /// How long, at most, after each batch is sent, keys sent with <see cref="Send"/> are passed untouched while they arrive. Once Excel has
    /// taken the first F2 it is editing, and the hook passes the rest anyway.
    /// </summary>
    private const int SynthesisTimeoutMilliseconds = 2000;

    /// <summary>How long <see cref="Send"/> waits, at most, for the user to let go of F2 and the modifiers.</summary>
    private const int ReleaseWaitMilliseconds = 1000;

    /// <summary>How often, while a batch waits for the Go To dialog, the hook checks whether it has the focus.</summary>
    private const int GoToPollMilliseconds = 20;

    /// <summary>How long, at most, a batch waits for the Go To dialog after the previous one was sent.</summary>
    private const int GoToWaitMilliseconds = 1500;

    /// <summary>How often, after Ctrl+Tab, the hook checks whether the target's window is in front.</summary>
    private const int SwitchPollMilliseconds = 20;

    /// <summary>How long, at most, the batch after Ctrl+Tab waits for the target's window.</summary>
    private const int SwitchWaitMilliseconds = 1500;

    /// <summary>The title of Excel's Go To dialog (F5).</summary>
    private const string GoToTitle = "Go To";

    /// <summary>The most keys the <c>TraceSynthKeys</c> line describes, and the most characters of a window title in it.</summary>
    private const int MaxProbes = 16;
    private const int MaxProbeTitle = 40;

    private const int VkReturn = 0x0D;
    private const int VkF2 = 0x71;
    private const int VkF5 = 0x74;
    private const uint InputKeyboard = 1;
    private const uint KeyEventFExtendedKey = 0x1;
    private const uint KeyEventFKeyUp = 0x2;
    private const uint KeyEventFUnicode = 0x4;

    // The hook holds only a native pointer to the delegate: it stays referenced for the life of the AppDomain (a
    // call can still be on its way in while the hook is being removed).
    private static readonly NativeMethods.HookProc HookCallback = Proc;

    // Where the keys named in _probeNames landed so far (the TraceSynthKeys line).
    private static readonly List<string> Probes = new List<string>();

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

    // Keys sent with Send that the hook has not seen yet (null when none are in flight), when they were sent (a
    // Stopwatch timestamp), and the timer that gives up on them.
    private static SyntheticKeyTracker? _synthetic;
    private static long _syntheticSent;
    private static Timer? _syntheticTimer;

    // The batches of those keys (ReferenceEdit.Batches), how many batches and keys have been sent, and the timer that
    // sends the next batch once the Go To dialog has the focus (until _gateUntil, a Stopwatch timestamp).
    private static IReadOnlyList<IReadOnlyList<SyntheticKey>>? _batches;
    private static int _batchesSent;
    private static int _keysSent;
    private static Timer? _gateTimer;
    private static long _gateUntil;

    // After a Ctrl+Tab: the window it should bring to the front (the target's; zero when the keys hold no Ctrl+Tab),
    // the one in front when the last Ctrl+Tab was sent, and how many more Ctrl+Tabs were sent (SwitchTick).
    private static IntPtr _switchTo;
    private static IntPtr _switchFrom;
    private static int _ctrlTabs;

    // For each sent key, its name in the TraceSynthKeys line if where it lands is logged (else null); the line's
    // entries so far are in Probes.
    private static string?[] _probeNames = new string?[0];

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
        EndSynthesis("the window closed");
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
            EndSynthesis("the hook was removed");
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
    /// Sends <paramref name="keys"/> to Excel as if typed (one <c>SendInput</c> batch per <see cref="ReferenceEdit.Batches"/>
    /// batch, so none of the user's keys come in between; a batch after Ctrl+Tab once <paramref name="switchTo"/>, the
    /// target's window, is in front; a batch after F5 once the Go To dialog has the focus), with the hook passing them
    /// untouched as they arrive (see the remarks). Waits first, up to <see cref="ReleaseWaitMilliseconds"/>, for the
    /// user to let go of F2, Shift, Ctrl and Alt, so a held key does not change what the keys do. Only to an Excel
    /// workbook window in the foreground (<paramref name="window"/>, if given), and only while the hook is installed.
    /// Returns false, with the reason in <paramref name="failure"/>, if nothing was sent. Main thread, outside the hook.
    /// </summary>
    internal static bool Send(IReadOnlyList<SyntheticKey> keys, out string failure, IntPtr window = default, IntPtr switchTo = default)
    {
        failure = string.Empty;
        if (!IsInstalled || _poster is null)
        {
            failure = "the keyboard is not connected";
            return false;
        }

        var waited = Stopwatch.StartNew();
        while (IsHeld(VkF2) || IsHeld(VkShift) || IsHeld(VkControl) || IsHeld(VkMenu))
        {
            if (waited.ElapsedMilliseconds > ReleaseWaitMilliseconds)
            {
                failure = "a key is still held down (F2, Shift, Ctrl or Alt)";
                return false;
            }

            System.Threading.Thread.Sleep(10);
        }

        var foreground = NativeMethods.GetForegroundWindow();
        if (!IsExcelWorkbookWindow(foreground))
        {
            failure = "Excel's workbook window is not in front";
            return false;
        }

        if (window != IntPtr.Zero && !SameWindow(foreground, window))
        {
            failure = "the edited cell's workbook window is not in front";
            return false;
        }

        EndSynthesis("replaced by new keys");
        _synthetic = new SyntheticKeyTracker(keys);
        _syntheticSent = Stopwatch.GetTimestamp();
        _batches = ReferenceEdit.Batches(keys);
        _batchesSent = 0;
        _keysSent = 0;
        _probeNames = ProbeNames(keys);
        _switchTo = switchTo;
        _switchFrom = foreground;
        _ctrlTabs = 0;
        Probes.Clear();

        var sendFailure = SendBatch();
        if (sendFailure is not null)
        {
            failure = sendFailure;
            return false;
        }

        return true;
    }

    // Sends the next batch (one SendInput call, so none of the user's keys come in between) and restarts the timeout;
    // if another batch follows, starts waiting for the target's window after a Ctrl+Tab (SwitchTick), else for the Go
    // To dialog (GoToTick). Null if sent; else why not, and the synthesis is over. Main thread, outside the hook.
    private static string? SendBatch()
    {
        var batches = _batches;
        if (_synthetic is null || batches is null || _batchesSent >= batches.Count)
        {
            return "nothing to send";
        }

        var batch = batches[_batchesSent];
        _batchesSent++;
        var failure = SendKeys(batch);
        if (failure is not null)
        {
            return failure;
        }

        if (_batchesSent < batches.Count)
        {
            var switching = ReferenceEdit.EndsWithCtrlTab(batch);
            _gateUntil = Stopwatch.GetTimestamp() + ((switching ? SwitchWaitMilliseconds : GoToWaitMilliseconds) * Stopwatch.Frequency / 1000);
            var gate = new Timer { Interval = switching ? SwitchPollMilliseconds : GoToPollMilliseconds };
            if (switching)
            {
                gate.Tick += (sender, e) => SwitchTick();
            }
            else
            {
                gate.Tick += (sender, e) => GoToTick();
            }

            gate.Start();
            _gateTimer = gate;
        }

        return null;
    }

    // Sends keys of the sequence in one SendInput call and restarts the timeout. Null if sent; else why not, and the
    // synthesis is over. Main thread, outside the hook.
    private static string? SendKeys(IReadOnlyList<SyntheticKey> keys)
    {
        var inputs = new NativeMethods.Input[keys.Count];
        for (var i = 0; i < keys.Count; i++)
        {
            inputs[i] = KeyInput(keys[i]);
        }

        StopTimer(ref _syntheticTimer);
        var timer = new Timer { Interval = SynthesisTimeoutMilliseconds };
        timer.Tick += (sender, e) => EndSynthesis("timed out");
        timer.Start();
        _syntheticTimer = timer;

        _keysSent += keys.Count;
        var sent = NativeMethods.SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(NativeMethods.Input)));
        if (sent != inputs.Length)
        {
            var error = Marshal.GetLastWin32Error();
            var failure = string.Format(CultureInfo.InvariantCulture, "Windows took {0} of {1} key events (error {2})", sent, inputs.Length, error);
            EndSynthesis("SendInput failed: " + failure);
            return failure;
        }

        return null;
    }

    // The window timer after a Ctrl+Tab: once the keys sent so far have all arrived, sends the next batch (F5) when the
    // target's window (_switchTo) is in front with the keyboard focus on Excel's thread; if Ctrl+Tab brought another
    // workbook window to the front instead, sends Ctrl+Tab again (at most once per workbook window) and waits on.
    // Gives up (the rest is not sent; Excel stays in Point mode, where Esc cancels) after SwitchWaitMilliseconds. Each
    // window a Ctrl+Tab brought to the front is a TraceSynthKeys entry. Never throws.
    private static void SwitchTick()
    {
        try
        {
            var synthetic = _synthetic;
            if (synthetic is null || _batches is null)
            {
                StopTimer(ref _gateTimer);
                return;
            }

            if (synthetic.Seen >= _keysSent)
            {
                var foreground = NativeMethods.GetForegroundWindow();
                var info = new NativeMethods.GuiThreadInfo { Size = Marshal.SizeOf(typeof(NativeMethods.GuiThreadInfo)) };
                if (_switchTo != IntPtr.Zero && SameWindow(foreground, _switchTo) &&
                    NativeMethods.GetGUIThreadInfo(_thread, ref info) && info.Focus != IntPtr.Zero)
                {
                    AddProbe("Ctrl+Tab@" + WhereKeysGo());
                    StopTimer(ref _gateTimer);
                    SendBatch();
                    return;
                }

                if (!SameWindow(foreground, _switchFrom) && !SameWindow(foreground, _switchTo) && IsExcelWorkbookWindow(foreground) &&
                    _ctrlTabs < WorkbookWindowCount())
                {
                    // Another workbook's window (a third workbook ahead of the target in Excel's window order).
                    AddProbe("Ctrl+Tab@" + WhereKeysGo() + "(not the target: again)");
                    _switchFrom = foreground;
                    _ctrlTabs++;
                    var ctrlTab = ReferenceEdit.CtrlTab;
                    synthetic.InsertNext(ctrlTab);
                    var names = new List<string?>(_probeNames);
                    names.InsertRange(_keysSent, new string?[ctrlTab.Count]);
                    _probeNames = names.ToArray();
                    SendKeys(ctrlTab);
                    return;
                }
            }

            if (Stopwatch.GetTimestamp() >= _gateUntil)
            {
                AddProbe("Ctrl+Tab@" + WhereKeysGo());
                EndSynthesis(string.Format(
                    CultureInfo.InvariantCulture,
                    "the target's window did not come to the front within {0} ms of Ctrl+Tab ({1} more Ctrl+Tab; keys arrived={2}/{3}; foreground={4}): the rest was not sent, Excel stays in Point mode",
                    SwitchWaitMilliseconds,
                    _ctrlTabs,
                    synthetic.Seen,
                    _keysSent,
                    WhereKeysGo()));
            }
        }
        catch (Exception ex)
        {
            EndSynthesis("waiting for the target's window failed: " + ex.Message);
        }
    }

    // Window handles are 32-bit values (Window.Hwnd is an Int32).
    private static bool SameWindow(IntPtr a, IntPtr b) => (a.ToInt64() & 0xFFFFFFFFL) == (b.ToInt64() & 0xFFFFFFFFL);

    // The number of visible workbook windows (XLMAIN) on Excel's thread: how many windows Ctrl+Tab can cycle through.
    private static int WorkbookWindowCount()
    {
        var count = 0;
        NativeMethods.EnumThreadWindows(_thread, (window, data) =>
        {
            if (NativeMethods.IsWindowVisible(window) && IsExcelWorkbookWindow(window))
            {
                count++;
            }

            return true;
        }, IntPtr.Zero);
        return count;
    }

    // The Go To timer: sends the next batch once the keys sent so far have all arrived and the Go To dialog has the
    // focus; gives up (the rest is not sent) after GoToWaitMilliseconds. Runs from whatever message loop Excel is in,
    // the dialog's too. Never throws.
    private static void GoToTick()
    {
        try
        {
            var synthetic = _synthetic;
            if (synthetic is null || _batches is null)
            {
                StopTimer(ref _gateTimer);
                return;
            }

            if (synthetic.Seen >= _keysSent && GoToHasFocus())
            {
                StopTimer(ref _gateTimer);
                SendBatch();
                return;
            }

            if (Stopwatch.GetTimestamp() >= _gateUntil)
            {
                EndSynthesis(string.Format(
                    CultureInfo.InvariantCulture,
                    "the Go To dialog did not get the focus within {0} ms (keys arrived={1}/{2}; foreground={3}): the rest was not sent",
                    GoToWaitMilliseconds,
                    synthetic.Seen,
                    _keysSent,
                    WhereKeysGo()));
            }
        }
        catch (Exception ex)
        {
            EndSynthesis("waiting for the Go To dialog failed: " + ex.Message);
        }
    }

    // True if the foreground window is a visible top-level window titled Go To on Excel's thread, and the keyboard
    // focus on that thread is in it.
    private static bool GoToHasFocus()
    {
        var foreground = NativeMethods.GetForegroundWindow();
        if (foreground == IntPtr.Zero || !NativeMethods.IsWindowVisible(foreground) ||
            NativeMethods.GetAncestor(foreground, GaRoot) != foreground ||
            NativeMethods.GetWindowThreadProcessId(foreground, out _) != _thread ||
            !string.Equals(WindowTitle(foreground), GoToTitle, StringComparison.Ordinal))
        {
            return false;
        }

        var info = new NativeMethods.GuiThreadInfo { Size = Marshal.SizeOf(typeof(NativeMethods.GuiThreadInfo)) };
        return NativeMethods.GetGUIThreadInfo(_thread, ref info) && info.Focus != IntPtr.Zero &&
            (info.Focus == foreground || NativeMethods.IsChild(foreground, info.Focus));
    }

    // "[title]/CLASS": the foreground window's title (shortened) and the class of the window with the keyboard focus
    // on Excel's thread. Never throws.
    private static string WhereKeysGo()
    {
        try
        {
            var title = WindowTitle(NativeMethods.GetForegroundWindow());
            if (title.Length > MaxProbeTitle)
            {
                title = title.Substring(0, MaxProbeTitle) + "...";
            }

            var focusClass = "none";
            var info = new NativeMethods.GuiThreadInfo { Size = Marshal.SizeOf(typeof(NativeMethods.GuiThreadInfo)) };
            if (NativeMethods.GetGUIThreadInfo(_thread, ref info) && info.Focus != IntPtr.Zero)
            {
                var name = new StringBuilder(64);
                focusClass = NativeMethods.GetClassName(info.Focus, name, name.Capacity) > 0 ? name.ToString() : "?";
            }

            return "[" + title + "]/" + focusClass;
        }
        catch (Exception ex)
        {
            return "(" + ex.GetType().Name + ")";
        }
    }

    private static string WindowTitle(IntPtr window)
    {
        if (window == IntPtr.Zero)
        {
            return string.Empty;
        }

        var text = new StringBuilder(256);
        return NativeMethods.InternalGetWindowText(window, text, text.Capacity) > 0 ? text.ToString() : string.Empty;
    }

    // The keys whose landing the TraceSynthKeys line records: each F5 and Enter press, and the first and last
    // character typed of each Go To text (presses).
    private static string?[] ProbeNames(IReadOnlyList<SyntheticKey> keys)
    {
        var names = new string?[keys.Count];
        for (var i = 0; i < keys.Count; i++)
        {
            var key = keys[i];
            if (key.KeyUp)
            {
                continue;
            }

            if (key.IsCharacter)
            {
                var first = i == 0 || !keys[i - 1].IsCharacter;
                var last = i + 2 >= keys.Count || !keys[i + 2].IsCharacter;
                names[i] = first || last ? "'" + key.Character + "'" : null;
            }
            else if (key.VirtualKey == VkF5)
            {
                names[i] = "F5";
            }
            else if (key.VirtualKey == VkReturn)
            {
                names[i] = "Enter";
            }
        }

        return names;
    }

    // Stops and disposes a timer. Never throws.
    private static void StopTimer(ref Timer? timer)
    {
        try
        {
            timer?.Stop();
            timer?.Dispose();
        }
        catch (Exception)
        {
            // A timer that will not stop only ends a wait that is already over.
        }

        timer = null;
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

    // Stops passing sent keys untouched (all arrived, another key, the timeout, the target's window not coming to the
    // front, the Go To dialog not opening, the window closed); batches not sent yet never are. Never throws.
    private static void EndSynthesis(string how)
    {
        var synthetic = _synthetic;
        if (synthetic is null)
        {
            return;
        }

        _synthetic = null;
        StopTimer(ref _syntheticTimer);
        StopTimer(ref _gateTimer);
        var batches = _batches?.Count ?? 0;
        _batches = null;
        var fields = new[]
        {
            how,
            "keys=" + synthetic.Seen.ToString(CultureInfo.InvariantCulture) + "/" + synthetic.Count.ToString(CultureInfo.InvariantCulture),
            "batches=" + _batchesSent.ToString(CultureInfo.InvariantCulture) + "/" + batches.ToString(CultureInfo.InvariantCulture),
            "ms=" + ((Stopwatch.GetTimestamp() - _syntheticSent) * 1000.0 / Stopwatch.Frequency).ToString("0.0", CultureInfo.InvariantCulture),
        };
        var probes = Probes.Count == 0 ? null : string.Join(" ", Probes);
        Probes.Clear();
        _probeNames = new string?[0];

        // From the hook only posted, as every log line there; elsewhere written at once (the hook may be going).
        if (_inProc)
        {
            if (probes is not null)
            {
                Log("TraceSynthKeys", probes);
            }

            Log("TraceSynth", fields);
        }
        else
        {
            if (probes is not null)
            {
                DiagnosticsLog.Write("TraceSynthKeys", probes);
            }

            DiagnosticsLog.Write("TraceSynth", fields);
        }
    }

    // In the hook, for a sent key it has just seen (index): where it landed, if the TraceSynthKeys line records it.
    private static void Probe(int index)
    {
        try
        {
            if (index < 0 || index >= _probeNames.Length || _probeNames[index] is not string name)
            {
                return;
            }

            AddProbe(name + "@" + WhereKeysGo());
        }
        catch (Exception)
        {
            // A diagnostic only.
        }
    }

    // An entry of the TraceSynthKeys line, unless it is full.
    private static void AddProbe(string entry)
    {
        if (Probes.Count < MaxProbes)
        {
            Probes.Add(entry);
        }
        else if (Probes.Count == MaxProbes)
        {
            Probes.Add("...");
        }
    }

    private static NativeMethods.Input KeyInput(SyntheticKey key)
    {
        if (key.IsCharacter)
        {
            // Typed as the character itself, whatever the keyboard layout: Windows delivers it as VK_PACKET.
            return new NativeMethods.Input
            {
                Type = InputKeyboard,
                Union = new NativeMethods.InputUnion
                {
                    Keyboard = new NativeMethods.KeyboardInput
                    {
                        ScanCode = key.Character,
                        Flags = KeyEventFUnicode | (key.KeyUp ? KeyEventFKeyUp : 0),
                    },
                },
            };
        }

        // The navigation keys (PgUp to Del) are extended keys: without the flag an arrow is the numeric keypad's,
        // which Windows may turn into a digit (Num Lock) or answer with a fake Shift release (Shift+Right).
        var extended = key.VirtualKey >= 0x21 && key.VirtualKey <= 0x2E;
        return new NativeMethods.Input
        {
            Type = InputKeyboard,
            Union = new NativeMethods.InputUnion
            {
                Keyboard = new NativeMethods.KeyboardInput
                {
                    VirtualKey = (ushort)key.VirtualKey,
                    ScanCode = (ushort)NativeMethods.MapVirtualKey((uint)key.VirtualKey, 0),
                    Flags = (key.KeyUp ? KeyEventFKeyUp : 0) | (extended ? KeyEventFExtendedKey : 0),
                },
            },
        };
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
            // Keys we sent (Send): passed untouched until the last one. A stray release is handled as usual below;
            // any other press ends them and is handled as usual.
            if (code == HcAction && _synthetic is SyntheticKeyTracker synthetic)
            {
                var key = wParam.ToInt32();
                var keyUp = (lParam.ToInt64() & 0x80000000L) != 0;
                switch (synthetic.Observe(key, keyUp))
                {
                    case SyntheticKeyMatch.Expected:
                        Probe(synthetic.Seen - 1);
                        return NativeMethods.CallNextHookEx(_hook, code, wParam, lParam);
                    case SyntheticKeyMatch.Completed:
                        Probe(synthetic.Seen - 1);
                        EndSynthesis("complete");
                        return NativeMethods.CallNextHookEx(_hook, code, wParam, lParam);
                    case SyntheticKeyMatch.Unexpected:
                        EndSynthesis("ended by another key (0x" + key.ToString("X2", CultureInfo.InvariantCulture) + ")");
                        break;
                }
            }

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
                    var modifiers = Modifiers();

                    // The key that ends an F2 edit goes to Excel (below); after it the session checks the formula.
                    if (session.AwaitsEditEnd && TraceKeys.EndsEdit(key, modifiers) && Context() == TraceKeyContext.Editing)
                    {
                        session.Enqueue(() => session.EndEdit(goBack: true, "key"));
                    }

                    var command = TraceKeys.Command(key, modifiers);
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

    // Physically down now (IsDown: as of the key message being handled).
    private static bool IsHeld(int virtualKey) => (NativeMethods.GetAsyncKeyState(virtualKey) & 0x8000) != 0;

    private static class NativeMethods
    {
        public delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

        [return: MarshalAs(UnmanagedType.Bool)]
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumThreadWindows(uint dwThreadId, EnumWindowsProc lpfn, IntPtr lParam);

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
        public static extern short GetAsyncKeyState(int vKey);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint SendInput(uint nInputs, [In] Input[] pInputs, int cbSize);

        [DllImport("user32.dll")]
        public static extern uint MapVirtualKey(uint uCode, uint uMapType);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

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

        // GetWindowText without WM_GETTEXT: a window of Excel's thread is not called from inside the hook.
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int InternalGetWindowText(IntPtr hWnd, StringBuilder pString, int cchMaxCount);

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

        /// <summary>INPUT (used for keys only; the union is as large as its largest member, MOUSEINPUT).</summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct Input
        {
            public uint Type;
            public InputUnion Union;
        }

        /// <summary>INPUT's union.</summary>
        [StructLayout(LayoutKind.Explicit)]
        public struct InputUnion
        {
            [FieldOffset(0)]
            public MouseInput Mouse;

            [FieldOffset(0)]
            public KeyboardInput Keyboard;
        }

        /// <summary>MOUSEINPUT (only for the union's size).</summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct MouseInput
        {
            public int Dx;
            public int Dy;
            public uint MouseData;
            public uint Flags;
            public uint Time;
            public IntPtr ExtraInfo;
        }

        /// <summary>KEYBDINPUT.</summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct KeyboardInput
        {
            public ushort VirtualKey;
            public ushort ScanCode;
            public uint Flags;
            public uint Time;
            public IntPtr ExtraInfo;
        }
    }
}
