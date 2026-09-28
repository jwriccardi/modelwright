using System;
using System.Diagnostics;
using System.Text;
using ExcelDna.Integration;

namespace EmtSpike
{
    /// <summary>
    /// Macro entry points bound by OnKey. Each one maps 1:1 to a key in <see cref="AddIn.Keys"/>
    /// so the log records exactly which binding fired.
    /// </summary>
    public static class Commands
    {
        // ---- K3: Macabacus keys ------------------------------------------------------------
        [ExcelCommand(Name = "EmtNumberCycle")] public static void NumberCycle() => Run("NumberCycle", "^+1", () => Cycles.Format(Cycles.Number));
        [ExcelCommand(Name = "EmtDateCycle")] public static void DateCycle() => Run("DateCycle", "^+2", () => Cycles.Format(Cycles.Date));
        [ExcelCommand(Name = "EmtCurrencyCycle")] public static void CurrencyCycle() => Run("CurrencyCycle", "^+4", () => Cycles.Format(Cycles.Currency));
        [ExcelCommand(Name = "EmtPercentCycle")] public static void PercentCycle() => Run("PercentCycle", "^+5", () => Cycles.Format(Cycles.Percent));
        [ExcelCommand(Name = "EmtMultipleCycle")] public static void MultipleCycle() => Run("MultipleCycle", "^+8", () => Cycles.Format(Cycles.Multiple));
        [ExcelCommand(Name = "EmtFontColorCycle")] public static void FontColorCycle() => Run("FontColorCycle", "^'", Cycles.FontColor);
        [ExcelCommand(Name = "EmtFillCycle")] public static void FillCycle() => Run("FillCycle", "^+k", Cycles.Fill);
        [ExcelCommand(Name = "EmtFillCycleUpperK")] public static void FillCycleUpperK() => Run("FillCycle", "^+K", Cycles.Fill);
        [ExcelCommand(Name = "EmtBlueBlackToggle")] public static void BlueBlackToggle() => Run("BlueBlackToggle", "^;", Cycles.BlueBlack);
        [ExcelCommand(Name = "EmtDecimalsIncrease")] public static void DecimalsIncrease() => Run("DecimalsIncrease", "^,", () => "log only");
        [ExcelCommand(Name = "EmtDecimalsDecrease")] public static void DecimalsDecrease() => Run("DecimalsDecrease", "^.", () => "log only");
        [ExcelCommand(Name = "EmtTracePrecedents")] public static void TracePrecedents() => Run("ProPrecedents(WPF)", "^+{[}", () => Trace.Open(Trace.Ui.Wpf, "^+{[}", false));
        [ExcelCommand(Name = "EmtTracePrecedentsBrace")] public static void TracePrecedentsBrace() => Run("ProPrecedents(WPF)", "^{{}", () => Trace.Open(Trace.Ui.Wpf, "^{{}", false));
        [ExcelCommand(Name = "EmtShowAllPrecedents")] public static void ShowAllPrecedents() => Run("ShowAllPrecedents", "^%{[}", () => "log only");
        [ExcelCommand(Name = "EmtLastAuditedCell")] public static void LastAuditedCell() => Run("LastAuditedCell", "^+\\", Trace.GotoLastAudited);
        [ExcelCommand(Name = "EmtLastAuditedCellPipe")] public static void LastAuditedCellPipe() => Run("LastAuditedCell", "^|", Trace.GotoLastAudited);
        [ExcelCommand(Name = "EmtOverride")] public static void Override() => Run("Override", "^%+{F1}", () => { AddIn.RegisterAll("Override"); return "all keys re-registered"; });

        // ---- K2: undo variants --------------------------------------------------------------
        [ExcelCommand(Name = "EmtK2Control")] public static void K2Control() => Run("K2.1 COM Font.Bold", "^%+{F5}", K2.Control);
        [ExcelCommand(Name = "EmtK2MsoDirect")] public static void K2MsoDirect() => Run("K2.2 ExecuteMso Bold direct", "^%+{F6}", K2.MsoBoldDirect);
        [ExcelCommand(Name = "EmtK2MsoDeferred")] public static void K2MsoDeferred() => Run("K2.3 ExecuteMso Bold deferred", "^%+{F7}", K2.MsoBoldDeferred);
        [ExcelCommand(Name = "EmtK2PercentMso")] public static void K2PercentMso() => Run("K2.4 ExecuteMso PercentStyle", "^%+{F8}", K2.MsoPercentStyle);
        [ExcelCommand(Name = "EmtK2CopyPasteSpecial")] public static void K2CopyPasteSpecial() => Run("K2.5 Copy+PasteSpecial formats", "^%+{F9}", K2.CopyPasteSpecial);
        [ExcelCommand(Name = "EmtK2CopyPasteMso")] public static void K2CopyPasteMso() => Run("K2.6 Copy+ExecuteMso PasteFormatting", "^%+{F10}", K2.CopyPasteMso);
        [ExcelCommand(Name = "EmtK2XlmFormatNumber")] public static void K2XlmFormatNumber() => Run("K2.7 xlcFormatNumber", "^%+{F11}", K2.XlmFormatNumber);
        [ExcelCommand(Name = "EmtK2Snapshot")] public static void K2Snapshot() => Run("K2.0 snapshot", "^%+{F12}", () => UndoProbe.Snapshot("manual", "K2.0"));

        // ---- K4 -----------------------------------------------------------------------------
        [ExcelCommand(Name = "EmtTraceWinForms")] public static void TraceWinForms() => Run("ProPrecedents(WinForms)", "^%+{F2}", () => Trace.Open(Trace.Ui.WinForms, "^%+{F2}", false));
        [ExcelCommand(Name = "EmtTraceHook")] public static void TraceHook() => Run("ProPrecedents(hook)", "^%+{F3}", () => Trace.Open(Trace.Ui.Hook, "^%+{F3}", false));

        /// <summary>Ribbon dispatch (always invoked via QueueAsMacro, so we are in macro context).</summary>
        internal static void Ribbon(string id)
        {
            const string k = "ribbon";
            switch (id)
            {
                case "k2_1": Run("K2.1 COM Font.Bold", k, K2.Control); break;
                case "k2_2": Run("K2.2 ExecuteMso Bold direct", k, K2.MsoBoldDirect); break;
                case "k2_3": Run("K2.3 ExecuteMso Bold deferred", k, K2.MsoBoldDeferred); break;
                case "k2_4": Run("K2.4 ExecuteMso PercentStyle", k, K2.MsoPercentStyle); break;
                case "k2_5": Run("K2.5 Copy+PasteSpecial formats", k, K2.CopyPasteSpecial); break;
                case "k2_6": Run("K2.6 Copy+ExecuteMso PasteFormatting", k, K2.CopyPasteMso); break;
                case "k2_7": Run("K2.7 xlcFormatNumber", k, K2.XlmFormatNumber); break;
                case "k2_0": Run("K2.0 snapshot", k, () => UndoProbe.Snapshot("manual", "K2.0")); break;
                case "openLog": Run("OpenLogFolder", k, OpenLogFolder); break;
                case "fixtures": Run("CreateK4Fixtures", k, Fixtures.Create); break;
                case "traceWpf": Run("ProPrecedents(WPF)", k, () => Trace.Open(Trace.Ui.Wpf, k, false)); break;
                case "traceWpfReact": Run("ProPrecedents(WPF,reactivate)", k, () => Trace.Open(Trace.Ui.Wpf, k, true)); break;
                case "traceWinForms": Run("ProPrecedents(WinForms)", k, () => Trace.Open(Trace.Ui.WinForms, k, false)); break;
                case "traceWinFormsReact": Run("ProPrecedents(WinForms,reactivate)", k, () => Trace.Open(Trace.Ui.WinForms, k, true)); break;
                case "traceHook": Run("ProPrecedents(hook)", k, () => Trace.Open(Trace.Ui.Hook, k, false)); break;
                case "override": Run("Override", k, () => { AddIn.RegisterAll("Override(ribbon)"); return "re-registered"; }); break;
                default: Log.Write("ribbonUnknown", new { id }); break;
            }
        }

        /// <summary>Times the action, logs {cmd,key,elapsedMs,selectionCellCount,method}, sets the status bar. Never throws.</summary>
        internal static void Run(string cmd, string key, Func<object> action)
        {
            var sw = Stopwatch.StartNew();
            try
            {
                object detail = action();
                sw.Stop();
                double ms = Xl.Ms(sw);
                Log.Write("cmd", new
                {
                    cmd, key, elapsedMs = ms,
                    selectionCellCount = Xl.SelectionCount(),
                    method = AddIn.MethodFor(key),
                    detail,
                });
                Xl.Status($"EMT spike: {Pretty(key)} {cmd} fired ({ms:0.0} ms)");
            }
            catch (Exception ex)
            {
                sw.Stop();
                Log.Write("cmdError", new { cmd, key, elapsedMs = Xl.Ms(sw), method = AddIn.MethodFor(key), type = ex.GetType().Name, message = ex.Message, stack = ex.ToString() });
                Xl.Status($"EMT spike: {Pretty(key)} {cmd} FAILED: {ex.Message}");
            }
        }

        private static object OpenLogFolder()
        {
            System.IO.Directory.CreateDirectory(Log.Dir);
            Process.Start("explorer.exe", "\"" + Log.Dir + "\"");
            return Log.Dir;
        }

        /// <summary>"^+{[}" -> "Ctrl+Shift+[".</summary>
        internal static string Pretty(string key)
        {
            if (string.IsNullOrEmpty(key) || key == "ribbon") return "Ribbon";
            var sb = new StringBuilder();
            int i = 0;
            for (; i < key.Length && "^+%".IndexOf(key[i]) >= 0; i++)
                sb.Append(key[i] == '^' ? "Ctrl+" : key[i] == '+' ? "Shift+" : "Alt+");
            string rest = key.Substring(i);
            if (rest.Length > 2 && rest[0] == '{' && rest[rest.Length - 1] == '}') rest = rest.Substring(1, rest.Length - 2);
            return sb + rest;
        }
    }
}
