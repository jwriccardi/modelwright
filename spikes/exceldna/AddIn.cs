using System;
using System.Collections.Generic;
using System.Diagnostics;
using ExcelDna.Integration;

namespace EmtSpike
{
    /// <summary>AutoOpen/AutoClose: key registration (K3) and hidden-workbook pre-creation (K2).</summary>
    public sealed class AddIn : IExcelAddIn
    {
        /// <summary>OnKey string -> macro name. Every key maps to its own macro so the log shows which binding fired.</summary>
        internal static readonly (string Key, string Macro)[] Keys =
        {
            // K3: Macabacus keys
            ("^+1", "EmtNumberCycle"),
            ("^+2", "EmtDateCycle"),
            ("^+4", "EmtCurrencyCycle"),
            ("^+5", "EmtPercentCycle"),
            ("^+8", "EmtMultipleCycle"),
            ("^'", "EmtFontColorCycle"),
            ("^+k", "EmtFillCycle"),
            ("^+K", "EmtFillCycleUpperK"),
            ("^;", "EmtBlueBlackToggle"),
            ("^,", "EmtDecimalsIncrease"),
            ("^.", "EmtDecimalsDecrease"),
            ("^+{[}", "EmtTracePrecedents"),
            ("^{{}", "EmtTracePrecedentsBrace"),
            ("^%{[}", "EmtShowAllPrecedents"),
            ("^+\\", "EmtLastAuditedCell"),
            ("^|", "EmtLastAuditedCellPipe"),
            // Spike-only commands live on Ctrl+Alt+Shift+F-keys to stay off Macabacus keys.
            ("^%+{F1}", "EmtOverride"),
            // K2: undo variants
            ("^%+{F5}", "EmtK2Control"),
            ("^%+{F6}", "EmtK2MsoDirect"),
            ("^%+{F7}", "EmtK2MsoDeferred"),
            ("^%+{F8}", "EmtK2PercentMso"),
            ("^%+{F9}", "EmtK2CopyPasteSpecial"),
            ("^%+{F10}", "EmtK2CopyPasteMso"),
            ("^%+{F11}", "EmtK2XlmFormatNumber"),
            ("^%+{F12}", "EmtK2Snapshot"),
            // K4: WinForms (B) and non-activating + hook (C) trace variants
            ("^%+{F2}", "EmtTraceWinForms"),
            ("^%+{F3}", "EmtTraceHook"),
        };

        /// <summary>
        /// K3 coverage set: remaining Macabacus keys. Each only logs {key, name} + status bar,
        /// to learn whether every punctuation/named key can be bound. Macros are registered at runtime.
        /// </summary>
        internal static readonly (string Key, string Name)[] Coverage =
        {
            ("%+;", "RatioCycle"),
            ("%+,", "ShiftDecimalLeft"),
            ("%+.", "ShiftDecimalRight"),
            ("^%+'", "BorderColorCycle"),
            ("^%.", "AutoColorCycle"),
            ("^%'", "CommentFormula"),
            ("^+{]}", "TraceOut"),
            ("^%+{[}", "AutoTracePrecedents"),
            ("^%+{]}", "AutoTraceDependents"),
            ("^%{]}", "ShowAllDependents"),
            ("^%\\", "ClearArrows"),
            ("^%=", "ZoomIn"),
            ("^%-", "ZoomOut"),
            ("^%+,", "GoToMin"),
            ("^%+.", "GoToMax"),
            ("%+=", "ExpandAllRows"),
            ("^%+=", "ExpandAllColumns"),
            ("%+-", "CollapseAllRows"),
            ("^%+-", "CollapseAllColumns"),
            ("^%+{UP}", "TopBorder"),
            ("%+{PGUP}", "RowHeightCycle"),
            ("^%+{INSERT}", "InsertColumn"),
            ("^{F2}", "AnchorFormulaCycle"),
            ("%{F12}", "QuickSaveAs"),
            ("^%{HOME}", "FirstSheet"),
            ("^+y", "BinaryCycle"),
        };

        private static string CoverageMacro(string name) => "EmtCov" + name;

        /// <summary>All (key, macro) pairs we bind: fixed commands + coverage set.</summary>
        private static IEnumerable<(string Key, string Macro)> AllBindings()
        {
            foreach (var b in Keys) yield return b;
            foreach (var (key, name) in Coverage) yield return (key, CoverageMacro(name));
        }

        /// <summary>Registers one log-only macro per coverage key (must run in AutoOpen).</summary>
        private static void RegisterCoverageMacros()
        {
            var delegates = new List<Delegate>();
            var attrs = new List<object>();
            var argAttrs = new List<List<object>>();
            foreach (var (key, name) in Coverage)
            {
                string k = key, n = name;
                delegates.Add(new Action(() => Commands.Run("Coverage." + n, k, () => new { key = k, name = n })));
                attrs.Add(new ExcelCommandAttribute { Name = CoverageMacro(n) });
                argAttrs.Add(new List<object>());
            }
            try
            {
                ExcelIntegration.RegisterDelegates(delegates, attrs, argAttrs);
                Log.Write("coverageMacros", new { ok = true, count = delegates.Count });
            }
            catch (Exception ex)
            {
                Log.Write("coverageMacros", new { ok = false, error = ex.ToString() });
            }
        }

        /// <summary>OnKey string -> "xlcOnKey" | "COM OnKey" | "failed".</summary>
        internal static readonly Dictionary<string, string> Methods = new Dictionary<string, string>();

        internal const string BuildStamp = "K2c (2026-09-28)";

        public void AutoOpen()
        {
            var sw = Stopwatch.StartNew();
            try
            {
                dynamic app = Xl.App;
                string version = null, build = null;
                try { version = Convert.ToString(app.Version); build = Convert.ToString(app.Build); } catch { }

                RegisterCoverageMacros();
                RegisterAll("AutoOpen");
                K2b.Install();
                sw.Stop();
                Log.Write("autoOpen", new
                {
                    build = BuildStamp,
                    excelVersion = version,
                    excelBuild = build,
                    bitness = Environment.Is64BitProcess ? 64 : 32,
                    xllPath = ExcelDnaUtil.XllPath,
                    loadMs = Xl.Ms(sw),
                });
                Xl.Status($"EMT spike {BuildStamp} loaded ({Xl.Ms(sw):0.0} ms). Log: {Log.FilePath}");

                // Pre-create the hidden add-in workbook outside any K2 test (creation itself may clear undo).
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try { K2.EnsureHiddenWorkbook("AutoOpen"); }
                    catch (Exception ex) { Log.Error("AutoOpen.hiddenWb", ex); }
                });
            }
            catch (Exception ex)
            {
                Log.Error("AutoOpen", ex);
            }
        }

        public void AutoClose()
        {
            foreach (var (key, macro) in AllBindings())
            {
                try
                {
                    XlCall.Excel(XlCall.xlcOnKey, key); // no macro => restore Excel default
                    Log.Write("unregister", new { key, macro, method = "xlcOnKey", ok = true });
                }
                catch (Exception ex)
                {
                    try
                    {
                        Xl.App.OnKey(key);
                        Log.Write("unregister", new { key, macro, method = "COM OnKey", ok = true, xlcError = ex.Message });
                    }
                    catch (Exception ex2)
                    {
                        Log.Write("unregister", new { key, macro, ok = false, xlcError = ex.Message, comError = ex2.Message });
                    }
                }
            }
            KeyHook.Uninstall("AutoClose");
            K2b.Uninstall();
            K2.CloseHiddenWorkbook();
            Log.Write("autoClose");
        }

        internal static void RegisterAll(string reason)
        {
            foreach (var (key, macro) in AllBindings())
            {
                var sw = Stopwatch.StartNew();
                string xlcError = null, comError = null, method;
                object result = null;
                try
                {
                    result = XlCall.Excel(XlCall.xlcOnKey, key, macro);
                    // The C API may "succeed" but return FALSE or an error value for a bad key string.
                    if (result is ExcelError || (result is bool ok && !ok))
                        throw new InvalidOperationException($"xlcOnKey returned {result}");
                    method = "xlcOnKey";
                }
                catch (Exception ex)
                {
                    xlcError = $"{ex.GetType().Name}: {ex.Message}";
                    try
                    {
                        Xl.App.OnKey(key, macro);
                        method = "COM OnKey";
                    }
                    catch (Exception ex2)
                    {
                        comError = $"{ex2.GetType().Name}: {ex2.Message}";
                        method = "failed";
                    }
                }
                Methods[key] = method;
                Log.Write("register", new
                {
                    reason, key, macro, method,
                    result = result?.ToString(),
                    xlcError, comError,
                    ms = Xl.Ms(sw),
                });
            }
        }

        internal static string MethodFor(string key) =>
            key != null && Methods.TryGetValue(key, out var m) ? m : "ribbon/none";
    }
}
