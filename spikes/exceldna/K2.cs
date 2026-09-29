using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace EmtSpike
{
    /// <summary>
    /// K2 variants: does any way of applying formatting keep Excel's native undo stack?
    /// Every variant logs an undo snapshot before and after, plus a queued-macro snapshot and one ~500 ms later.
    /// </summary>
    internal static class K2
    {
        private const int XlPasteFormats = -4122;
        private static dynamic _hiddenWb;

        public static object Control() => Variant("K2.1", () =>
        {
            dynamic font = Xl.App.Selection.Font;
            object b = font.Bold;
            bool bold = b is bool v && v; // DBNull when mixed
            font.Bold = !bold;
            return new { boldWas = b?.ToString(), set = !bold };
        });

        public static object MsoBoldDirect() => MsoBold("K2.2");

        public static object MsoBold(string variant) => Variant(variant, () =>
        {
            Xl.App.CommandBars.ExecuteMso("Bold");
            return "ExecuteMso(Bold) direct";
        });

        public static object MsoBoldDeferred() => Variant("K2.3", () =>
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    UndoProbe.Snapshot("deferred:beforeExecute", "K2.3");
                    Xl.App.CommandBars.ExecuteMso("Bold");
                    UndoProbe.Snapshot("deferred:afterExecute", "K2.3");
                }
                catch (Exception ex) { Log.Error("K2.3 deferred", ex); }
            });
            return "ExecuteMso(Bold) queued via QueueAsMacro";
        });

        public static object MsoPercentStyle() => Variant("K2.4", () =>
        {
            try
            {
                Xl.App.CommandBars.ExecuteMso("PercentStyle");
                return "ExecuteMso(PercentStyle) ok";
            }
            catch (Exception ex)
            {
                Log.Write("k2", new { variant = "K2.4", note = "PercentStyle failed (invalid idMso?)", error = ex.Message });
                return "PercentStyle failed: " + ex.Message;
            }
        });

        public static object CopyPasteSpecial() => Variant("K2.5", () =>
        {
            dynamic app = Xl.App;
            dynamic src = SourceCell();
            src.Copy();
            UndoProbe.Snapshot("afterCopy", "K2.5");
            app.Selection.PasteSpecial(XlPasteFormats);
            UndoProbe.Snapshot("afterPasteSpecial", "K2.5");
            app.CutCopyMode = false;
            return "Copy + PasteSpecial(xlPasteFormats) + CutCopyMode=false";
        });

        public static object CopyPasteMso() => CopyPasteMso("K2.6");

        public static object CopyPasteMso(string variant) => Variant(variant, () =>
        {
            dynamic app = Xl.App;
            dynamic src = SourceCell();
            src.Copy();
            UndoProbe.Snapshot("afterCopy", variant);

            var probes = new List<object>();
            string used = null;
            foreach (string id in new[] { "PasteFormatting", "PasteFormats" })
            {
                object enabled = null; string error = null;
                try { enabled = app.CommandBars.GetEnabledMso(id); }
                catch (Exception ex) { error = ex.Message; }
                probes.Add(new { id, exists = error == null, enabled, error });
            }
            foreach (string id in new[] { "PasteFormatting", "PasteFormats" })
            {
                try { app.CommandBars.ExecuteMso(id); used = id; break; }
                catch (Exception ex) { probes.Add(new { id, executeError = ex.Message }); }
            }
            UndoProbe.Snapshot("afterPasteMso", variant);
            app.CutCopyMode = false;
            return new { used, probes };
        });

        public static object XlmFormatNumber() => Variant("K2.7", () =>
        {
            object r = XlCall.Excel(XlCall.xlcFormatNumber, "0.0%");
            return "xlcFormatNumber(0.0%) returned " + r;
        });

        /// <summary>Before snapshot, action, after snapshot, then queued and ~500 ms snapshots.</summary>
        private static object Variant(string variant, Func<object> action)
        {
            UndoProbe.Snapshot("before", variant);
            object result = action();
            UndoProbe.Snapshot("after", variant);
            ExcelAsyncUtil.QueueAsMacro(() => UndoProbe.Snapshot("afterQueued", variant));
            Task.Delay(500).ContinueWith(_ =>
                ExcelAsyncUtil.QueueAsMacro(() => UndoProbe.Snapshot("after500ms", variant)));
            return result;
        }

        // ---- K2c (called from the K2b hook, i.e. outside macro context) ----------------------

        /// <summary>K2.11: plain COM write of the font color (red/blue toggle). Does it wipe history? Is it undoable?</summary>
        public static object ComFontColorNoMacro() => Variant("K2.11", () =>
        {
            dynamic font = Xl.App.Selection.Font;
            double cur = Convert.ToDouble(font.Color);
            double next = cur == 255 ? 16711680 : 255; // BGR: red (255) <-> blue (0xFF0000)
            font.Color = next;
            return new { was = cur, set = next };
        });

        /// <summary>
        /// K2.12: make the hidden template cell a clone of the active cell's formats via COM (hidden workbook only),
        /// change only its number format to the next of two codes, then copy it and ExecuteMso("PasteFormatting").
        /// Question: does the COM write to the HIDDEN workbook wipe the user's undo history?
        /// </summary>
        public static object CloneTemplatePaste() => Variant("K2.12", () =>
        {
            dynamic app = Xl.App;
            dynamic target = app.ActiveCell;
            dynamic tpl = SourceCell();
            string curFmt = Convert.ToString(target.NumberFormat);
            string nextFmt = curFmt == "0.00%" ? "#,##0.0" : "0.00%";

            tpl.Font.Bold = target.Font.Bold;
            tpl.Font.Italic = target.Font.Italic;
            tpl.Font.Color = target.Font.Color;
            object pattern = target.Interior.Pattern;
            if (Convert.ToInt32(pattern) == -4142) tpl.Interior.Pattern = -4142; // xlNone
            else tpl.Interior.Color = target.Interior.Color;
            tpl.HorizontalAlignment = target.HorizontalAlignment;
            tpl.NumberFormat = nextFmt;
            UndoProbe.Snapshot("afterTemplateWrite", "K2.12");

            tpl.Copy();
            UndoProbe.Snapshot("afterCopy", "K2.12");
            app.CommandBars.ExecuteMso("PasteFormatting");
            UndoProbe.Snapshot("afterPasteMso", "K2.12");
            app.CutCopyMode = false;
            return new { curFmt, nextFmt, targetBold = Convert.ToString(target.Font.Bold), targetColor = Convert.ToString(target.Font.Color), pattern = Convert.ToString(pattern) };
        });

        /// <summary>A1 of the hidden add-in workbook, pre-formatted as 0.0% so the test itself writes nothing there.</summary>
        private static dynamic SourceCell()
        {
            EnsureHiddenWorkbook("lazy");
            return _hiddenWb.Worksheets[1].Range["A1"];
        }

        public static void EnsureHiddenWorkbook(string reason)
        {
            try
            {
                if (_hiddenWb != null) { string _ = _hiddenWb.Name; return; } // still alive?
            }
            catch { _hiddenWb = null; }

            UndoProbe.Snapshot("hiddenWb:beforeCreate", reason);
            dynamic app = Xl.App;
            dynamic wb = app.Workbooks.Add();
            dynamic a1 = wb.Worksheets[1].Range["A1"];
            a1.Value2 = 0.5;
            a1.NumberFormat = "0.0%";
            wb.Windows[1].Visible = false;
            wb.Saved = true;
            _hiddenWb = wb;
            UndoProbe.Snapshot("hiddenWb:afterCreate", reason);
            Log.Write("hiddenWb", new { reason, name = Convert.ToString(wb.Name) });
        }

        public static void CloseHiddenWorkbook()
        {
            try { _hiddenWb?.Close(false); } catch (Exception ex) { Log.Error("CloseHiddenWorkbook", ex); }
            _hiddenWb = null;
        }
    }
}
