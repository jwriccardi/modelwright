using System;
using System.IO;

namespace EmtSpike
{
    /// <summary>K4: creates EMT_Fixture_B.xlsx and EMT_Fixture_A.xlsx in %LOCALAPPDATA%\EmtSpike and leaves both open.</summary>
    internal static class Fixtures
    {
        private const int XlOpenXmlWorkbook = 51;
        private const int XlSheetHidden = 0;
        private const string BookA = "EMT_Fixture_A.xlsx", BookB = "EMT_Fixture_B.xlsx";

        public static object Create()
        {
            dynamic app = Xl.App;
            Directory.CreateDirectory(Log.Dir);
            CloseIfOpen(BookA);
            CloseIfOpen(BookB);

            bool alerts = app.DisplayAlerts;
            app.DisplayAlerts = false; // overwrite existing files silently
            try
            {
                // Book B: Data!C3:C5
                dynamic wbB = app.Workbooks.Add();
                dynamic data = wbB.Worksheets[1];
                data.Name = "Data";
                data.Range["C3"].Value2 = 10;
                data.Range["C4"].Value2 = 20;
                data.Range["C5"].Value2 = 30;
                wbB.SaveAs(Path.Combine(Log.Dir, BookB), XlOpenXmlWorkbook);

                // Book A: Sheet1, Sheet2, Hidden (create all sheets before writing the formula)
                dynamic wbA = app.Workbooks.Add();
                while (Convert.ToInt32(wbA.Worksheets.Count) < 3)
                    wbA.Worksheets.Add(Type.Missing, wbA.Worksheets[wbA.Worksheets.Count]);
                dynamic s1 = wbA.Worksheets[1], s2 = wbA.Worksheets[2], hidden = wbA.Worksheets[3];
                s1.Name = "Sheet1";
                s2.Name = "Sheet2";
                hidden.Name = "Hidden";

                s2.Range["B2"].Value2 = 5;
                s2.Range["B3"].Value2 = 7;
                for (int r = 5; r <= 10; r++) s2.Range["B" + r].Value2 = r - 4;
                s1.Range["D10"].Value2 = 100;
                hidden.Range["A1"].Value2 = 42;

                s1.Range["A1"].Formula =
                    "=Sheet2!B2+Sheet2!B3*2+'[EMT_Fixture_B.xlsx]Data'!C3+Sheet1!D10+SUM(Sheet2!B5:B10)+Hidden!A1";
                hidden.Visible = XlSheetHidden;

                wbA.SaveAs(Path.Combine(Log.Dir, BookA), XlOpenXmlWorkbook);

                wbA.Activate();
                s1.Activate();
                s1.Range["A1"].Select();
                return new { dir = Log.Dir, formula = Convert.ToString(s1.Range["A1"].Formula) };
            }
            finally
            {
                app.DisplayAlerts = alerts;
            }
        }

        private static void CloseIfOpen(string name)
        {
            try { Xl.App.Workbooks[name].Close(false); }
            catch { /* not open */ }
        }
    }
}
