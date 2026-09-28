using System;

namespace EmtSpike
{
    /// <summary>K3 cycles with Macabacus-like defaults. Read the active cell, write the whole Selection in one COM call.</summary>
    internal static class Cycles
    {
        public static readonly string[] Number =
        {
            "_(#,##0_)_%;(#,##0)_%;_(\"–\"_)_%;_(@_)_%",
            "_(#,##0.0_)_%;(#,##0.0)_%;_(\"–\"_)_%;_(@_)_%",
            "_(#,##0.00_)_%;(#,##0.00)_%;_(\"–\"_)_%;_(@_)_%",
            "#,##0;(#,##0);\"–\";@",
        };

        public static readonly string[] Percent =
        {
            "0.0%_);(0.0%);\"–\"_)",
            "0%_);(0%);\"–\"_)",
            "0.00%_);(0.00%);\"–\"_)",
        };

        public static readonly string[] Multiple =
        {
            "0.0\"x\"_);(0.0\"x\");\"–\"_)",
            "0.00\"x\"_);(0.00\"x\");\"–\"_)",
        };

        public static readonly string[] Currency =
        {
            "_([$$]#,##0_)_%;([$$]#,##0)_%;_(\"–\"_)_%;_(@_)_%",
            "_([$$]#,##0.00_)_%;([$$]#,##0.00)_%;_(\"–\"_)_%;_(@_)_%",
        };

        public static readonly string[] Date =
        {
            "yyyy\"A\"",
            "mm-dd-yyyy",
            "yyyy-mm-dd",
            "mmm-yy",
        };

        /// <summary>Excel COM colors are BGR: R + G*256 + B*65536.</summary>
        private static int Rgb(int r, int g, int b) => r + g * 256 + b * 65536;

        private static readonly int Black = Rgb(0, 0, 0), Blue = Rgb(0, 0, 255);
        // Owner's real Macabacus orders.
        private static readonly int[] FontColors = { Blue, Rgb(0, 128, 0), Rgb(128, 0, 128), Rgb(255, 0, 0), Rgb(255, 255, 255), Black };
        private static readonly string[] FontNames = { "blue", "green", "purple", "red", "white", "black" };

        private const int NoFill = -1;
        private const int XlNone = -4142;
        private static readonly int[] Fills = { Rgb(201, 218, 248), Rgb(210, 242, 255), Rgb(244, 204, 204), Rgb(252, 229, 205), Rgb(28, 69, 135), NoFill };
        private static readonly string[] FillNames = { "rgb(201,218,248)", "rgb(210,242,255)", "rgb(244,204,204)", "rgb(252,229,205)", "rgb(28,69,135)", "no fill" };

        /// <summary>Index of the next item: current match -> next (wrapping); no match -> first.</summary>
        private static int Next(int currentIndex, int length) => currentIndex < 0 ? 0 : (currentIndex + 1) % length;

        public static object Format(string[] list)
        {
            dynamic app = Xl.App;
            string current = Convert.ToString(app.ActiveCell.NumberFormat);
            int i = Array.IndexOf(list, current);
            string next = list[Next(i, list.Length)];
            app.Selection.NumberFormat = next;
            return new { current, matchedIndex = i, applied = next };
        }

        private static int ActiveFontColor()
        {
            try { return Convert.ToInt32(Convert.ToDouble(Xl.App.ActiveCell.Font.Color)); } catch { return -2; }
        }

        public static object FontColor()
        {
            int current = ActiveFontColor();
            int i = Array.IndexOf(FontColors, current);
            int n = Next(i, FontColors.Length);
            Xl.App.Selection.Font.Color = FontColors[n];
            return new { current, matchedIndex = i, applied = FontNames[n] };
        }

        public static object BlueBlack()
        {
            int current = ActiveFontColor();
            int next = current == Blue ? Black : Blue;
            Xl.App.Selection.Font.Color = next;
            return new { current, applied = next == Blue ? "blue" : "black" };
        }

        public static object Fill()
        {
            dynamic app = Xl.App;
            int current;
            try
            {
                dynamic interior = app.ActiveCell.Interior;
                current = Convert.ToInt32(interior.Pattern) == XlNone ? NoFill : Convert.ToInt32(Convert.ToDouble(interior.Color));
            }
            catch { current = -2; }

            int i = Array.IndexOf(Fills, current);
            int n = Next(i, Fills.Length);
            if (Fills[n] == NoFill) app.Selection.Interior.Pattern = XlNone;
            else app.Selection.Interior.Color = Fills[n];
            return new { current, matchedIndex = i, applied = FillNames[n] };
        }
    }
}
