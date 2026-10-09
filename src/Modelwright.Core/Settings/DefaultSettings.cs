using System.Collections.Generic;
using Modelwright.Core.Formatting;

namespace Modelwright.Core.Settings;

/// <summary>
/// The factory defaults: Macabacus's factory settings (MacabacusSettings 9.9.5: <c>NumberFormatCycles</c>,
/// <c>ColorCycles</c>, <c>AutoColors</c>, <c>DefaultColors</c>) on the Macabacus keys (<c>ExcelShortcuts</c>;
/// docs/PLAN.md section 4.2). Names and codes are Macabacus's, character for character; the dash inside the quoted
/// zero sections is an en dash (U+2013).
/// </summary>
/// <remarks>
/// No default cycle is provisional (<see cref="CycleDefinition.Provisional"/>): these lists replace the Date,
/// Currency, Percent and Multiple placeholders of earlier builds (see <see cref="ToolkitSettings.FromJson"/> for
/// how a settings file that still has them is brought up to date).
/// </remarks>
public static class DefaultSettings
{
    /// <summary>Creates the default settings.</summary>
    public static ToolkitSettings Create() => new ToolkitSettings(CreateCycles(), CreateKeymap());

    private static List<CycleDefinition> CreateCycles() => new List<CycleDefinition>
    {
        // Factory GeneralNumberCycle.
        new CycleDefinition(ActionIds.NumberCycle, "Number", CycleKind.NumberFormat, new CycleItem[]
        {
            new NumberFormatItem("Comma 0 Dec Lg Align", "_(#,##0_)_%;(#,##0)_%;_(\"–\"_)_%;_(@_)_%"),
            new NumberFormatItem("Comma 1 Dec Lg Align", "_(#,##0.0_)_%;(#,##0.0)_%;_(\"–\"_)_%;_(@_)_%"),
            new NumberFormatItem("Comma 2 Dec Lg Align", "_(#,##0.00_)_%;(#,##0.00)_%;_(\"–\"_)_%;_(@_)_%"),
            new NumberFormatItem("Comma 0 Dec No Align", "#,##0;(#,##0);\"–\";@"),
        }),

        // Factory DateCycle.
        new CycleDefinition(ActionIds.DateCycle, "Date", CycleKind.NumberFormat, new CycleItem[]
        {
            new NumberFormatItem("m/d/yyyy", "m/d/yyyy;@"),
            new NumberFormatItem("Date Text Long", "mmmm d, yyyy;@"),
            new NumberFormatItem("Date Actual Year", "0000\\A"),
            new NumberFormatItem("Date Estimated Year", "0000\\E"),
        }),

        // Factory LocalCurrencyCycle.
        new CycleDefinition(ActionIds.CurrencyCycle, "Currency", CycleKind.NumberFormat, new CycleItem[]
        {
            new NumberFormatItem("USD 0 Dec Lg Align", "_([$$]#,##0_)_%;([$$]#,##0)_%;_(\"–\"_)_%;_(@_)_%"),
            new NumberFormatItem("USD 1 Dec Lg Align", "_([$$]#,##0.0_)_%;([$$]#,##0.0)_%;_(\"–\"_)_%;_(@_)_%"),
            new NumberFormatItem("USD 2 Dec Lg Align", "_([$$]#,##0.00_)_%;([$$]#,##0.00)_%;_(\"–\"_)_%;_(@_)_%"),
            new NumberFormatItem("USD 0 Dec No Align", "[$$]#,##0;([$$]#,##0);\"–\";@"),
            new NumberFormatItem("EUR 0 Dec Lg Align", "_([$€-2]#,##0_)_%;([$€-2]#,##0)_%;_(\"–\"_)_%;_(@_)_%"),
            new NumberFormatItem("GBP 0 Dec Lg Align", "_([$£-809]#,##0_)_%;([$£-809]#,##0)_%;_(\"–\"_)_%;_(@_)_%"),
            new NumberFormatItem("YEN 0 Dec Lg Align", "_([$¥-2]#,##0_)_%;([$¥-2]#,##0)_%;_(\"–\"_)_%;_(@_)_%"),
        }),

        // Factory PercentCycle.
        new CycleDefinition(ActionIds.PercentCycle, "Percent", CycleKind.NumberFormat, new CycleItem[]
        {
            new NumberFormatItem("Percent Aligned Neg Pct", "_(#,##0.0%_);(#,##0.0%);_(\"–\"_)_%;_(@_)_%"),
            new NumberFormatItem("Percent Unaligned", "#,##0.0%;(#,##0.0%);\"–\";@"),
            new NumberFormatItem("Hard Percent Aligned Neg Pct", "_(#,##0.0\"%\"_);(#,##0.0\"%\");_(\"–\"_)_%;_(@_)_%"),
            new NumberFormatItem("Hard Percent Unaligned", "#,##0.0\"%\";(#,##0.0\"%\");\"–\";@"),
            new NumberFormatItem("SOFR +", "\"S\"+0_)_%;\"S\"-0_)_%;\"S\"+0_)_%"),
            new NumberFormatItem("LIBOR +", "L+0_)_%;L-0_)_%;L+0_)_%"),
        }),

        // Factory MultipleCycle: the x is unquoted and _' pads by the width of an apostrophe, as Macabacus has it.
        new CycleDefinition(ActionIds.MultipleCycle, "Multiple", CycleKind.NumberFormat, new CycleItem[]
        {
            new NumberFormatItem("Mult 1 Decimal Aligned Neg Pct", "_(0.0x_)_)_';_((0.0x)_'_';_(\"–\"_)_%;_(@_)_%"),
            new NumberFormatItem("Mult 2 Decimal Aligned Neg Pct", "_(0.00x_)_)_';_((0.00x)_'_';_(\"–\"_)_%;_(@_)_%"),
            new NumberFormatItem("Mult 1 Decimal Unaligned", "0.0x;(0.0x);\"–\""),
            new NumberFormatItem("Mult 2 Decimal Unaligned", "0.00x;(0.00x);\"–\""),
        }),

        // Factory BinaryCycle: a positive number shows the first word, zero the second, a negative or text ERROR.
        new CycleDefinition(ActionIds.BinaryCycle, "Binary", CycleKind.NumberFormat, new CycleItem[]
        {
            new NumberFormatItem("Yes/No", "\"Yes\";\"ERROR\";\"No\";\"ERROR\""),
            new NumberFormatItem("Y/N", "\"Y\";\"ERROR\";\"N\";\"ERROR\""),
            new NumberFormatItem("On/Off", "\"On\";\"ERROR\";\"Off\";\"ERROR\""),
            new NumberFormatItem("True/False", "\"True\";\"ERROR\";\"False\";\"ERROR\""),
        }),

        // Factory RatioCycle.
        new CycleDefinition(ActionIds.RatioCycle, "Ratio", CycleKind.NumberFormat, new CycleItem[]
        {
            new NumberFormatItem("Exchange Ratio", "0.0\\:1_);(0.0)\\:1_);0.0\\:1_);@_)"),
            new NumberFormatItem("Fraction 1", "# ?/?"),
            new NumberFormatItem("Fraction 2", "# ??/??"),
            new NumberFormatItem("Fraction 3", "# ???/???"),
            new NumberFormatItem("Halves", "# ?/2"),
            new NumberFormatItem("Thirds", "# ?/3"),
        }),

        // Factory FontColorCycle. Starts at blue; black is last.
        new CycleDefinition(ActionIds.FontColorCycle, "Font Color", CycleKind.FontColor, new CycleItem[]
        {
            new ColorItem("Blue", OleColor.FromRgb(0, 0, 255)),
            new ColorItem("Green", OleColor.FromRgb(0, 128, 0)),
            new ColorItem("Purple", OleColor.FromRgb(128, 0, 128)),
            new ColorItem("Red", OleColor.FromRgb(255, 0, 0)),
            new ColorItem("White", OleColor.FromRgb(255, 255, 255)),
            new ColorItem("Black", OleColor.FromRgb(0, 0, 0)),
        }),

        // Factory FillColorCycle, ending with No fill (an empty Color).
        new CycleDefinition(ActionIds.FillColorCycle, "Fill Color", CycleKind.FillColor, new CycleItem[]
        {
            new ColorItem("Light Blue", OleColor.FromRgb(201, 218, 248)),
            new ColorItem("Light Cyan", OleColor.FromRgb(210, 242, 255)),
            new ColorItem("Light Pink", OleColor.FromRgb(244, 204, 204)),
            new ColorItem("Peach", OleColor.FromRgb(252, 229, 205)),
            new ColorItem("Navy", OleColor.FromRgb(28, 69, 135)),
            new ColorItem("No Fill", OleColor.NoFill),
        }),

        // Factory AutoColors ColorInputs (blue) and DefaultColors FontColor (black).
        new CycleDefinition(ActionIds.BlueBlackToggle, "Blue/Black", CycleKind.FontColor, new CycleItem[]
        {
            new ColorItem("Blue", OleColor.FromRgb(0, 0, 255)),
            new ColorItem("Black", OleColor.FromRgb(0, 0, 0)),
        }),
    };

    private static Dictionary<string, string> CreateKeymap() => new Dictionary<string, string>
    {
        [ActionIds.NumberCycle] = "Ctrl+Shift+1",
        [ActionIds.DateCycle] = "Ctrl+Shift+2",
        [ActionIds.CurrencyCycle] = "Ctrl+Shift+4",
        [ActionIds.PercentCycle] = "Ctrl+Shift+5",
        [ActionIds.MultipleCycle] = "Ctrl+Shift+8",
        [ActionIds.BinaryCycle] = "Ctrl+Shift+Y",
        [ActionIds.RatioCycle] = "Alt+Shift+;",
        [ActionIds.FontColorCycle] = "Ctrl+'",
        [ActionIds.FillColorCycle] = "Ctrl+Shift+K",
        [ActionIds.BlueBlackToggle] = "Ctrl+;",
        [ActionIds.TraceIn] = "Ctrl+Shift+[",
        [ActionIds.LastAuditedCell] = "Ctrl+Shift+\\",
        [ActionIds.About] = "Ctrl+Alt+Shift+F12",
    };
}
