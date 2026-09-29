using System.Collections.Generic;
using ExcelModelingToolkit.Core.Formatting;

namespace ExcelModelingToolkit.Core.Settings;

/// <summary>
/// The factory defaults: the owner's installed Macabacus configuration (docs/research/06) on the Macabacus keys
/// (docs/PLAN.md section 4.2). The dash inside the quoted zero sections is an en dash (U+2013).
/// </summary>
/// <remarks>
/// The Date, Currency, Percent and Multiple cycles are <b>provisional</b>
/// (<see cref="CycleDefinition.Provisional"/>): Macabacus-style placeholders until the owner supplies the real
/// lists (open-questions A1). The General Number, Font Color, Fill Color and Blue-Black cycles are the owner's
/// observed values.
/// </remarks>
public static class DefaultSettings
{
    /// <summary>Creates the default settings.</summary>
    public static ToolkitSettings Create() => new ToolkitSettings(CreateCycles(), CreateKeymap());

    private static List<CycleDefinition> CreateCycles() => new List<CycleDefinition>
    {
        // Observed (research/06): General Number Cycle.
        new CycleDefinition(ActionIds.NumberCycle, "Number", CycleKind.NumberFormat, new CycleItem[]
        {
            new NumberFormatItem("Comma 0 Dec Lg Align", "_(#,##0_)_%;(#,##0)_%;_(\"–\"_)_%;_(@_)_%"),
            new NumberFormatItem("Comma 1 Dec Lg Align", "_(#,##0.0_)_%;(#,##0.0)_%;_(\"–\"_)_%;_(@_)_%"),
            new NumberFormatItem("Comma 2 Dec Lg Align", "_(#,##0.00_)_%;(#,##0.00)_%;_(\"–\"_)_%;_(@_)_%"),
            new NumberFormatItem("Comma 0 Dec No Align", "#,##0;(#,##0);\"–\";@"),
        }),

        // Provisional: the owner's Date list has not been captured yet.
        new CycleDefinition(ActionIds.DateCycle, "Date", CycleKind.NumberFormat, new CycleItem[]
        {
            new NumberFormatItem("Year Actual", "yyyy\"A\""),
            new NumberFormatItem("Year Estimate", "yyyy\"E\""),
            new NumberFormatItem("Month-Day-Year", "mm-dd-yyyy"),
            new NumberFormatItem("Year-Month-Day", "yyyy-mm-dd"),
        }, provisional: true),

        // Provisional: the owner's Currency list has not been captured yet.
        new CycleDefinition(ActionIds.CurrencyCycle, "Currency", CycleKind.NumberFormat, new CycleItem[]
        {
            new NumberFormatItem("Currency 0 Dec Lg Align", "_([$$]#,##0_)_%;([$$]#,##0)_%;_(\"–\"_)_%;_(@_)_%"),
            new NumberFormatItem("Currency 1 Dec Lg Align", "_([$$]#,##0.0_)_%;([$$]#,##0.0)_%;_(\"–\"_)_%;_(@_)_%"),
            new NumberFormatItem("Currency 2 Dec Lg Align", "_([$$]#,##0.00_)_%;([$$]#,##0.00)_%;_(\"–\"_)_%;_(@_)_%"),
        }, provisional: true),

        // Provisional: the owner's Percent list has not been captured yet.
        new CycleDefinition(ActionIds.PercentCycle, "Percent", CycleKind.NumberFormat, new CycleItem[]
        {
            new NumberFormatItem("Percent 1 Dec", "_(0.0%_);(0.0%);_(\"–\"_)_%;_(@_)_%"),
            new NumberFormatItem("Percent 0 Dec", "_(0%_);(0%);_(\"–\"_)_%;_(@_)_%"),
            new NumberFormatItem("Percent 2 Dec", "_(0.00%_);(0.00%);_(\"–\"_)_%;_(@_)_%"),
        }, provisional: true),

        // Provisional: the owner's Multiple list has not been captured yet.
        new CycleDefinition(ActionIds.MultipleCycle, "Multiple", CycleKind.NumberFormat, new CycleItem[]
        {
            new NumberFormatItem("Multiple 1 Dec", "_(0.0\"x\"_)_%;(0.0\"x\")_%;_(\"–\"_)_%;_(@_)_%"),
            new NumberFormatItem("Multiple 2 Dec", "_(0.00\"x\"_)_%;(0.00\"x\")_%;_(\"–\"_)_%;_(@_)_%"),
            new NumberFormatItem("Multiple 0 Dec", "_(0\"x\"_)_%;(0\"x\")_%;_(\"–\"_)_%;_(@_)_%"),
        }, provisional: true),

        // Observed (research/06): Font Colors. Starts at blue; black is last.
        new CycleDefinition(ActionIds.FontColorCycle, "Font Color", CycleKind.FontColor, new CycleItem[]
        {
            new ColorItem("Blue", OleColor.FromRgb(0, 0, 255)),
            new ColorItem("Green", OleColor.FromRgb(0, 128, 0)),
            new ColorItem("Purple", OleColor.FromRgb(128, 0, 128)),
            new ColorItem("Red", OleColor.FromRgb(255, 0, 0)),
            new ColorItem("White", OleColor.FromRgb(255, 255, 255)),
            new ColorItem("Black", OleColor.FromRgb(0, 0, 0)),
        }),

        // Observed (research/06): Fill Colors, ending with No fill.
        new CycleDefinition(ActionIds.FillColorCycle, "Fill Color", CycleKind.FillColor, new CycleItem[]
        {
            new ColorItem("Light Blue", OleColor.FromRgb(201, 218, 248)),
            new ColorItem("Light Cyan", OleColor.FromRgb(210, 242, 255)),
            new ColorItem("Light Pink", OleColor.FromRgb(244, 204, 204)),
            new ColorItem("Peach", OleColor.FromRgb(252, 229, 205)),
            new ColorItem("Navy", OleColor.FromRgb(28, 69, 135)),
            new ColorItem("No Fill", OleColor.NoFill),
        }),

        // Observed (research/06): AutoColor Inputs blue and the default font color black.
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
        [ActionIds.FontColorCycle] = "Ctrl+'",
        [ActionIds.FillColorCycle] = "Ctrl+Shift+K",
        [ActionIds.BlueBlackToggle] = "Ctrl+;",
        [ActionIds.About] = "Ctrl+Alt+Shift+F12",
    };
}
