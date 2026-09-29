using System;
using ExcelModelingToolkit.Core.Formatting;

namespace ExcelModelingToolkit.AddIn;

/// <summary>
/// Reads and writes the format properties the cycles change, over late-bound COM. Shared by the cycles and by
/// undo. Main thread only. Any write clears Excel's undo history (spike K2c).
/// </summary>
internal static class CellFormats
{
    /// <summary><c>xlNone</c>: <c>Interior.Pattern</c> of a cell with no fill.</summary>
    private const int XlNone = -4142;

    /// <summary>
    /// The value of <paramref name="kind"/> that every cell of <paramref name="range"/> holds; unknown if they
    /// differ (Excel returns null for a mixed range) or it cannot be interpreted.
    /// </summary>
    public static CycleValue Read(object range, CycleKind kind)
    {
        dynamic cells = range;
        switch (kind)
        {
            case CycleKind.NumberFormat:
                object format = cells.NumberFormat;
                return format is string code ? CycleValue.FromNumberFormat(code) : CycleValue.Unknown;
            case CycleKind.FontColor:
                object fontColor = cells.Font.Color;
                return ToColor(fontColor);
            case CycleKind.FillColor:
                dynamic interior = cells.Interior;
                object pattern = interior.Pattern;
                var patternValue = ToInt(pattern);
                if (patternValue == XlNone)
                {
                    return CycleValue.FromColor(OleColor.NoFill);
                }

                if (patternValue is null)
                {
                    // Mixed patterns: some cells may have no fill, which reads as white in Interior.Color.
                    return CycleValue.Unknown;
                }

                object fillColor = interior.Color;
                return ToColor(fillColor);
            default:
                return CycleValue.Unknown;
        }
    }

    /// <summary>Writes <paramref name="value"/> to every cell of <paramref name="range"/> with one COM property write.</summary>
    /// <exception cref="InvalidOperationException"><paramref name="value"/> is not a value of <paramref name="kind"/>.</exception>
    public static void Write(object range, CycleKind kind, CycleValue value)
    {
        dynamic cells = range;
        switch (kind)
        {
            case CycleKind.NumberFormat when value.NumberFormat is string code:
                cells.NumberFormat = code;
                break;
            case CycleKind.FontColor when value.Color is OleColor color && !color.IsNoFill:
                cells.Font.Color = color.OleValue;
                break;
            case CycleKind.FillColor when value.Color is OleColor color && color.IsNoFill:
                cells.Interior.Pattern = XlNone;
                break;
            case CycleKind.FillColor when value.Color is OleColor color:
                cells.Interior.Color = color.OleValue;
                break;
            default:
                throw new InvalidOperationException($"\"{value}\" cannot be applied as a {kind}.");
        }
    }

    /// <summary>
    /// True if <paramref name="worksheet"/> is protected and its protection does not allow formatting cells, so
    /// every write would fail.
    /// </summary>
    public static bool FormattingIsProtected(object worksheet)
    {
        dynamic sheet = worksheet;
        object protectContents = sheet.ProtectContents;
        if (!(protectContents is bool isProtected && isProtected))
        {
            return false;
        }

        object allowFormatting = sheet.Protection.AllowFormattingCells;
        return !(allowFormatting is bool allowed && allowed);
    }

    /// <summary>An OLE color from a COM value (Excel returns a double); unknown for DBNull (mixed) or out of range.</summary>
    private static CycleValue ToColor(object? value)
    {
        var ole = ToInt(value);
        return ole is int v && v >= 0 && v <= OleColor.MaxOleValue
            ? CycleValue.FromColor(OleColor.FromOle(v))
            : CycleValue.Unknown;
    }

    private static int? ToInt(object? value) => value switch
    {
        int i => i,
        double d when d >= int.MinValue && d <= int.MaxValue && d == Math.Floor(d) => (int)d,
        _ => null,
    };
}
