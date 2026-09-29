using System;
using ExcelModelingToolkit.Core.Formatting;
using ExcelModelingToolkit.Core.Undo;

namespace ExcelModelingToolkit.AddIn;

/// <summary>
/// Reads and writes the format properties the cycles change, over late-bound COM. Shared by the cycles and by
/// undo. Main thread only. Any write clears Excel's undo history (spike K2c).
/// </summary>
/// <remarks>
/// Colors are read and written as RGB (<c>Font.Color</c>, <c>Interior.Color</c>). A theme color's link to the
/// workbook theme (<c>ThemeColor</c>, <c>TintAndShade</c>) is therefore not preserved when undo puts a color back:
/// the cell gets the same RGB value as a plain color. Accepted for v1.
/// </remarks>
internal static class CellFormats
{
    /// <summary><c>xlNone</c>: <c>Interior.Pattern</c> of a cell with no fill.</summary>
    private const int XlNone = -4142;

    /// <summary><c>xlSolid</c>: <c>Interior.Pattern</c> of a plain color fill.</summary>
    private const int XlSolid = 1;

    /// <summary><c>xlColorIndexAutomatic</c>: <c>Font.ColorIndex</c> of a font set to Automatic.</summary>
    private const int XlColorIndexAutomatic = -4105;

    /// <summary>
    /// The value of <paramref name="kind"/> that every cell of <paramref name="range"/> holds, for a cycle; unknown
    /// if they differ (Excel returns null for a mixed range) or it cannot be interpreted. An Automatic font reads as
    /// its color (usually black).
    /// </summary>
    public static CycleValue Read(object range, CycleKind kind) => Read(range, kind, forUndo: false);

    /// <summary>
    /// Like <see cref="Read(object, CycleKind)"/>, but exact enough for undo to put the value back: an Automatic
    /// font reads as <see cref="CycleValue.Automatic"/> (a range mixing it with an explicit black is mixed), and a
    /// pattern or gradient fill cannot be restored.
    /// </summary>
    /// <exception cref="UnrestorableFormatException">Every cell has a pattern or gradient fill.</exception>
    public static CycleValue ReadForUndo(object range, CycleKind kind) => Read(range, kind, forUndo: true);

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
            case CycleKind.FontColor when value.IsAutomatic:
                cells.Font.ColorIndex = XlColorIndexAutomatic;
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

    private static CycleValue Read(object range, CycleKind kind, bool forUndo)
    {
        dynamic cells = range;
        switch (kind)
        {
            case CycleKind.NumberFormat:
                object format = cells.NumberFormat;
                return format is string code ? CycleValue.FromNumberFormat(code) : CycleValue.Unknown;
            case CycleKind.FontColor:
                dynamic font = cells.Font;
                if (forUndo)
                {
                    // Font.Color reports Automatic as black; ColorIndex tells them apart (null when mixed).
                    object colorIndex = font.ColorIndex;
                    var index = ToInt(colorIndex);
                    if (index is null)
                    {
                        return CycleValue.Unknown;
                    }

                    if (index == XlColorIndexAutomatic)
                    {
                        return CycleValue.Automatic;
                    }
                }

                object fontColor = font.Color;
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

                if (forUndo && patternValue != XlSolid)
                {
                    // A pattern or gradient: undo only writes a plain color back, which would lose it.
                    throw new UnrestorableFormatException("pattern or gradient fill");
                }

                object fillColor = interior.Color;
                return ToColor(fillColor);
            default:
                return CycleValue.Unknown;
        }
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
