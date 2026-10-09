using System;
using System.Globalization;
using System.Text;

namespace Modelwright.Core.Trace;

/// <summary>
/// The text of the Trace In tree's Value column and badges, from what Excel returns over COM: <c>Range.Text</c>
/// (the value as displayed), <c>Range.Value2</c> or <c>Worksheet.Evaluate</c> (a double, string, bool, error code,
/// array or nothing).
/// </summary>
public static class TraceValueText
{
    /// <summary>The longest value text shown; longer text is cut with an ellipsis.</summary>
    public const int MaxLength = 255;

    // A COM VT_ERROR from Excel is 0x800A0000 plus Excel's error number (2042 for #N/A).
    private const int ComErrorBase = unchecked((int)0x800A0000);

    /// <summary>
    /// The value as Excel displays it: <paramref name="text"/> (<c>Range.Text</c>, the cell's formatted value) when
    /// it is usable, else <paramref name="value"/> formatted by <see cref="FromValue"/>. Text is not usable when it
    /// is null, empty while the cell holds a value, or all <c>#</c> (the column is too narrow to show the number).
    /// </summary>
    /// <param name="text">The cell's <c>Range.Text</c>, or null if it could not be read.</param>
    /// <param name="value">The cell's <c>Range.Value2</c>.</param>
    /// <param name="culture">Formats numbers when the text is not usable.</param>
    public static string Display(string? text, object? value, IFormatProvider culture)
    {
        if (!string.IsNullOrEmpty(text) && !IsAllHashes(text!))
        {
            return SingleLine(text!);
        }

        return FromValue(value, culture);
    }

    /// <summary>
    /// A value from <c>Value2</c> or <c>Evaluate</c>: empty for nothing; a number in general format (up to 15
    /// significant digits); <c>TRUE</c> or <c>FALSE</c>; an error as Excel writes it (<c>#N/A</c>); text on one line,
    /// at most <see cref="MaxLength"/> characters; an array as its first element followed by its size
    /// (<c>1 {3x2}</c>).
    /// </summary>
    /// <param name="value">The value.</param>
    /// <param name="culture">Formats numbers and dates.</param>
    public static string FromValue(object? value, IFormatProvider culture)
    {
        switch (value)
        {
            case null:
                return string.Empty;
            case string text:
                return SingleLine(text);
            case bool flag:
                return flag ? "TRUE" : "FALSE";
            case double number:
                return number.ToString("G15", culture);
            case int code:
                return ErrorText(code) ?? code.ToString(culture);
            case DateTime date:
                return date.ToString(culture);
            case Array array:
                return FromArray(array, culture);
            case IFormattable formattable:
                return SingleLine(formattable.ToString(null, culture));
            default:
                return SingleLine(value.ToString() ?? string.Empty);
        }
    }

    /// <summary>
    /// A value <c>Worksheet.Evaluate</c> returned for a fragment of a formula (Trace In's evaluate mode): as
    /// <see cref="FromValue"/>, except that a number shows its digits grouped, as a cell formatted <c>#,##0.##</c>
    /// would (<c>15,120</c>, <c>-1,234.5</c>), still to 15 significant digits; a very large or small number stays in
    /// scientific notation.
    /// </summary>
    /// <param name="value">The value.</param>
    /// <param name="culture">Formats numbers (separators) and dates.</param>
    public static string FromEvaluated(object? value, IFormatProvider culture)
    {
        if (value is not double number || double.IsNaN(number) || double.IsInfinity(number))
        {
            return FromValue(value, culture);
        }

        var general = (number == 0 ? 0d : number).ToString("G15", CultureInfo.InvariantCulture); // no "-0"
        if (general.IndexOf('E') >= 0)
        {
            return number.ToString("G15", culture);
        }

        var point = general.IndexOf('.');
        var decimals = point < 0 ? 0 : general.Length - point - 1;
        return double.Parse(general, CultureInfo.InvariantCulture).ToString("N" + decimals.ToString(CultureInfo.InvariantCulture), culture);
    }

    /// <summary>
    /// Excel's text for an error value (<c>#DIV/0!</c>, <c>#N/A</c>, <c>#SPILL!</c>...), given as a COM error code
    /// (<c>-2146826246</c>, how <c>Value2</c> returns <c>#N/A</c>) or as Excel's error number (<c>2042</c>); null if
    /// <paramref name="code"/> is neither.
    /// </summary>
    public static string? ErrorText(int code)
    {
        var number = code >= 2000 && code < 2100 ? code : code - ComErrorBase;
        switch (number)
        {
            case 2000:
                return "#NULL!";
            case 2007:
                return "#DIV/0!";
            case 2015:
                return "#VALUE!";
            case 2023:
                return "#REF!";
            case 2029:
                return "#NAME?";
            case 2036:
                return "#NUM!";
            case 2042:
                return "#N/A";
            case 2043:
                return "#GETTING_DATA";
            case 2045:
                return "#SPILL!";
            case 2046:
                return "#CONNECT!";
            case 2047:
                return "#BLOCKED!";
            case 2048:
                return "#UNKNOWN!";
            case 2049:
                return "#FIELD!";
            case 2050:
                return "#CALC!";
            default:
                return null;
        }
    }

    /// <summary>
    /// The Value column of a range: its first cell's value, followed by the cell count when there is more than one
    /// cell (<c>12.5 (60 cells)</c>).
    /// </summary>
    /// <param name="firstCell">The first cell's value text.</param>
    /// <param name="cellCount">The number of cells in the range.</param>
    /// <param name="culture">Formats the count.</param>
    public static string ForRange(string firstCell, long cellCount, IFormatProvider culture)
    {
        if (firstCell is null)
        {
            throw new ArgumentNullException(nameof(firstCell));
        }

        if (cellCount <= 1)
        {
            return firstCell;
        }

        var count = cellCount.ToString("N0", culture) + " cells";
        return firstCell.Length == 0 ? "(" + count + ")" : firstCell + " (" + count + ")";
    }

    /// <summary>
    /// The badge for a target that cannot be seen, or null if it is visible: <c>hidden workbook</c>,
    /// <c>hidden sheet</c>, <c>hidden rows</c>, <c>hidden columns</c> or <c>hidden rows and columns</c> (the first
    /// that applies, rows and columns together). For rows and columns, null means some but not all are hidden
    /// (Excel's <c>Range.EntireRow.Hidden</c> is null for a mix), which counts as hidden.
    /// </summary>
    public static string? HiddenNote(bool workbookHidden, bool sheetHidden, bool? rowsHidden, bool? columnsHidden)
    {
        if (workbookHidden)
        {
            return "hidden workbook";
        }

        if (sheetHidden)
        {
            return "hidden sheet";
        }

        var rows = rowsHidden != false;
        var columns = columnsHidden != false;
        if (rows && columns)
        {
            return "hidden rows and columns";
        }

        return rows ? "hidden rows" : columns ? "hidden columns" : null;
    }

    private static string FromArray(Array array, IFormatProvider culture)
    {
        if (array.Length == 0)
        {
            return string.Empty;
        }

        object? first;
        string size;
        if (array.Rank == 2)
        {
            first = array.GetValue(array.GetLowerBound(0), array.GetLowerBound(1));
            size = string.Format(CultureInfo.InvariantCulture, "{0}x{1}", array.GetLength(0), array.GetLength(1));
        }
        else if (array.Rank == 1)
        {
            first = array.GetValue(array.GetLowerBound(0));
            size = array.Length.ToString(CultureInfo.InvariantCulture);
        }
        else
        {
            return "{array}";
        }

        var text = first is Array ? "{array}" : FromValue(first, culture);
        return (text.Length == 0 ? string.Empty : text + " ") + "{" + size + "}";
    }

    private static bool IsAllHashes(string text)
    {
        foreach (var c in text)
        {
            if (c != '#')
            {
                return false;
            }
        }

        return true;
    }

    // Line breaks become a return symbol; tabs and other control characters a space. Cut at MaxLength.
    private static string SingleLine(string text)
    {
        var sb = new StringBuilder(Math.Min(text.Length, MaxLength + 1));
        for (var i = 0; i < text.Length && sb.Length <= MaxLength; i++)
        {
            var c = text[i];
            if (c == '\r')
            {
                sb.Append('↵');
                if (i + 1 < text.Length && text[i + 1] == '\n')
                {
                    i++;
                }
            }
            else if (c == '\n')
            {
                sb.Append('↵');
            }
            else
            {
                sb.Append(char.IsControl(c) ? ' ' : c);
            }
        }

        if (sb.Length <= MaxLength)
        {
            return sb.ToString();
        }

        // Do not split a surrogate pair.
        var cut = char.IsHighSurrogate(sb[MaxLength - 2]) ? MaxLength - 2 : MaxLength - 1;
        return sb.ToString(0, cut) + "…";
    }
}
