using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelModelingToolkit.Core.Trace;

/// <summary>
/// Excel's <c>Worksheet.Evaluate</c>, which Trace In uses for the calls that compute a reference (INDEX, OFFSET,
/// INDIRECT, CHOOSE) and for names that hold a formula, has no calling cell: <c>ROW()</c> is 1 there whatever cell
/// the formula is in. These helpers make a fragment evaluate as it would in its own cell where that can be done in the
/// text, and say where it cannot (docs/PLAN.md section 4.5).
/// </summary>
public static class CallingCell
{
    // R1C1 text that names one absolute cell or range (R2C3, Sheet1!R1C1:R5C2): exact wherever it is evaluated.
    private static readonly Regex AbsoluteR1C1 = new Regex(
        @"^(.*!)?R\d+C\d+(:R\d+C\d+)?$", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    /// <summary>
    /// <paramref name="expression"/> with every argument-less <c>ROW()</c> and <c>COLUMN()</c> replaced by
    /// <paramref name="row"/> and <paramref name="column"/>, the formula cell's. Matched in any case and with spaces
    /// inside the parentheses; never inside a string constant, a quoted sheet name or a bracketed part (a table column,
    /// a workbook name), never <c>ROWS</c>, <c>COLUMNS</c> or another name ending in ROW, and never <c>ROW(ref)</c>,
    /// which does not depend on the calling cell.
    /// </summary>
    /// <param name="expression">A formula or a fragment of one, with or without the leading <c>=</c>.</param>
    /// <param name="row">The formula cell's row (1-based).</param>
    /// <param name="column">The formula cell's column (1-based).</param>
    /// <exception cref="ArgumentNullException"><paramref name="expression"/> is null.</exception>
    public static string SubstituteRowAndColumn(string expression, int row, int column)
    {
        if (expression is null)
        {
            throw new ArgumentNullException(nameof(expression));
        }

        var result = new StringBuilder(expression.Length + 8);
        var index = 0;
        while (index < expression.Length)
        {
            var skipped = SkipQuoted(expression, index);
            if (skipped > index)
            {
                result.Append(expression, index, skipped - index);
                index = skipped;
                continue;
            }

            if (!IsTokenChar(expression[index]))
            {
                result.Append(expression[index]);
                index++;
                continue;
            }

            var end = TokenEnd(expression, index);
            var token = expression.Substring(index, end - index);
            var isRow = string.Equals(token, "ROW", StringComparison.OrdinalIgnoreCase);
            if ((isRow || string.Equals(token, "COLUMN", StringComparison.OrdinalIgnoreCase)) &&
                EmptyCallEnd(expression, end) is int callEnd)
            {
                result.Append((isRow ? row : column).ToString(CultureInfo.InvariantCulture));
                index = callEnd;
                continue;
            }

            result.Append(token);
            index = end;
        }

        return result.ToString();
    }

    /// <summary>
    /// True if <paramref name="expression"/> calls INDIRECT in R1C1 style (a second argument that is not omitted,
    /// <c>TRUE</c> or <c>1</c>) with text that may hold a relative reference (<c>RC[-1]</c>): evaluated without its
    /// calling cell, such a reference is relative to some other cell, so the result may be wrong. An R1C1 string
    /// constant naming only absolute cells (<c>"R2C3"</c>) is exact and not counted.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="expression"/> is null.</exception>
    public static bool HasRelativeR1C1Indirect(string expression)
    {
        if (expression is null)
        {
            throw new ArgumentNullException(nameof(expression));
        }

        var index = 0;
        while (index < expression.Length)
        {
            var skipped = SkipQuoted(expression, index);
            if (skipped > index)
            {
                index = skipped;
                continue;
            }

            if (!IsTokenChar(expression[index]))
            {
                index++;
                continue;
            }

            var end = TokenEnd(expression, index);
            var token = expression.Substring(index, end - index);
            index = end;
            if (!string.Equals(token, "INDIRECT", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(token, "_xlfn.INDIRECT", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var arguments = Arguments(expression, end);
            if (arguments is null || arguments.Count < 2)
            {
                continue;
            }

            var a1 = arguments[1].Trim();
            if (a1.Length == 0 || string.Equals(a1, "TRUE", StringComparison.OrdinalIgnoreCase) || a1 == "1")
            {
                continue;
            }

            var text = arguments[0].Trim();
            var isConstant = text.Length >= 2 && text[0] == '"' && text[text.Length - 1] == '"' &&
                text.IndexOf('"', 1) == text.Length - 1;
            if (!(isConstant && AbsoluteR1C1.IsMatch(text.Substring(1, text.Length - 2))))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// True if <paramref name="reference"/> (a cell, range, whole column or whole row) has a row or column written
    /// without <c>$</c>, so it moves with the cell it is evaluated from: in a defined name, relative to the active
    /// cell (<c>Name.RefersTo</c>) unless converted for the formula's cell. False for other kinds.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="reference"/> is null.</exception>
    public static bool IsRelative(FormulaReference reference)
    {
        if (reference is null)
        {
            throw new ArgumentNullException(nameof(reference));
        }

        switch (reference.Kind)
        {
            case FormulaReferenceKind.Cell:
            case FormulaReferenceKind.Range:
            case FormulaReferenceKind.WholeColumn:
            case FormulaReferenceKind.WholeRow:
                break;
            default:
                return false;
        }

        var text = reference.Text;
        var address = text.Substring(text.LastIndexOf('!') + 1).TrimEnd('#');
        for (var index = 0; index < address.Length; index++)
        {
            var c = address[index];
            var startsRun = (char.IsLetter(c) && (index == 0 || !char.IsLetter(address[index - 1]))) ||
                (char.IsDigit(c) && (index == 0 || !char.IsDigit(address[index - 1])));
            if (startsRun && (index == 0 || address[index - 1] != '$'))
            {
                return true;
            }
        }

        return false;
    }

    // Characters of a function name, defined name, cell address or number: read as one token, so ROWS, MYROW and
    // A1ROW are never mistaken for ROW.
    private static bool IsTokenChar(char c) => char.IsLetterOrDigit(c) || c == '_' || c == '.' || c == '\\' || c == '?';

    private static int TokenEnd(string text, int start)
    {
        var end = start;
        while (end < text.Length && IsTokenChar(text[end]))
        {
            end++;
        }

        return end;
    }

    // The position after "( )" (spaces allowed) at index, or null if no empty call follows.
    private static int? EmptyCallEnd(string text, int index)
    {
        index = SkipSpaces(text, index);
        if (index >= text.Length || text[index] != '(')
        {
            return null;
        }

        index = SkipSpaces(text, index + 1);
        return index < text.Length && text[index] == ')' ? index + 1 : (int?)null;
    }

    private static int SkipSpaces(string text, int index)
    {
        while (index < text.Length && char.IsWhiteSpace(text[index]))
        {
            index++;
        }

        return index;
    }

    // If a string constant ("..."), quoted sheet name ('...') or bracketed part ([...], nested, with ' escapes) starts
    // at index, the position after it (the rest of the text if it is not closed); else index.
    private static int SkipQuoted(string text, int index)
    {
        var c = text[index];
        if (c == '"' || c == '\'')
        {
            var i = index + 1;
            while (i < text.Length)
            {
                if (text[i] == c)
                {
                    if (i + 1 < text.Length && text[i + 1] == c)
                    {
                        i += 2; // "" or '' is one quote character
                        continue;
                    }

                    return i + 1;
                }

                i++;
            }

            return text.Length;
        }

        if (c == '[')
        {
            var depth = 0;
            var i = index;
            while (i < text.Length)
            {
                switch (text[i])
                {
                    case '\'':
                        i += 2; // a table column's escape: the next character is literal
                        continue;
                    case '[':
                        depth++;
                        break;
                    case ']':
                        if (--depth == 0)
                        {
                            return i + 1;
                        }

                        break;
                }

                i++;
            }

            return text.Length;
        }

        return index;
    }

    // The top-level arguments of the call whose name ends at index (before its '('), or null if no '(' follows.
    private static List<string>? Arguments(string text, int index)
    {
        index = SkipSpaces(text, index);
        if (index >= text.Length || text[index] != '(')
        {
            return null;
        }

        var arguments = new List<string>();
        var depth = 0;
        var start = index + 1;
        var i = start;
        while (i < text.Length)
        {
            var skipped = SkipQuoted(text, i);
            if (skipped > i)
            {
                i = skipped;
                continue;
            }

            switch (text[i])
            {
                case '(':
                case '{':
                    depth++;
                    break;
                case ')':
                case '}':
                    if (depth == 0)
                    {
                        arguments.Add(text.Substring(start, i - start));
                        return arguments;
                    }

                    depth--;
                    break;
                case ',':
                    if (depth == 0)
                    {
                        arguments.Add(text.Substring(start, i - start));
                        start = i + 1;
                    }

                    break;
            }

            i++;
        }

        arguments.Add(text.Substring(start));
        return arguments;
    }
}
