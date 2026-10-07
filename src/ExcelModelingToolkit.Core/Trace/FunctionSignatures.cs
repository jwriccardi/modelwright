using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ExcelModelingToolkit.Core.Trace;

/// <summary>
/// Excel's parameter names for common functions, for the Argument column of Trace In's evaluate mode
/// (docs/PLAN.md section 4.5; research/07). Names are written the way Excel's formula tooltip writes them: optional
/// parameters in brackets (<c>IF(logical_test, [value_if_true], [value_if_false])</c>) and a repeating tail as
/// <c>number1, [number2], ...</c>.
/// </summary>
/// <remarks>
/// Source: the "Syntax" section of each function's page in Microsoft's public Excel function reference
/// (https://support.microsoft.com/en-us/office/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188),
/// which lists every parameter's name and whether it is optional. Where Excel's tooltip brackets a parameter the
/// reference calls required (IF's value_if_true, which can be left empty), the tooltip is followed, as Macabacus does.
/// INDEX has two forms; its entry uses the array form's names and adds the reference form's <c>[area_num]</c>.
/// </remarks>
public static class FunctionSignatures
{
    private const string Repeat = "...";

    private static readonly Dictionary<string, Signature> Signatures = Build(new Dictionary<string, string>
    {
        ["ABS"] = "number",
        ["ADDRESS"] = "row_num, column_num, [abs_num], [a1], [sheet_text]",
        ["AGGREGATE"] = "function_num, options, ref1, [ref2], ...",
        ["AND"] = "logical1, [logical2], ...",
        ["AVERAGE"] = "number1, [number2], ...",
        ["AVERAGEIF"] = "range, criteria, [average_range]",
        ["AVERAGEIFS"] = "average_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...",
        ["BYCOL"] = "array, lambda",
        ["BYROW"] = "array, lambda",
        ["CEILING"] = "number, significance",
        ["CEILING.MATH"] = "number, [significance], [mode]",
        ["CELL"] = "info_type, [reference]",
        ["CHOOSE"] = "index_num, value1, [value2], ...",
        ["CHOOSECOLS"] = "array, col_num1, [col_num2], ...",
        ["CHOOSEROWS"] = "array, row_num1, [row_num2], ...",
        ["COLUMN"] = "[reference]",
        ["COLUMNS"] = "array",
        ["CONCAT"] = "text1, [text2], ...",
        ["CONCATENATE"] = "text1, [text2], ...",
        ["CORREL"] = "array1, array2",
        ["COUNT"] = "value1, [value2], ...",
        ["COUNTA"] = "value1, [value2], ...",
        ["COUNTBLANK"] = "range",
        ["COUNTIF"] = "range, criteria",
        ["COUNTIFS"] = "criteria_range1, criteria1, [criteria_range2, criteria2], ...",
        ["CUMIPMT"] = "rate, nper, pv, start_period, end_period, type",
        ["CUMPRINC"] = "rate, nper, pv, start_period, end_period, type",
        ["DATE"] = "year, month, day",
        ["DATEDIF"] = "start_date, end_date, unit",
        ["DATEVALUE"] = "date_text",
        ["DAY"] = "serial_number",
        ["DAYS"] = "end_date, start_date",
        ["DB"] = "cost, salvage, life, period, [month]",
        ["DDB"] = "cost, salvage, life, period, [factor]",
        ["DROP"] = "array, rows, [columns]",
        ["EDATE"] = "start_date, months",
        ["EFFECT"] = "nominal_rate, npery",
        ["EOMONTH"] = "start_date, months",
        ["EXP"] = "number",
        ["EXPAND"] = "array, rows, [columns], [pad_with]",
        ["FALSE"] = "",
        ["FILTER"] = "array, include, [if_empty]",
        ["FIND"] = "find_text, within_text, [start_num]",
        ["FLOOR"] = "number, significance",
        ["FLOOR.MATH"] = "number, [significance], [mode]",
        ["FORECAST.LINEAR"] = "x, known_y's, known_x's",
        ["FORMULATEXT"] = "reference",
        ["FV"] = "rate, nper, pmt, [pv], [type]",
        ["HLOOKUP"] = "lookup_value, table_array, row_index_num, [range_lookup]",
        ["HSTACK"] = "array1, [array2], ...",
        ["HYPERLINK"] = "link_location, [friendly_name]",
        ["IF"] = "logical_test, [value_if_true], [value_if_false]",
        ["IFERROR"] = "value, value_if_error",
        ["IFNA"] = "value, value_if_na",
        ["IFS"] = "logical_test1, value_if_true1, [logical_test2, value_if_true2], ...",
        ["INDEX"] = "array, row_num, [column_num], [area_num]",
        ["INDIRECT"] = "ref_text, [a1]",
        ["INT"] = "number",
        ["INTERCEPT"] = "known_y's, known_x's",
        ["IPMT"] = "rate, per, nper, pv, [fv], [type]",
        ["IRR"] = "values, [guess]",
        ["ISBLANK"] = "value",
        ["ISERR"] = "value",
        ["ISERROR"] = "value",
        ["ISEVEN"] = "number",
        ["ISFORMULA"] = "reference",
        ["ISLOGICAL"] = "value",
        ["ISNA"] = "value",
        ["ISNUMBER"] = "value",
        ["ISODD"] = "number",
        ["ISREF"] = "value",
        ["ISTEXT"] = "value",
        ["LARGE"] = "array, k",
        ["LEFT"] = "text, [num_chars]",
        ["LEN"] = "text",
        ["LET"] = "name1, name_value1, calculation_or_name2, [name_value2, calculation_or_name3], ...",
        ["LN"] = "number",
        ["LOG"] = "number, [base]",
        ["LOOKUP"] = "lookup_value, lookup_vector, [result_vector]",
        ["LOWER"] = "text",
        ["MAKEARRAY"] = "rows, cols, lambda",
        ["MATCH"] = "lookup_value, lookup_array, [match_type]",
        ["MAX"] = "number1, [number2], ...",
        ["MAXIFS"] = "max_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...",
        ["MEDIAN"] = "number1, [number2], ...",
        ["MID"] = "text, start_num, num_chars",
        ["MIN"] = "number1, [number2], ...",
        ["MINIFS"] = "min_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...",
        ["MIRR"] = "values, finance_rate, reinvest_rate",
        ["MOD"] = "number, divisor",
        ["MONTH"] = "serial_number",
        ["MROUND"] = "number, multiple",
        ["N"] = "value",
        ["NA"] = "",
        ["NETWORKDAYS"] = "start_date, end_date, [holidays]",
        ["NETWORKDAYS.INTL"] = "start_date, end_date, [weekend], [holidays]",
        ["NOMINAL"] = "effect_rate, npery",
        ["NOT"] = "logical",
        ["NOW"] = "",
        ["NPER"] = "rate, pmt, pv, [fv], [type]",
        ["NPV"] = "rate, value1, [value2], ...",
        ["OFFSET"] = "reference, rows, cols, [height], [width]",
        ["OR"] = "logical1, [logical2], ...",
        ["PERCENTILE.INC"] = "array, k",
        ["PMT"] = "rate, nper, pv, [fv], [type]",
        ["POWER"] = "number, power",
        ["PPMT"] = "rate, per, nper, pv, [fv], [type]",
        ["PRODUCT"] = "number1, [number2], ...",
        ["PROPER"] = "text",
        ["PV"] = "rate, nper, pmt, [fv], [type]",
        ["QUARTILE.INC"] = "array, quart",
        ["QUOTIENT"] = "numerator, denominator",
        ["RAND"] = "",
        ["RANDBETWEEN"] = "bottom, top",
        ["RANK.EQ"] = "number, ref, [order]",
        ["RATE"] = "nper, pmt, pv, [fv], [type], [guess]",
        ["REDUCE"] = "[initial_value], array, lambda",
        ["REPT"] = "text, number_times",
        ["RIGHT"] = "text, [num_chars]",
        ["ROUND"] = "number, num_digits",
        ["ROUNDDOWN"] = "number, num_digits",
        ["ROUNDUP"] = "number, num_digits",
        ["ROW"] = "[reference]",
        ["ROWS"] = "array",
        ["SCAN"] = "[initial_value], array, lambda",
        ["SEARCH"] = "find_text, within_text, [start_num]",
        ["SEQUENCE"] = "rows, [columns], [start], [step]",
        ["SIGN"] = "number",
        ["SLN"] = "cost, salvage, life",
        ["SLOPE"] = "known_y's, known_x's",
        ["SMALL"] = "array, k",
        ["SORT"] = "array, [sort_index], [sort_order], [by_col]",
        ["SORTBY"] = "array, by_array1, [sort_order1], [by_array2, sort_order2], ...",
        ["SQRT"] = "number",
        ["STDEV.P"] = "number1, [number2], ...",
        ["STDEV.S"] = "number1, [number2], ...",
        ["SUBSTITUTE"] = "text, old_text, new_text, [instance_num]",
        ["SUBTOTAL"] = "function_num, ref1, [ref2], ...",
        ["SUM"] = "number1, [number2], ...",
        ["SUMIF"] = "range, criteria, [sum_range]",
        ["SUMIFS"] = "sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...",
        ["SUMPRODUCT"] = "array1, [array2], [array3], ...",
        ["SUMSQ"] = "number1, [number2], ...",
        ["SWITCH"] = "expression, value1, result1, [default_or_value2, result2], ...",
        ["TAKE"] = "array, rows, [columns]",
        ["TEXT"] = "value, format_text",
        ["TEXTAFTER"] = "text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found]",
        ["TEXTBEFORE"] = "text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found]",
        ["TEXTJOIN"] = "delimiter, ignore_empty, text1, [text2], ...",
        ["TEXTSPLIT"] = "text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with]",
        ["TOCOL"] = "array, [ignore], [scan_by_column]",
        ["TODAY"] = "",
        ["TOROW"] = "array, [ignore], [scan_by_column]",
        ["TRANSPOSE"] = "array",
        ["TREND"] = "known_y's, [known_x's], [new_x's], [const]",
        ["TRIM"] = "text",
        ["TRUE"] = "",
        ["TRUNC"] = "number, [num_digits]",
        ["UNIQUE"] = "array, [by_col], [exactly_once]",
        ["UPPER"] = "text",
        ["VALUE"] = "text",
        ["VLOOKUP"] = "lookup_value, table_array, col_index_num, [range_lookup]",
        ["VSTACK"] = "array1, [array2], ...",
        ["WEEKDAY"] = "serial_number, [return_type]",
        ["WORKDAY"] = "start_date, days, [holidays]",
        ["WORKDAY.INTL"] = "start_date, days, [weekend], [holidays]",
        ["XIRR"] = "values, dates, [guess]",
        ["XLOOKUP"] = "lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]",
        ["XMATCH"] = "lookup_value, lookup_array, [match_mode], [search_mode]",
        ["XNPV"] = "rate, values, dates",
        ["XOR"] = "logical1, [logical2], ...",
        ["YEAR"] = "serial_number",
        ["YEARFRAC"] = "start_date, end_date, [basis]",
    });

    /// <summary>The functions in the table (upper case), plus LAMBDA, whose names depend on its argument count.</summary>
    public static IReadOnlyCollection<string> FunctionNames { get; } =
        Signatures.Keys.Concat(new[] { "LAMBDA" }).OrderBy(name => name, StringComparer.Ordinal).ToList();

    /// <summary>
    /// A function name as the table keys it: upper case, without the <c>_xlfn.</c> and <c>_xlws.</c> prefixes Excel
    /// stores for functions newer than Excel 2007 (<c>_xlfn._xlws.SORT</c> is <c>SORT</c>).
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="functionName"/> is null.</exception>
    public static string Normalize(string functionName)
    {
        if (functionName is null)
        {
            throw new ArgumentNullException(nameof(functionName));
        }

        var name = functionName.Trim().ToUpperInvariant();
        while (true)
        {
            if (name.StartsWith("_XLFN.", StringComparison.Ordinal) || name.StartsWith("_XLWS.", StringComparison.Ordinal))
            {
                name = name.Substring(6);
                continue;
            }

            return name;
        }
    }

    /// <summary>True if the function's parameter names are known.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="functionName"/> is null.</exception>
    public static bool Contains(string functionName)
    {
        var name = Normalize(functionName);
        return name == "LAMBDA" || Signatures.ContainsKey(name);
    }

    /// <summary>
    /// The function's tooltip-style syntax, <c>IF(logical_test, [value_if_true], [value_if_false])</c>, or null if it
    /// is not in the table.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="functionName"/> is null.</exception>
    public static string? GetSyntax(string functionName)
    {
        var name = Normalize(functionName);
        if (name == "LAMBDA")
        {
            return "LAMBDA([parameter1, parameter2, ...], calculation)";
        }

        return Signatures.TryGetValue(name, out var signature) ? name + "(" + signature.Text + ")" : null;
    }

    /// <summary>
    /// The display name of a function's argument, as Excel's tooltip writes it: <c>logical_test</c>,
    /// <c>[value_if_true]</c>; a repeating tail continues its numbering (SUM's fourth argument is
    /// <c>[number4]</c>, SUMIFS's sixth is <c>[criteria_range3]</c>).
    /// </summary>
    /// <param name="functionName">The function, in any case, with or without <c>_xlfn.</c>.</param>
    /// <param name="argumentIndex">The 0-based argument position.</param>
    /// <param name="argumentCount">The number of arguments in the call (LAMBDA's last argument is its calculation).</param>
    /// <returns>The name, or null if the function is unknown or takes no argument at that position.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="functionName"/> is null.</exception>
    public static string? GetParameterName(string functionName, int argumentIndex, int argumentCount)
    {
        var name = Normalize(functionName);
        if (argumentIndex < 0 || argumentIndex >= argumentCount)
        {
            return null;
        }

        if (name == "LAMBDA")
        {
            return argumentIndex == argumentCount - 1
                ? "calculation"
                : "[parameter" + (argumentIndex + 1).ToString(CultureInfo.InvariantCulture) + "]";
        }

        if (!Signatures.TryGetValue(name, out var signature))
        {
            return null;
        }

        if (argumentIndex < signature.Parameters.Count)
        {
            return signature.Parameters[argumentIndex].Display;
        }

        if (signature.RepeatStart < 0)
        {
            return null;
        }

        // Repeat the last bracketed group, numbering each name on from where the group starts.
        var groupSize = signature.Parameters.Count - signature.RepeatStart;
        var offset = argumentIndex - signature.RepeatStart;
        var parameter = signature.Parameters[signature.RepeatStart + (offset % groupSize)];
        var renamed = Renumber(parameter.Name, offset / groupSize);
        return parameter.Optional ? "[" + renamed + "]" : renamed;
    }

    // number2 advanced by 3 is number5; a name without a trailing number is repeated as is.
    private static string Renumber(string name, int step)
    {
        var digits = name.Length;
        while (digits > 0 && char.IsDigit(name[digits - 1]))
        {
            digits--;
        }

        if (digits == name.Length)
        {
            return name;
        }

        var number = int.Parse(name.Substring(digits), NumberStyles.None, CultureInfo.InvariantCulture);
        return name.Substring(0, digits) + (number + step).ToString(CultureInfo.InvariantCulture);
    }

    private static Dictionary<string, Signature> Build(Dictionary<string, string> table)
    {
        var result = new Dictionary<string, Signature>(StringComparer.Ordinal);
        foreach (var entry in table)
        {
            result.Add(entry.Key, Signature.Parse(entry.Value));
        }

        return result;
    }

    private sealed class Parameter
    {
        public Parameter(string name, bool optional, int group)
        {
            Name = name;
            Optional = optional;
            Group = group;
        }

        public string Name { get; }

        public bool Optional { get; }

        // Parameters written in one pair of brackets share a group; required ones have group -1.
        public int Group { get; }

        public string Display => Optional ? "[" + Name + "]" : Name;
    }

    private sealed class Signature
    {
        private Signature(string text, List<Parameter> parameters, int repeatStart)
        {
            Text = text;
            Parameters = parameters;
            RepeatStart = repeatStart;
        }

        public string Text { get; }

        public IReadOnlyList<Parameter> Parameters { get; }

        // The index of the first parameter of the group that repeats after "...", or -1.
        public int RepeatStart { get; }

        // "sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ..."
        public static Signature Parse(string text)
        {
            var parameters = new List<Parameter>();
            var repeats = false;
            var group = -1;
            var groups = 0;
            foreach (var raw in text.Split(','))
            {
                var token = raw.Trim();
                if (token.Length == 0)
                {
                    continue;
                }

                if (token == Repeat)
                {
                    repeats = true;
                    continue;
                }

                if (token.StartsWith("[", StringComparison.Ordinal))
                {
                    group = groups++;
                    token = token.Substring(1);
                }

                var closes = token.EndsWith("]", StringComparison.Ordinal);
                if (closes)
                {
                    token = token.Substring(0, token.Length - 1);
                }

                parameters.Add(new Parameter(token, group >= 0, group));
                if (closes)
                {
                    group = -1;
                }
            }

            var repeatStart = -1;
            if (repeats)
            {
                var lastGroup = parameters[parameters.Count - 1].Group;
                repeatStart = lastGroup < 0
                    ? parameters.Count - 1
                    : parameters.FindIndex(parameter => parameter.Group == lastGroup);
            }

            return new Signature(text, parameters, repeatStart);
        }
    }
}
