using System;
using System.Linq;
using ExcelModelingToolkit.Core.Trace;
using Xunit;

namespace ExcelModelingToolkit.Core.Tests.Trace;

public class FunctionSignaturesTests
{
    [Fact]
    public void Table_covers_at_least_100_functions_including_the_modeling_core()
    {
        Assert.True(FunctionSignatures.FunctionNames.Count >= 100, $"{FunctionSignatures.FunctionNames.Count} functions.");
        foreach (var name in new[]
        {
            "SUM", "IF", "IFERROR", "INDEX", "MATCH", "XLOOKUP", "VLOOKUP", "HLOOKUP", "OFFSET", "INDIRECT", "SUMIFS",
            "SUMPRODUCT", "AVERAGE", "MIN", "MAX", "ROUND", "ROUNDUP", "ROUNDDOWN", "ABS", "AND", "OR", "NOT", "CHOOSE",
            "EOMONTH", "EDATE", "YEAR", "DATE", "IRR", "XIRR", "NPV", "XNPV", "PMT", "LET", "LAMBDA",
        })
        {
            Assert.True(FunctionSignatures.Contains(name), name);
        }
    }

    [Theory]
    [InlineData("IF", "IF(logical_test, [value_if_true], [value_if_false])")]
    [InlineData("if", "IF(logical_test, [value_if_true], [value_if_false])")]
    [InlineData("_xlfn.XLOOKUP", "XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])")]
    [InlineData("SUM", "SUM(number1, [number2], ...)")]
    [InlineData("TODAY", "TODAY()")]
    [InlineData("LAMBDA", "LAMBDA([parameter1, parameter2, ...], calculation)")]
    [InlineData("NOSUCHFUNCTION", null)]
    public void Syntax_is_the_tooltip_text(string function, string? expected)
    {
        Assert.Equal(expected, FunctionSignatures.GetSyntax(function));
    }

    [Theory]
    [InlineData("IF", 0, 3, "logical_test")]
    [InlineData("IF", 1, 3, "[value_if_true]")]
    [InlineData("IF", 2, 3, "[value_if_false]")]
    [InlineData("IF", 3, 4, null)]
    [InlineData("VLOOKUP", 3, 4, "[range_lookup]")]
    [InlineData("INDEX", 2, 3, "[column_num]")]
    [InlineData("OFFSET", 4, 5, "[width]")]
    [InlineData("SUM", 0, 1, "number1")]
    [InlineData("SUM", 9, 10, "[number10]")]
    [InlineData("NPV", 3, 4, "[value3]")]
    [InlineData("SUMPRODUCT", 1, 4, "[array2]")]
    [InlineData("SUMPRODUCT", 3, 4, "[array4]")]
    [InlineData("SUMIFS", 5, 7, "[criteria_range3]")]
    [InlineData("SUMIFS", 6, 7, "[criteria3]")]
    [InlineData("COUNTIFS", 2, 4, "[criteria_range2]")]
    [InlineData("IFS", 4, 6, "[logical_test3]")]
    [InlineData("IFS", 5, 6, "[value_if_true3]")]
    [InlineData("LET", 4, 7, "[calculation_or_name3]")]
    [InlineData("LET", 5, 7, "[name_value3]")]
    [InlineData("LET", 6, 7, "[calculation_or_name4]")]
    [InlineData("SWITCH", 4, 6, "[result2]")]
    [InlineData("SWITCH", 5, 6, "[default_or_value3]")]
    [InlineData("CHOOSE", 4, 5, "[value4]")]
    [InlineData("TEXTJOIN", 4, 5, "[text3]")]
    [InlineData("LAMBDA", 0, 3, "[parameter1]")]
    [InlineData("LAMBDA", 2, 3, "calculation")]
    [InlineData("_xlfn._xlws.SORT", 1, 2, "[sort_index]")]
    [InlineData("NOSUCHFUNCTION", 0, 1, null)]
    [InlineData("SUM", -1, 1, null)]
    [InlineData("SUM", 1, 1, null)]
    public void Parameter_names_follow_excel_tooltips(string function, int index, int count, string? expected)
    {
        Assert.Equal(expected, FunctionSignatures.GetParameterName(function, index, count));
    }

    [Fact]
    public void Every_listed_function_names_its_listed_parameters()
    {
        foreach (var name in FunctionSignatures.FunctionNames)
        {
            var syntax = FunctionSignatures.GetSyntax(name)!;
            var listed = syntax.Substring(syntax.IndexOf('(') + 1).TrimEnd(')')
                .Split(',').Select(part => part.Trim()).Count(part => part.Length > 0 && part != "...");
            for (var index = 0; index < listed; index++)
            {
                var parameter = FunctionSignatures.GetParameterName(name, index, listed);
                Assert.False(string.IsNullOrWhiteSpace(parameter), name + " " + index);
                Assert.DoesNotContain(" ", parameter);
            }
        }
    }

    [Theory]
    [InlineData("sum", "SUM")]
    [InlineData(" _xlfn.XLOOKUP ", "XLOOKUP")]
    [InlineData("_xlfn._xlws.FILTER", "FILTER")]
    [InlineData("STDEV.S", "STDEV.S")]
    public void Normalize_removes_storage_prefixes(string name, string expected)
    {
        Assert.Equal(expected, FunctionSignatures.Normalize(name));
    }

    [Fact]
    public void Null_names_are_rejected()
    {
        Assert.Throws<ArgumentNullException>(() => FunctionSignatures.Normalize(null!));
        Assert.Throws<ArgumentNullException>(() => FunctionSignatures.GetParameterName(null!, 0, 1));
    }
}
