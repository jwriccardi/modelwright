using System;
using System.Linq;
using ExcelModelingToolkit.Core.Trace;
using Xunit;

namespace ExcelModelingToolkit.Core.Tests.Trace;

// GC.GetTotalMemory counts every thread's allocations, so these tests run alone, after the parallel ones.
[CollectionDefinition(nameof(FormulaMemoryTests), DisableParallelization = true)]
public class FormulaMemoryCollection
{
}

[Collection(nameof(FormulaMemoryTests))]
public class FormulaMemoryTests
{
    private static readonly FormulaContext Context = new FormulaContext("Model.xlsx", "Calc");

    public static TheoryData<string> LongFormulas => new TheoryData<string>
    {
        // 8,190 characters of A1+A1+...: 2,730 references in an operator chain 2,729 levels deep.
        "=" + string.Join("+", Enumerable.Repeat("A1", 2730)),
        "=" + new string('(', 4000) + "A1" + new string(')', 4000),
        "=" + string.Concat(Enumerable.Repeat("INDEX(", 1300)) + "B1" + string.Concat(Enumerable.Repeat(",1)", 1300)),
        "=SUM(" + string.Join(":", Enumerable.Range(1, 1500).Select(row => "A" + row)) + ")",
    };

    [Theory]
    [MemberData(nameof(LongFormulas))]
    public void Parsed_formula_memory_grows_linearly_with_its_nesting(string formula)
    {
        // Each node's text is a span of its parent's, so stored texts would add up to the square of the nesting:
        // 22 MB for the operator chain, 31 MB for the parentheses. Without them each is under 1 MB.
        var other = "=" + string.Join("+", Enumerable.Repeat("B1", 200));
        FormulaParser.Parse(other, Context);
        var before = GC.GetTotalMemory(true);

        var parsed = FormulaParser.Parse(formula, Context);

        // XLParser keeps its last parse tree per thread; parsing another long formula lets that one go.
        FormulaParser.Parse(other, Context);
        var retained = GC.GetTotalMemory(true) - before;
        Assert.True(parsed.IsParsed, parsed.Error);
        Assert.True(retained < 5 * 1024 * 1024, $"The parsed formula holds {retained / 1024:N0} KB.");
        Assert.Equal(formula.Substring(1), parsed.Structure!.Text);
        GC.KeepAlive(parsed);
    }
}
