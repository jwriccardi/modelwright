using System;
using System.Diagnostics;
using System.Linq;
using ExcelModelingToolkit.Core.Trace;
using Xunit;

namespace ExcelModelingToolkit.Core.Tests.Trace;

// Timings run alone, after the parallel tests: other tests' long formulas share the parse thread, and on a busy CI
// runner waiting behind them would be measured too.
[CollectionDefinition(nameof(FormulaParserTimingTests), DisableParallelization = true)]
public class FormulaParserTimingCollection
{
}

[Collection(nameof(FormulaParserTimingTests))]
public class FormulaParserTimingTests
{
    private static readonly FormulaContext Context = new FormulaContext("Model.xlsx", "Calc");

    [Theory]
    [InlineData(5)]
    [InlineData(20)]
    public void Parsing_is_fast_after_the_first_call(int references)
    {
        // docs/PLAN.md Phase 4: a formula with 20 references opens in 300 ms, Excel reads included. Five references
        // parse on the caller's thread; twenty (over 256 characters) on the parse thread.
        var formula = "=" + string.Join("+", Enumerable.Range(1, references).Select(i => "'Sheet " + i + "'!$B$" + i));
        Assert.Equal(references, FormulaParser.Parse(formula, Context).References.Count);

        var times = new double[51];
        for (var i = 0; i < times.Length; i++)
        {
            var watch = Stopwatch.StartNew();
            var parsed = FormulaParser.Parse(formula, Context);
            times[i] = watch.Elapsed.TotalMilliseconds;
            Assert.Equal(references, parsed.References.Count);
        }

        // The median, so one pause (a GC, the runner scheduling another process) can't fail it.
        Array.Sort(times);
        var median = times[times.Length / 2];
        Assert.True(median < 30, $"median parse took {median:0.00} ms.");
    }
}
