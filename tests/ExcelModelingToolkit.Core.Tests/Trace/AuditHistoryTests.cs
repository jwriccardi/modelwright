using System;
using ExcelModelingToolkit.Core.Trace;
using Xunit;
using static ExcelModelingToolkit.Core.Tests.Trace.FakePrecedentProvider;

namespace ExcelModelingToolkit.Core.Tests.Trace;

public class AuditHistoryTests
{
    [Fact]
    public void Last_audited_cell_returns_through_three_levels()
    {
        var history = new AuditHistory();
        history.Push(Cell("A1"));
        history.Push(Cell("B5", "Debt"));
        history.Push(Cell("C9", "Returns"));

        Assert.Equal("Returns!C9", history.Pop()!.Label);
        Assert.Equal("Debt!B5", history.Pop()!.Label);
        Assert.Equal("Calc!A1", history.Pop()!.Label);
        Assert.Null(history.Pop());
    }

    [Fact]
    public void Holds_twenty_audits_and_drops_the_oldest()
    {
        var history = new AuditHistory();
        for (var i = 1; i <= 25; i++)
        {
            history.Push(Cell("A" + i));
        }

        Assert.Equal(20, history.Capacity);
        Assert.Equal(20, history.Count);
        PrecedentItem? last = null;
        while (history.Count > 0)
        {
            last = history.Pop();
        }

        Assert.Equal("Calc!A6", last!.Label);
    }

    [Fact]
    public void Reauditing_the_newest_cell_records_nothing()
    {
        var history = new AuditHistory();
        history.Push(Cell("A1"));
        history.Push(new PrecedentItem(PrecedentKind.Cell, "same", "MODEL.xlsx", "calc", "A1"));
        history.Push(Cell("B1"));
        history.Push(Cell("A1"));

        Assert.Equal(3, history.Count);
    }

    [Fact]
    public void Peek_and_clear()
    {
        var history = new AuditHistory(capacity: 2);
        Assert.Null(history.Peek());
        history.Push(Cell("A1"));

        Assert.Equal("Calc!A1", history.Peek()!.Label);
        Assert.Equal(1, history.Count);
        history.Clear();
        Assert.Equal(0, history.Count);
    }

    [Fact]
    public void Arguments_are_checked()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new AuditHistory(0));
        Assert.Throws<ArgumentNullException>(() => new AuditHistory().Push(null!));
    }
}
