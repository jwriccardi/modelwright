using Modelwright.Core.Trace;
using Xunit;

namespace Modelwright.Core.Tests.Trace;

public class TraceUiStateTests
{
    [Fact]
    public void State_round_trips()
    {
        var state = new TraceUiState(new WindowRect(-1200, 300, 620, 380), wrapFormula: true);

        var json = state.ToJson();
        var read = TraceUiState.FromJson(json);

        Assert.Equal(new WindowRect(-1200, 300, 620, 380), read.Bounds);
        Assert.True(read.WrapFormula);
        Assert.Equal(json, read.ToJson());
        Assert.Contains("\"schemaVersion\": 1", json);
    }

    [Fact]
    public void Json_is_readable()
    {
        var json = new TraceUiState(new WindowRect(10, 20, 30, 40)).ToJson();

        Assert.Equal(
            "{\r\n  \"schemaVersion\": 1,\r\n  \"traceWindow\": {\r\n    \"left\": 10,\r\n    \"top\": 20,\r\n" +
            "    \"width\": 30,\r\n    \"height\": 40,\r\n    \"wrapFormula\": false\r\n  }\r\n}\r\n",
            json);
    }

    [Fact]
    public void Unknown_bounds_are_left_out()
    {
        var json = new TraceUiState().ToJson();

        Assert.DoesNotContain("left", json);
        Assert.Null(TraceUiState.FromJson(json).Bounds);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("not json")]
    [InlineData("[]")]
    [InlineData("{}")]
    [InlineData("{ \"traceWindow\": 5 }")]
    [InlineData("{ \"traceWindow\": { \"left\": 1, \"top\": 2, \"width\": 3 } }")]
    [InlineData("{ \"traceWindow\": { \"left\": 1, \"top\": 2, \"width\": 0, \"height\": 4 } }")]
    [InlineData("{ \"traceWindow\": { \"left\": 1, \"top\": 2, \"width\": -3, \"height\": 4 } }")]
    [InlineData("{ \"traceWindow\": { \"left\": 1.5, \"top\": 2, \"width\": 3, \"height\": 4 } }")]
    [InlineData("{ \"traceWindow\": { \"left\": \"1\", \"top\": 2, \"width\": 3, \"height\": 4 } }")]
    [InlineData("{ \"traceWindow\": { \"left\": 99999999999, \"top\": 2, \"width\": 3, \"height\": 4 } }")]
    [InlineData("{ \"traceWindow\": { \"left\": 1, \"top\": 2, \"width\": 3000000, \"height\": 4 } }")]
    public void Unusable_bounds_give_the_default_place(string? json)
    {
        var state = TraceUiState.FromJson(json);

        Assert.Null(state.Bounds);
        Assert.False(state.WrapFormula);
    }

    [Fact]
    public void Unknown_properties_are_ignored_and_bad_values_cost_only_their_setting()
    {
        var state = TraceUiState.FromJson(
            "{ \"schemaVersion\": 7, \"other\": [1], \"traceWindow\": { \"left\": 1, \"top\": 2, \"width\": 3, \"height\": 4, " +
            "\"wrapFormula\": \"yes\", \"columns\": [100, 50] } }");

        Assert.Equal(new WindowRect(1, 2, 3, 4), state.Bounds);
        Assert.False(state.WrapFormula);
    }

    [Fact]
    public void With_changes_one_setting()
    {
        var state = new TraceUiState(new WindowRect(1, 2, 3, 4));

        Assert.True(state.WithWrapFormula(true).WrapFormula);
        Assert.Equal(state.Bounds, state.WithWrapFormula(true).Bounds);
        Assert.Null(state.WithBounds(null).Bounds);
        Assert.Equal(new WindowRect(5, 6, 7, 8), state.WithBounds(new WindowRect(5, 6, 7, 8)).Bounds);
    }
}
