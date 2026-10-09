using Modelwright.Core.Trace;
using Xunit;

namespace Modelwright.Core.Tests.Trace;

public class TraceUiStateTests
{
    [Fact]
    public void State_round_trips()
    {
        var state = new TraceUiState(new WindowRect(-1200, 300, 620, 380), wrapFormula: true, evaluateFunctions: true);

        var json = state.ToJson();
        var read = TraceUiState.FromJson(json);

        Assert.Equal(new WindowRect(-1200, 300, 620, 380), read.Bounds);
        Assert.True(read.WrapFormula);
        Assert.True(read.EvaluateFunctions);
        Assert.Equal(json, read.ToJson());
        Assert.Contains("\"schemaVersion\": 1", json);
    }

    [Fact]
    public void Json_is_readable()
    {
        var json = new TraceUiState(new WindowRect(10, 20, 30, 40)).ToJson();

        Assert.Equal(
            "{\r\n  \"schemaVersion\": 1,\r\n  \"traceWindow\": {\r\n    \"left\": 10,\r\n    \"top\": 20,\r\n" +
            "    \"width\": 30,\r\n    \"height\": 40,\r\n    \"wrapFormula\": false,\r\n    \"evaluateFunctions\": false\r\n  }\r\n}\r\n",
            json);
    }

    [Fact]
    public void Macabacus_notice_flag_round_trips_and_is_written_only_when_true()
    {
        var state = new TraceUiState(new WindowRect(1, 2, 3, 4), wrapFormula: true).WithMacabacusNoticeShown(true);

        var json = state.ToJson();
        var read = TraceUiState.FromJson(json);

        Assert.True(read.MacabacusNoticeShown);
        Assert.Equal(new WindowRect(1, 2, 3, 4), read.Bounds);
        Assert.True(read.WrapFormula);
        Assert.Equal(json, read.ToJson());
        Assert.EndsWith("  },\r\n  \"macabacusNoticeShown\": true\r\n}\r\n", json);
        Assert.DoesNotContain("macabacus", read.WithMacabacusNoticeShown(false).ToJson());

        // The other With methods keep it.
        Assert.True(read.WithBounds(null).WithWrapFormula(false).WithEvaluateFunctions(true).MacabacusNoticeShown);
    }

    [Theory]
    [InlineData("{ \"macabacusNoticeShown\": true }", true)]
    [InlineData("{ \"schemaVersion\": 1, \"traceWindow\": 5, \"macabacusNoticeShown\": true }", true)]
    [InlineData("{ \"macabacusNoticeShown\": false }", false)]
    [InlineData("{ \"macabacusNoticeShown\": \"true\" }", false)]
    [InlineData("{ \"macabacusNoticeShown\": 1 }", false)]
    [InlineData("{ \"traceWindow\": { \"macabacusNoticeShown\": true } }", false)]
    [InlineData("{}", false)]
    public void Macabacus_notice_flag_is_read_from_the_top_level(string json, bool expected)
    {
        Assert.Equal(expected, TraceUiState.FromJson(json).MacabacusNoticeShown);
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
        Assert.False(state.EvaluateFunctions);
    }

    [Fact]
    public void A_file_from_before_evaluate_mode_reads_with_it_off()
    {
        var state = TraceUiState.FromJson(
            "{ \"schemaVersion\": 1, \"traceWindow\": { \"left\": 1, \"top\": 2, \"width\": 3, \"height\": 4, \"wrapFormula\": true } }");

        Assert.Equal(new WindowRect(1, 2, 3, 4), state.Bounds);
        Assert.True(state.WrapFormula);
        Assert.False(state.EvaluateFunctions);
        Assert.Contains("\"evaluateFunctions\": false", state.ToJson());
    }

    [Theory]
    [InlineData("true", true)]
    [InlineData("false", false)]
    [InlineData("1", false)]
    [InlineData("\"true\"", false)]
    [InlineData("null", false)]
    public void Evaluate_mode_is_on_only_for_true(string value, bool expected)
    {
        var state = TraceUiState.FromJson("{ \"traceWindow\": { \"evaluateFunctions\": " + value + " } }");

        Assert.Equal(expected, state.EvaluateFunctions);
        Assert.Null(state.Bounds);
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
        Assert.True(state.WithEvaluateFunctions(true).EvaluateFunctions);
        Assert.Equal(state.Bounds, state.WithEvaluateFunctions(true).Bounds);
        Assert.False(state.WithEvaluateFunctions(true).WrapFormula);
        Assert.True(state.WithEvaluateFunctions(true).WithBounds(null).WithWrapFormula(true).EvaluateFunctions);
        Assert.Null(state.WithBounds(null).Bounds);
        Assert.Equal(new WindowRect(5, 6, 7, 8), state.WithBounds(new WindowRect(5, 6, 7, 8)).Bounds);
    }
}
