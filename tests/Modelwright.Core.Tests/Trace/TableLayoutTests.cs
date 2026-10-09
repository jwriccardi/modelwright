using System;
using System.Linq;
using Modelwright.Core.Trace;
using Modelwright.Core.Undo;
using Xunit;

namespace Modelwright.Core.Tests.Trace;

public class TableLayoutTests
{
    private static readonly FormulaContext Context = new FormulaContext("Model.xlsx", "Calc");

    // Sales on B2:E12: header row 2, data rows 3-11, totals row 12; columns Region, Q1, Q2, Amount (B to E).
    private static readonly TableLayout Sales = new TableLayout(
        new CellRect(2, 2, 11, 4), hasHeaders: true, hasTotals: true, new[] { "Region", "Q1", "Q2", "Amount" });

    private static string? Resolve(string formula, int? row = 5, TableLayout? table = null)
    {
        var reference = FormulaParser.Parse(formula, Context).References.Single();
        var rect = (table ?? Sales).Resolve(reference.TableSpecifiers, reference.TableColumns, row, out var error);
        return rect?.Address ?? "error: " + error;
    }

    [Theory]
    [InlineData("=SUM(Sales[Amount])", "$E$3:$E$11")]
    [InlineData("=SUM(Sales[amount])", "$E$3:$E$11")]
    [InlineData("=ROWS(Sales[#All])", "$B$2:$E$12")]
    [InlineData("=ROWS(Sales[#Data])", "$B$3:$E$11")]
    [InlineData("=Sales[[#Headers],[Q1]]", "$C$2")]
    [InlineData("=Sales[#Totals]", "$B$12:$E$12")]
    [InlineData("=SUM(Sales[[#Headers],[#Data],[Q2]])", "$D$2:$D$11")]
    [InlineData("=SUM(Sales[[#Data],[#Totals],[Q2]])", "$D$3:$D$12")]
    [InlineData("=SUM(Sales[[#Data],[Q1]:[Q2]])", "$C$3:$D$11")]
    [InlineData("=SUM(Sales[[Q2]:[Q1]])", "$C$3:$D$11")]
    [InlineData("=Sales[@Amount]", "$E$5")]
    [InlineData("=Sales[[#This Row],[Amount]]", "$E$5")]
    [InlineData("=[@Amount]", "$E$5")]
    [InlineData("=Sales[@[Q1]:[Q2]]", "$C$5:$D$5")]
    public void Structured_references_resolve_to_their_cells(string formula, string address)
    {
        Assert.Equal(address, Resolve(formula));
    }

    [Theory]
    [InlineData("=Sales[@Amount]", 2)]
    [InlineData("=Sales[@Amount]", 12)]
    [InlineData("=Sales[@Amount]", null)]
    public void This_row_outside_the_data_is_an_error(string formula, int? row)
    {
        Assert.Equal("error: this row is outside the table's data rows", Resolve(formula, row));
    }

    [Fact]
    public void Unknown_columns_and_hidden_rows_are_errors()
    {
        var bare = new TableLayout(new CellRect(2, 2, 9, 4), hasHeaders: false, hasTotals: false, new[] { "Region", "Q1", "Q2", "Amount" });

        Assert.Equal("error: the table has no column 'Q9'", Resolve("=SUM(Sales[Q9])"));
        Assert.Equal("error: the table's header row is hidden", Resolve("=Sales[#Headers]", table: bare));
        Assert.Equal("error: the table's totals row is hidden", Resolve("=Sales[#Totals]", table: bare));
        Assert.Equal("$B$2:$E$10", Resolve("=ROWS(Sales[#Data])", table: bare));
    }

    [Fact]
    public void Invalid_combinations_are_errors()
    {
        Assert.Null(Sales.Resolve(new[] { "#Headers", "#Totals" }, new string[0], 5, out var error));
        Assert.Equal("the header and totals rows are not next to each other", error);
        Assert.Null(Sales.Resolve(new[] { "@", "#Totals" }, new string[0], 5, out error));
        Assert.Equal("this row cannot be combined with other item specifiers", error);
        Assert.Null(Sales.Resolve(new[] { "#Nope" }, new string[0], 5, out error));
        Assert.Equal("'#Nope' is not a table item specifier", error);
        Assert.Null(Sales.Resolve(new string[0], new[] { "Q1", "Q2", "Amount" }, 5, out error));
        Assert.Equal("a structured reference can name at most two columns (a span)", error);
    }

    [Fact]
    public void A_table_with_only_a_header_has_no_data()
    {
        var empty = new TableLayout(new CellRect(2, 2, 1, 1), hasHeaders: true, hasTotals: false, new[] { "Only" });

        Assert.Equal(3, empty.FirstDataRow);
        Assert.Equal(2, empty.LastDataRow);
        Assert.Null(empty.Resolve(new string[0], new string[0], 2, out var error));
        Assert.Equal("the table has no data rows", error);
        Assert.Equal(new CellRect(2, 2, 1, 1), empty.Resolve(new[] { "#Headers" }, new string[0], null, out _));
    }

    [Fact]
    public void Layouts_are_checked()
    {
        Assert.Throws<ArgumentNullException>(() => new TableLayout(new CellRect(1, 1, 2, 2), true, false, null!));
        Assert.Throws<ArgumentException>(() => new TableLayout(new CellRect(1, 1, 2, 2), true, false, new[] { "A" }));
        Assert.Throws<ArgumentException>(() => new TableLayout(new CellRect(1, 1, 1, 1), true, true, new[] { "A" }));
        Assert.Throws<ArgumentNullException>(() => Sales.Resolve(null!, new string[0], 1, out _));
        Assert.Throws<ArgumentNullException>(() => Sales.Resolve(new string[0], null!, 1, out _));
        Assert.Equal(new[] { "Region", "Q1", "Q2", "Amount" }, Sales.Columns);
        Assert.True(Sales.HasHeaders);
        Assert.True(Sales.HasTotals);
        Assert.Equal(new CellRect(2, 2, 11, 4), Sales.Range);
    }
}
