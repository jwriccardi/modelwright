using System;
using System.Collections.Generic;
using System.Linq;
using ExcelModelingToolkit.Core.Json;
using Xunit;

namespace ExcelModelingToolkit.Core.Tests.Json;

public class JsonTests
{
    [Fact]
    public void Parses_every_value_type()
    {
        var root = Assert.IsType<JsonObject>(JsonParser.Parse(
            "{ \"s\": \"x\", \"t\": true, \"f\": false, \"n\": null, \"i\": -12, \"d\": 1.5e-3, \"a\": [1, [], {}], \"o\": { \"k\": \"v\" } }"));

        Assert.Equal(new[] { "s", "t", "f", "n", "i", "d", "a", "o" }, root.Properties.Select(p => p.Key));
        Assert.Equal("x", Get(root, "s"));
        Assert.Equal(true, Get(root, "t"));
        Assert.Equal(false, Get(root, "f"));
        Assert.Null(Get(root, "n"));
        Assert.Equal("-12", Assert.IsType<JsonNumber>(Get(root, "i")).Text);
        Assert.Equal("1.5e-3", Get(root, "d")!.ToString());
        var array = Assert.IsType<List<object?>>(Get(root, "a"));
        Assert.Equal(3, array.Count);
        Assert.Empty(Assert.IsType<List<object?>>(array[1]));
        Assert.Empty(Assert.IsType<JsonObject>(array[2]).Properties);
        Assert.Equal("v", Get(Assert.IsType<JsonObject>(Get(root, "o")), "k"));
        Assert.False(root.TryGet("missing", out _));
    }

    [Theory]
    [InlineData("\"a\\\"b\"", "a\"b")]
    [InlineData("\"a\\\\b\"", "a\\b")]
    [InlineData("\"a\\/b\"", "a/b")]
    [InlineData("\"\\b\\f\\n\\r\\t\"", "\b\f\n\r\t")]
    [InlineData("\"\\u2013\"", "\u2013")]
    [InlineData("\"\\u00e9\\u00E9\"", "\u00e9\u00e9")]
    [InlineData("\"\u2013\u20ac\"", "\u2013\u20ac")]
    [InlineData("\"\"", "")]
    public void Parses_string_escapes(string json, string expected)
    {
        Assert.Equal(expected, JsonParser.Parse(json));
    }

    [Theory]
    [InlineData("0")]
    [InlineData("-0")]
    [InlineData("10000")]
    [InlineData("3.25")]
    [InlineData("1E10")]
    [InlineData("2e+2")]
    public void Parses_numbers_verbatim(string json)
    {
        Assert.Equal(json, Assert.IsType<JsonNumber>(JsonParser.Parse(json)).Text);
    }

    [Theory]
    [InlineData("", "line 1, column 1: unexpected end of text")]
    [InlineData("   ", "line 1, column 4: unexpected end of text")]
    [InlineData("{", "line 1, column 2: expected a property name")]
    [InlineData("{\n  \"a\": 1,\n}", "line 3, column 1: expected a property name")]
    [InlineData("[1,]", "line 1, column 4: unexpected character ']'")]
    [InlineData("[1 2]", "line 1, column 4: expected ']' but found '2'")]
    [InlineData("[1", "line 1, column 3: expected ']' but found the end of the text")]
    [InlineData("{\"a\" 1}", "line 1, column 6: expected ':'")]
    [InlineData("{\"a\": 1 \"b\": 2}", "line 1, column 9: expected '}'")]
    [InlineData("{\"a\": 1, \"a\": 2}", "line 1, column 10: duplicate property \"a\"")]
    [InlineData("{a: 1}", "expected a property name in double quotes")]
    [InlineData("\"abc", "unterminated string")]
    [InlineData("\"abc\\", "unterminated string")]
    [InlineData("\"a\tb\"", "control character in a string")]
    [InlineData("\"\\x\"", "invalid escape '\\x'")]
    [InlineData("\"\\u12\"", "\\u must be followed by four hex digits")]
    [InlineData("\"\\u12G4\"", "\\u must be followed by four hex digits")]
    [InlineData("01", "unexpected text after the end")]
    [InlineData("-", "invalid number.")]
    [InlineData("1.", "expected digits after '.'")]
    [InlineData("1e", "expected digits in the exponent")]
    [InlineData("+1", "unexpected character '+'")]
    [InlineData("tru", "expected a value such as true")]
    [InlineData("nul", "expected a value such as null")]
    [InlineData("falsy", "expected a value such as false")]
    [InlineData("{} {}", "unexpected text after the end")]
    [InlineData("// comment\n{}", "unexpected character '/'")]
    public void Rejects_invalid_json_with_a_position(string json, string expected)
    {
        var ex = Assert.Throws<FormatException>(() => JsonParser.Parse(json));

        Assert.Contains(expected, ex.Message);
        Assert.StartsWith("line ", ex.Message);
    }

    [Fact]
    public void Rejects_excessive_nesting()
    {
        var deep = new string('[', JsonParser.MaxDepth + 1) + new string(']', JsonParser.MaxDepth + 1);
        var ok = new string('[', JsonParser.MaxDepth) + new string(']', JsonParser.MaxDepth);

        Assert.Contains("nesting is deeper than 64", Assert.Throws<FormatException>(() => JsonParser.Parse(deep)).Message);
        Assert.IsType<List<object?>>(JsonParser.Parse(ok));
        Assert.Throws<ArgumentNullException>(() => JsonParser.Parse(null!));
    }

    [Fact]
    public void Skips_a_byte_order_mark_and_whitespace()
    {
        Assert.Equal(true, JsonParser.Parse("\uFEFF \r\n\t true \r\n"));
    }

    [Fact]
    public void Writer_output_parses_back_to_the_same_tree()
    {
        var root = new JsonObject();
        root.Add("text", "quote \" backslash \\ newline \n return \r tab \t bell \u0007 dash \u2013 slash /");
        root.Add("number", 42);
        root.Add("raw", new JsonNumber("1.5e3"));
        root.Add("yes", true);
        root.Add("no", false);
        root.Add("nothing", null);
        root.Add("empty", new JsonObject());
        root.Add("list", new List<object?>());
        var nested = new JsonObject();
        nested.Add("inner", new List<object?> { 1, "two" });
        root.Add("items", new List<object?> { nested, "scalar", new List<object?> { true } });

        var json = JsonWriter.Write(root);
        var parsed = Assert.IsType<JsonObject>(JsonParser.Parse(json));

        Assert.Equal(json, JsonWriter.Write(parsed));
        Assert.Equal(root.Properties.Select(p => p.Key), parsed.Properties.Select(p => p.Key));
        Assert.Equal(Get(root, "text"), Get(parsed, "text"));
        Assert.Contains("\\u0007", json);
        Assert.Contains("\u2013", json);
        Assert.Contains("\"empty\": {},", json);
        Assert.Contains("\"list\": [],", json);
    }

    [Fact]
    public void Writer_inlines_scalar_objects_in_arrays_only()
    {
        var item = new JsonObject();
        item.Add("name", "Blue");
        item.Add("n", 1);
        var root = new JsonObject();
        root.Add("map", item);
        root.Add("items", new List<object?> { item });

        var json = JsonWriter.Write(root);

        Assert.Equal(
            "{\r\n  \"map\": {\r\n    \"name\": \"Blue\",\r\n    \"n\": 1\r\n  },\r\n  \"items\": [\r\n    { \"name\": \"Blue\", \"n\": 1 }\r\n  ]\r\n}\r\n",
            json);
    }

    [Fact]
    public void Writer_rejects_unknown_node_types()
    {
        Assert.Throws<ArgumentException>(() => JsonWriter.Write(1.5));
    }

    [Fact]
    public void JsonObject_rejects_duplicate_names()
    {
        var obj = new JsonObject();
        obj.Add("a", 1);

        Assert.Throws<ArgumentException>(() => obj.Add("a", 2));
        Assert.True(obj.Contains("a"));
        Assert.False(obj.Contains("A"));
    }

    private static object? Get(JsonObject obj, string name)
    {
        Assert.True(obj.TryGet(name, out var value));
        return value;
    }
}
