using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace ExcelModelingToolkit.Core.Json;

/// <summary>
/// Writes a node tree (see <see cref="JsonObject"/>) as indented JSON for people to edit: two-space indents,
/// CRLF line ends, a final line end, and non-ASCII characters written as themselves (the file is UTF-8).
/// An object inside an array whose values are all scalars goes on one line.
/// </summary>
internal static class JsonWriter
{
    private const string NewLine = "\r\n";
    private const string Indent = "  ";

    /// <summary>Writes <paramref name="value"/>.</summary>
    /// <exception cref="ArgumentException">The tree contains a node type the writer does not know.</exception>
    public static string Write(object? value)
    {
        var sb = new StringBuilder();
        WriteValue(sb, value, 0, inArray: false);
        return sb.Append(NewLine).ToString();
    }

    private static void WriteValue(StringBuilder sb, object? value, int depth, bool inArray)
    {
        switch (value)
        {
            case JsonObject obj:
                WriteObject(sb, obj, depth, inArray);
                break;
            case List<object?> array:
                WriteArray(sb, array, depth);
                break;
            default:
                WriteScalar(sb, value);
                break;
        }
    }

    private static void WriteObject(StringBuilder sb, JsonObject obj, int depth, bool inArray)
    {
        if (obj.Properties.Count == 0)
        {
            sb.Append("{}");
            return;
        }

        if (inArray && obj.Properties.All(p => IsScalar(p.Value)))
        {
            sb.Append("{ ");
            for (var i = 0; i < obj.Properties.Count; i++)
            {
                if (i > 0)
                {
                    sb.Append(", ");
                }

                WriteString(sb, obj.Properties[i].Key);
                sb.Append(": ");
                WriteScalar(sb, obj.Properties[i].Value);
            }

            sb.Append(" }");
            return;
        }

        sb.Append('{').Append(NewLine);
        for (var i = 0; i < obj.Properties.Count; i++)
        {
            AppendIndent(sb, depth + 1);
            WriteString(sb, obj.Properties[i].Key);
            sb.Append(": ");
            WriteValue(sb, obj.Properties[i].Value, depth + 1, inArray: false);
            if (i < obj.Properties.Count - 1)
            {
                sb.Append(',');
            }

            sb.Append(NewLine);
        }

        AppendIndent(sb, depth);
        sb.Append('}');
    }

    private static void WriteArray(StringBuilder sb, List<object?> array, int depth)
    {
        if (array.Count == 0)
        {
            sb.Append("[]");
            return;
        }

        sb.Append('[').Append(NewLine);
        for (var i = 0; i < array.Count; i++)
        {
            AppendIndent(sb, depth + 1);
            WriteValue(sb, array[i], depth + 1, inArray: true);
            if (i < array.Count - 1)
            {
                sb.Append(',');
            }

            sb.Append(NewLine);
        }

        AppendIndent(sb, depth);
        sb.Append(']');
    }

    private static bool IsScalar(object? value) => !(value is JsonObject) && !(value is List<object?>);

    private static void WriteScalar(StringBuilder sb, object? value)
    {
        switch (value)
        {
            case null:
                sb.Append("null");
                break;
            case string text:
                WriteString(sb, text);
                break;
            case bool flag:
                sb.Append(flag ? "true" : "false");
                break;
            case int number:
                sb.Append(number.ToString(CultureInfo.InvariantCulture));
                break;
            case JsonNumber number:
                sb.Append(number.Text);
                break;
            default:
                throw new ArgumentException($"Cannot write a {value.GetType().Name} as JSON.", nameof(value));
        }
    }

    private static void WriteString(StringBuilder sb, string text)
    {
        sb.Append('"');
        foreach (var c in text)
        {
            switch (c)
            {
                case '"':
                    sb.Append("\\\"");
                    break;
                case '\\':
                    sb.Append("\\\\");
                    break;
                case '\n':
                    sb.Append("\\n");
                    break;
                case '\r':
                    sb.Append("\\r");
                    break;
                case '\t':
                    sb.Append("\\t");
                    break;
                default:
                    if (c < ' ')
                    {
                        sb.Append("\\u").Append(((int)c).ToString("x4", CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        sb.Append(c);
                    }

                    break;
            }
        }

        sb.Append('"');
    }

    private static void AppendIndent(StringBuilder sb, int depth)
    {
        for (var i = 0; i < depth; i++)
        {
            sb.Append(Indent);
        }
    }
}
