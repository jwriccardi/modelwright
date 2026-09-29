using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcelModelingToolkit.Core.Json;

/// <summary>
/// A strict RFC 8259 JSON parser into <see cref="JsonObject"/>, <see cref="List{T}"/>, <see cref="string"/>,
/// <see cref="bool"/>, <see cref="JsonNumber"/> and null. Rejects duplicate property names, comments, trailing
/// commas, unpaired surrogate escapes and nesting deeper than <see cref="MaxDepth"/>. A leading byte order mark
/// is ignored.
/// </summary>
internal sealed class JsonParser
{
    /// <summary>Deepest nesting of objects and arrays accepted.</summary>
    public const int MaxDepth = 64;

    private readonly string _text;
    private int _position;
    private int _depth;

    private JsonParser(string text) => _text = text;

    /// <summary>Parses one JSON value that makes up the whole of <paramref name="text"/>.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="text"/> is null.</exception>
    /// <exception cref="FormatException">The text is not valid JSON; the message gives the line and column.</exception>
    public static object? Parse(string text)
    {
        if (text is null)
        {
            throw new ArgumentNullException(nameof(text));
        }

        var parser = new JsonParser(text);
        if (text.Length > 0 && text[0] == '\uFEFF')
        {
            parser._position = 1;
        }

        parser.SkipWhitespace();
        var value = parser.ParseValue();
        parser.SkipWhitespace();
        if (parser._position < text.Length)
        {
            throw parser.Error("unexpected text after the end of the JSON value.");
        }

        return value;
    }

    private object? ParseValue()
    {
        if (_position >= _text.Length)
        {
            throw Error("unexpected end of text; expected a value.");
        }

        var c = _text[_position];
        switch (c)
        {
            case '{':
                return ParseObject();
            case '[':
                return ParseArray();
            case '"':
                return ParseString();
            case 't':
                return ParseLiteral("true", true);
            case 'f':
                return ParseLiteral("false", false);
            case 'n':
                return ParseLiteral("null", null);
            default:
                if (c == '-' || (c >= '0' && c <= '9'))
                {
                    return ParseNumber();
                }

                throw Error($"unexpected character '{c}'; expected a value.");
        }
    }

    private JsonObject ParseObject()
    {
        Enter();
        _position++; // '{'
        var result = new JsonObject();
        SkipWhitespace();
        if (Peek() == '}')
        {
            _position++;
            _depth--;
            return result;
        }

        while (true)
        {
            if (Peek() != '"')
            {
                throw Error("expected a property name in double quotes.");
            }

            var nameStart = _position;
            var name = ParseString();
            if (result.Contains(name))
            {
                _position = nameStart;
                throw Error($"duplicate property \"{name}\".");
            }

            SkipWhitespace();
            Expect(':');
            SkipWhitespace();
            result.Add(name, ParseValue());
            SkipWhitespace();
            if (Peek() == ',')
            {
                _position++;
                SkipWhitespace();
                continue;
            }

            Expect('}');
            _depth--;
            return result;
        }
    }

    private List<object?> ParseArray()
    {
        Enter();
        _position++; // '['
        var result = new List<object?>();
        SkipWhitespace();
        if (Peek() == ']')
        {
            _position++;
            _depth--;
            return result;
        }

        while (true)
        {
            result.Add(ParseValue());
            SkipWhitespace();
            if (Peek() == ',')
            {
                _position++;
                SkipWhitespace();
                continue;
            }

            Expect(']');
            _depth--;
            return result;
        }
    }

    private string ParseString()
    {
        _position++; // opening quote
        var sb = new StringBuilder();
        while (true)
        {
            if (_position >= _text.Length)
            {
                throw Error("unterminated string.");
            }

            var c = _text[_position];
            if (c == '"')
            {
                _position++;
                return sb.ToString();
            }

            if (c < ' ')
            {
                throw Error("control character in a string; write it as an escape such as \\n or \\t.");
            }

            if (c != '\\')
            {
                sb.Append(c);
                _position++;
                continue;
            }

            var escapeStart = _position;
            _position++;
            if (_position >= _text.Length)
            {
                throw Error("unterminated string.");
            }

            var escape = _text[_position];
            switch (escape)
            {
                case '"':
                case '\\':
                case '/':
                    sb.Append(escape);
                    break;
                case 'b':
                    sb.Append('\b');
                    break;
                case 'f':
                    sb.Append('\f');
                    break;
                case 'n':
                    sb.Append('\n');
                    break;
                case 'r':
                    sb.Append('\r');
                    break;
                case 't':
                    sb.Append('\t');
                    break;
                case 'u':
                    var code = ReadHex4();
                    if (char.IsLowSurrogate(code))
                    {
                        _position = escapeStart;
                        throw Error($"\\u{(int)code:X4} is a lone low surrogate; it must follow a high surrogate escape (\\uD800-\\uDBFF).");
                    }

                    if (char.IsHighSurrogate(code))
                    {
                        // A character outside the Basic Multilingual Plane: both halves must be escaped, in order.
                        var low = _position + 2 < _text.Length && _text[_position + 1] == '\\' && _text[_position + 2] == 'u'
                            ? PeekHex4(_position + 2)
                            : null;
                        if (low is null || !char.IsLowSurrogate(low.Value))
                        {
                            _position = escapeStart;
                            throw Error($"\\u{(int)code:X4} is a lone high surrogate; it must be followed by a low surrogate escape (\\uDC00-\\uDFFF).");
                        }

                        sb.Append(code).Append(low.Value);
                        _position += 6;
                        break;
                    }

                    sb.Append(code);
                    break;
                default:
                    throw Error($"invalid escape '\\{escape}' in a string.");
            }

            _position++;
        }
    }

    /// <summary>Reads the four hex digits after the <c>u</c> at the current position, leaving the position on the last digit.</summary>
    private char ReadHex4()
    {
        var code = PeekHex4(_position) ?? throw Error("\\u must be followed by four hex digits.");
        _position += 4;
        return code;
    }

    /// <summary>The character whose four hex digits follow the <c>u</c> at <paramref name="uPosition"/>, or null if there are not four hex digits.</summary>
    private char? PeekHex4(int uPosition) =>
        uPosition + 4 < _text.Length &&
        int.TryParse(_text.Substring(uPosition + 1, 4), NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out var code)
            ? (char)code
            : null;

    private JsonNumber ParseNumber()
    {
        var start = _position;
        if (Peek() == '-')
        {
            _position++;
        }

        if (Peek() == '0')
        {
            _position++;
        }
        else if (!ReadDigits())
        {
            throw Error("invalid number.");
        }

        if (Peek() == '.')
        {
            _position++;
            if (!ReadDigits())
            {
                throw Error("invalid number: expected digits after '.'.");
            }
        }

        if (Peek() == 'e' || Peek() == 'E')
        {
            _position++;
            if (Peek() == '+' || Peek() == '-')
            {
                _position++;
            }

            if (!ReadDigits())
            {
                throw Error("invalid number: expected digits in the exponent.");
            }
        }

        return new JsonNumber(_text.Substring(start, _position - start));
    }

    private bool ReadDigits()
    {
        var start = _position;
        while (_position < _text.Length && _text[_position] >= '0' && _text[_position] <= '9')
        {
            _position++;
        }

        return _position > start;
    }

    private object? ParseLiteral(string literal, object? value)
    {
        if (string.CompareOrdinal(_text, _position, literal, 0, literal.Length) != 0)
        {
            throw Error($"unexpected text; expected a value such as {literal}.");
        }

        _position += literal.Length;
        return value;
    }

    private void Enter()
    {
        if (++_depth > MaxDepth)
        {
            throw Error($"nesting is deeper than {MaxDepth} levels.");
        }
    }

    private void Expect(char expected)
    {
        if (Peek() != expected)
        {
            var found = _position < _text.Length ? $"'{_text[_position]}'" : "the end of the text";
            throw Error($"expected '{expected}' but found {found}.");
        }

        _position++;
    }

    private char Peek() => _position < _text.Length ? _text[_position] : '\0';

    private void SkipWhitespace()
    {
        while (_position < _text.Length)
        {
            var c = _text[_position];
            if (c != ' ' && c != '\t' && c != '\n' && c != '\r')
            {
                return;
            }

            _position++;
        }
    }

    private FormatException Error(string message)
    {
        var line = 1;
        var column = 1;
        for (var i = 0; i < _position && i < _text.Length; i++)
        {
            if (_text[i] == '\n')
            {
                line++;
                column = 1;
            }
            else
            {
                column++;
            }
        }

        return new FormatException($"line {line}, column {column}: {message}");
    }
}
