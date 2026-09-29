using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelModelingToolkit.Core.Formatting;

/// <summary>
/// An RGB color in Excel's OLE color form (<c>Range.Font.Color</c>, <c>Range.Interior.Color</c>):
/// the integer <c>R + G*256 + B*65536</c>, i.e. BGR byte order. Also represents "no fill"
/// (<see cref="NoFill"/>), which has no color value: it is applied as <c>Interior.Pattern = xlNone</c>.
/// </summary>
/// <remarks><c>default(OleColor)</c> is black.</remarks>
public readonly struct OleColor : IEquatable<OleColor>
{
    /// <summary>The largest OLE RGB value (white, <c>0xFFFFFF</c>).</summary>
    public const int MaxOleValue = 0xFFFFFF;

    private const int NoFillValue = -1;

    private static readonly Regex RgbPattern = new Regex(
        @"^\s*rgb\s*\(\s*([0-9]{1,3})\s*,\s*([0-9]{1,3})\s*,\s*([0-9]{1,3})\s*\)\s*$",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex HexPattern = new Regex(
        @"^\s*#([0-9A-Fa-f]{6})\s*$",
        RegexOptions.CultureInvariant);

    private readonly int _value;

    private OleColor(int value) => _value = value;

    /// <summary>The "no fill" sentinel (a cell with no interior color). It is not equal to any RGB color.</summary>
    /// <remarks>
    /// Excel reports a no-fill cell's <c>Interior.Color</c> as white (<c>0xFFFFFF</c>), the same as a white
    /// fill. Adapters reading a cell must check <c>Interior.Pattern == xlNone</c> (<c>-4142</c>) first and map
    /// it to <see cref="NoFill"/>; only otherwise is <c>Interior.Color</c> the fill color.
    /// </remarks>
    public static OleColor NoFill { get; } = new OleColor(NoFillValue);

    /// <summary>True for <see cref="NoFill"/>.</summary>
    public bool IsNoFill => _value == NoFillValue;

    /// <summary>Red component, 0-255.</summary>
    /// <exception cref="InvalidOperationException">This is <see cref="NoFill"/>.</exception>
    public byte R => (byte)(RequireColor() & 0xFF);

    /// <summary>Green component, 0-255.</summary>
    /// <exception cref="InvalidOperationException">This is <see cref="NoFill"/>.</exception>
    public byte G => (byte)((RequireColor() >> 8) & 0xFF);

    /// <summary>Blue component, 0-255.</summary>
    /// <exception cref="InvalidOperationException">This is <see cref="NoFill"/>.</exception>
    public byte B => (byte)((RequireColor() >> 16) & 0xFF);

    /// <summary>The Excel OLE color integer, <c>R + G*256 + B*65536</c>.</summary>
    /// <exception cref="InvalidOperationException">This is <see cref="NoFill"/>, which has no color value.</exception>
    public int OleValue => RequireColor();

    /// <summary>Creates a color from red, green and blue components (each 0-255).</summary>
    /// <exception cref="ArgumentOutOfRangeException">A component is outside 0-255.</exception>
    public static OleColor FromRgb(int r, int g, int b)
    {
        CheckComponent(r, nameof(r));
        CheckComponent(g, nameof(g));
        CheckComponent(b, nameof(b));
        return new OleColor(r | (g << 8) | (b << 16));
    }

    /// <summary>Creates a color from an Excel OLE color integer (0 to <see cref="MaxOleValue"/>).</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is negative or above <see cref="MaxOleValue"/>.</exception>
    public static OleColor FromOle(int oleValue)
    {
        if (oleValue < 0 || oleValue > MaxOleValue)
        {
            throw new ArgumentOutOfRangeException(
                nameof(oleValue),
                oleValue,
                $"An OLE RGB color must be between 0 and {MaxOleValue} (0xFFFFFF).");
        }

        return new OleColor(oleValue);
    }

    /// <summary>
    /// Parses <c>rgb(r,g,b)</c> (case-insensitive, spaces allowed), <c>#RRGGBB</c> (hex, either case),
    /// or <c>none</c> for <see cref="NoFill"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="text"/> is null.</exception>
    /// <exception cref="FormatException">The text is not in one of those forms, or a component exceeds 255.</exception>
    public static OleColor Parse(string text)
    {
        if (text is null)
        {
            throw new ArgumentNullException(nameof(text));
        }

        if (string.Equals(text.Trim(), "none", StringComparison.OrdinalIgnoreCase))
        {
            return NoFill;
        }

        var hex = HexPattern.Match(text);
        if (hex.Success)
        {
            var rrggbb = int.Parse(hex.Groups[1].Value, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture);
            return FromRgb((rrggbb >> 16) & 0xFF, (rrggbb >> 8) & 0xFF, rrggbb & 0xFF);
        }

        var rgb = RgbPattern.Match(text);
        if (rgb.Success)
        {
            var r = int.Parse(rgb.Groups[1].Value, CultureInfo.InvariantCulture);
            var g = int.Parse(rgb.Groups[2].Value, CultureInfo.InvariantCulture);
            var b = int.Parse(rgb.Groups[3].Value, CultureInfo.InvariantCulture);
            if (r > 255 || g > 255 || b > 255)
            {
                throw new FormatException($"Invalid color \"{text}\": each rgb() component must be 0-255.");
            }

            return FromRgb(r, g, b);
        }

        throw new FormatException($"Invalid color \"{text}\": expected rgb(r,g,b), #RRGGBB or none.");
    }

    /// <summary>Returns <c>#RRGGBB</c> (upper-case hex), or <c>none</c> for <see cref="NoFill"/>. <see cref="Parse"/> accepts it back.</summary>
    public override string ToString() =>
        IsNoFill ? "none" : string.Format(CultureInfo.InvariantCulture, "#{0:X2}{1:X2}{2:X2}", R, G, B);

    /// <inheritdoc />
    public bool Equals(OleColor other) => _value == other._value;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OleColor other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => _value;

    /// <summary>Value equality.</summary>
    public static bool operator ==(OleColor left, OleColor right) => left.Equals(right);

    /// <summary>Value inequality.</summary>
    public static bool operator !=(OleColor left, OleColor right) => !left.Equals(right);

    private int RequireColor()
    {
        if (IsNoFill)
        {
            throw new InvalidOperationException("\"No fill\" has no RGB/OLE color value; clear the interior pattern instead.");
        }

        return _value;
    }

    private static void CheckComponent(int value, string name)
    {
        if (value < 0 || value > 255)
        {
            throw new ArgumentOutOfRangeException(name, value, "A color component must be between 0 and 255.");
        }
    }
}
