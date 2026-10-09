using System;

namespace Modelwright.Core.Formatting;

/// <summary>
/// A cell's current value for a cycle: a number format code, a color (including <see cref="OleColor.NoFill"/>),
/// Excel's automatic font color (<see cref="Automatic"/>), or <see cref="Unknown"/> when it could not be read or is
/// mixed (for example, rich text in several colors).
/// </summary>
/// <remarks><c>default(CycleValue)</c> is <see cref="Unknown"/>. Number formats compare ordinally.</remarks>
public readonly struct CycleValue : IEquatable<CycleValue>
{
    private readonly string? _numberFormat;
    private readonly OleColor _color;
    private readonly bool _isColor;
    private readonly bool _isAutomatic;

    private CycleValue(string? numberFormat, OleColor color, bool isColor, bool isAutomatic)
    {
        _numberFormat = numberFormat;
        _color = color;
        _isColor = isColor;
        _isAutomatic = isAutomatic;
    }

    /// <summary>No usable value: mixed, unreadable, or not applicable.</summary>
    public static CycleValue Unknown => default;

    /// <summary>
    /// Excel's automatic font color (<c>Font.ColorIndex = xlColorIndexAutomatic</c>). It usually displays as black
    /// but is not equal to an explicit black, so undo can put it back as automatic. Not a <see cref="Color"/>: no
    /// cycle item has it.
    /// </summary>
    public static CycleValue Automatic { get; } = new CycleValue(null, default, false, true);

    /// <summary>True for <see cref="Unknown"/>.</summary>
    public bool IsUnknown => _numberFormat is null && !_isColor && !_isAutomatic;

    /// <summary>True when this is a number format code.</summary>
    public bool IsNumberFormat => _numberFormat is not null;

    /// <summary>True when this is a color (including <see cref="OleColor.NoFill"/>, but not <see cref="Automatic"/>).</summary>
    public bool IsColor => _isColor;

    /// <summary>True for <see cref="Automatic"/>.</summary>
    public bool IsAutomatic => _isAutomatic;

    /// <summary>The number format code, or null if this is not a number format.</summary>
    public string? NumberFormat => _numberFormat;

    /// <summary>The color, or null if this is not a color.</summary>
    public OleColor? Color => _isColor ? _color : (OleColor?)null;

    /// <summary>A number format value.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="code"/> is null.</exception>
    public static CycleValue FromNumberFormat(string code) =>
        new CycleValue(code ?? throw new ArgumentNullException(nameof(code)), default, false, false);

    /// <summary>A color value (<see cref="OleColor.NoFill"/> allowed).</summary>
    public static CycleValue FromColor(OleColor color) => new CycleValue(null, color, true, false);

    /// <inheritdoc />
    public bool Equals(CycleValue other) =>
        _isColor == other._isColor &&
        _isAutomatic == other._isAutomatic &&
        _color == other._color &&
        string.Equals(_numberFormat, other._numberFormat, StringComparison.Ordinal);

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is CycleValue other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() =>
        _numberFormat is not null
            ? StringComparer.Ordinal.GetHashCode(_numberFormat)
            : (_isColor ? _color.GetHashCode() : (_isAutomatic ? -2 : 0));

    /// <summary>Value equality.</summary>
    public static bool operator ==(CycleValue left, CycleValue right) => left.Equals(right);

    /// <summary>Value inequality.</summary>
    public static bool operator !=(CycleValue left, CycleValue right) => !left.Equals(right);

    /// <summary>The code, the color (<c>#RRGGBB</c> or <c>none</c>), <c>automatic</c>, or <c>(unknown)</c>.</summary>
    public override string ToString() =>
        _numberFormat ?? (_isColor ? _color.ToString() : (_isAutomatic ? "automatic" : "(unknown)"));
}
