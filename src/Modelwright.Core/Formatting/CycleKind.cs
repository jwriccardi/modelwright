namespace Modelwright.Core.Formatting;

/// <summary>The cell property a <see cref="CycleDefinition"/> changes.</summary>
public enum CycleKind
{
    /// <summary>The number format code (<c>Range.NumberFormat</c>). Items are <see cref="NumberFormatItem"/>s.</summary>
    NumberFormat,

    /// <summary>The font color (<c>Range.Font.Color</c>). Items are <see cref="ColorItem"/>s, never "no fill".</summary>
    FontColor,

    /// <summary>
    /// The fill (<c>Range.Interior.Color</c>, or <c>Interior.Pattern = xlNone</c> for <see cref="OleColor.NoFill"/>).
    /// Items are <see cref="ColorItem"/>s.
    /// </summary>
    FillColor,
}
