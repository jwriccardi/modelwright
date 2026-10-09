using System;

namespace Modelwright.Core.Formatting;

/// <summary>A font or fill color cycle item: a name and a color, which may be <see cref="OleColor.NoFill"/> (fill cycles only).</summary>
public sealed class ColorItem : CycleItem
{
    /// <summary>Creates an item.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="name"/> is null.</exception>
    public ColorItem(string name, OleColor color)
        : base(name)
    {
        Color = color;
    }

    /// <summary>The color, or <see cref="OleColor.NoFill"/>.</summary>
    public OleColor Color { get; }

    /// <inheritdoc />
    public override CycleValue Value => CycleValue.FromColor(Color);

    /// <inheritdoc />
    public override string ToString() => $"{Name}: {Color}";
}
