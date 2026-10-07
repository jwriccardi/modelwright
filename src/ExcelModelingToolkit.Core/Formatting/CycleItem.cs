using System;

namespace ExcelModelingToolkit.Core.Formatting;

/// <summary>One step of a <see cref="CycleDefinition"/>: a <see cref="NumberFormatItem"/> or a <see cref="ColorItem"/>.</summary>
public abstract class CycleItem
{
    private protected CycleItem(string name)
    {
        Name = name ?? throw new ArgumentNullException(nameof(name));
    }

    /// <summary>The name shown to people, e.g. <c>Comma 1 Dec Lg Align</c> or <c>Blue</c>.</summary>
    public string Name { get; }

    /// <summary>The value this item applies, in the form the cycle engine compares against the cell.</summary>
    public abstract CycleValue Value { get; }
}
