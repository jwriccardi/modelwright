using System;

namespace ExcelModelingToolkit.Core.Formatting;

/// <summary>A number-format cycle item: a name and an Excel number format code (en-US syntax, as <c>Range.NumberFormat</c>).</summary>
public sealed class NumberFormatItem : CycleItem
{
    /// <summary>Creates an item.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="name"/> or <paramref name="code"/> is null.</exception>
    public NumberFormatItem(string name, string code)
        : base(name)
    {
        Code = code ?? throw new ArgumentNullException(nameof(code));
    }

    /// <summary>The number format code, e.g. <c>#,##0;(#,##0);"–";@</c>.</summary>
    public string Code { get; }

    /// <inheritdoc />
    public override CycleValue Value => CycleValue.FromNumberFormat(Code);

    /// <inheritdoc />
    public override string ToString() => $"{Name}: {Code}";
}
