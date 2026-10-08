using System;
using System.Globalization;

namespace ExcelModelingToolkit.Core.Trace;

/// <summary>
/// Where a precedent is written: the cell whose formula holds the reference (its owner), that formula as it was
/// parsed, and the reference's position in it. F2 in the Trace In window edits this text (see
/// <see cref="ReferenceEdit"/>).
/// </summary>
public sealed class ReferenceSpan
{
    /// <summary>Creates a span.</summary>
    /// <param name="workbook">The owner cell's workbook name (<c>Model.xlsx</c>).</param>
    /// <param name="sheet">The owner cell's sheet name.</param>
    /// <param name="address">The owner cell's sheet-local address (<c>B5</c>).</param>
    /// <param name="formula">The owner cell's formula as it was parsed, starting with <c>=</c>.</param>
    /// <param name="start">The reference's 0-based position in <paramref name="formula"/>.</param>
    /// <param name="length">The reference's length (at least 1).</param>
    /// <exception cref="ArgumentNullException">A string argument is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">The reference is not inside <paramref name="formula"/>.</exception>
    public ReferenceSpan(string workbook, string sheet, string address, string formula, int start, int length)
    {
        Workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
        Sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
        Address = address ?? throw new ArgumentNullException(nameof(address));
        Formula = formula ?? throw new ArgumentNullException(nameof(formula));
        if (start < 0 || length < 1 || start + length > formula.Length)
        {
            throw new ArgumentOutOfRangeException(nameof(start), start, "The reference must lie inside the formula.");
        }

        Start = start;
        Length = length;
    }

    /// <summary>The owner cell's workbook name.</summary>
    public string Workbook { get; }

    /// <summary>The owner cell's sheet name.</summary>
    public string Sheet { get; }

    /// <summary>The owner cell's sheet-local address.</summary>
    public string Address { get; }

    /// <summary>The owner cell's formula as it was parsed (starting with <c>=</c>).</summary>
    public string Formula { get; }

    /// <summary>The reference's 0-based position in <see cref="Formula"/>.</summary>
    public int Start { get; }

    /// <summary>The reference's length.</summary>
    public int Length { get; }

    /// <summary>The reference as written.</summary>
    public string Text => Formula.Substring(Start, Length);

    /// <summary><c>Sheet!B5[start+length]</c>, for the diagnostics log.</summary>
    public override string ToString() =>
        Sheet + "!" + Address + "[" + Start.ToString(CultureInfo.InvariantCulture) + "+" +
        Length.ToString(CultureInfo.InvariantCulture) + "]";
}
