using System;
using System.Collections.Generic;
using System.Linq;

namespace Modelwright.Core.Trace;

/// <summary>
/// The rows the classic Trace In tree shows under a cell (docs/PLAN.md section 4.5): one per reference in the order
/// written, duplicates included, with each call that computes a reference (INDEX, OFFSET, INDIRECT, CHOOSE) listed
/// where it starts, just before the references written inside it. LET and LAMBDA local names are left out.
/// </summary>
public static class ClassicPrecedents
{
    /// <summary>The rows for <paramref name="formula"/>; empty if it could not be parsed.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="formula"/> is null.</exception>
    public static IReadOnlyList<ClassicPrecedent> Of(ParsedFormula formula)
    {
        if (formula is null)
        {
            throw new ArgumentNullException(nameof(formula));
        }

        var rows = new List<ClassicPrecedent>(formula.References.Count + formula.DynamicReferences.Count);
        var dynamics = formula.DynamicReferences.OrderBy(d => d.Start).ToList(); // a stable sort: outer calls first
        var dynamicIndex = 0;
        foreach (var reference in formula.References.OrderBy(r => r.Start))
        {
            while (dynamicIndex < dynamics.Count && dynamics[dynamicIndex].Start <= reference.Start)
            {
                rows.Add(new ClassicPrecedent(null, dynamics[dynamicIndex++]));
            }

            if (reference.Kind != FormulaReferenceKind.LocalName)
            {
                rows.Add(new ClassicPrecedent(reference, null));
            }
        }

        while (dynamicIndex < dynamics.Count)
        {
            rows.Add(new ClassicPrecedent(null, dynamics[dynamicIndex++]));
        }

        return rows;
    }
}

/// <summary>A row of <see cref="ClassicPrecedents.Of"/>: a reference, or a call that computes one.</summary>
public sealed class ClassicPrecedent
{
    internal ClassicPrecedent(FormulaReference? reference, DynamicReference? dynamic)
    {
        Reference = reference;
        Dynamic = dynamic;
    }

    /// <summary>The reference, or null for a <see cref="Dynamic"/> row.</summary>
    public FormulaReference? Reference { get; }

    /// <summary>The call that computes a reference, or null for a <see cref="Reference"/> row.</summary>
    public DynamicReference? Dynamic { get; }

    /// <summary>The text as written.</summary>
    public string Text => Reference?.Text ?? Dynamic!.Text;

    /// <summary>The text as written.</summary>
    public override string ToString() => Text;
}
