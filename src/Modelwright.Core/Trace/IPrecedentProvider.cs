using System.Collections.Generic;

namespace Modelwright.Core.Trace;

/// <summary>
/// Reads what the Trace In tree shows under an item. The add-in implements it over Excel (formulas parsed with
/// <see cref="FormulaParser"/>, values from <c>Value2</c>); tests use a fake. <see cref="PrecedentTree"/> calls it
/// only when a node is first expanded, so nothing is read until it is needed.
/// </summary>
public interface IPrecedentProvider
{
    /// <summary>
    /// The items directly below a non-range item: a cell's precedents in the order written (duplicates included), a
    /// name's or table reference's target, a dynamic reference's resolved range, or (in evaluate mode) a function's
    /// arguments. Empty if there are none.
    /// </summary>
    IReadOnlyList<PrecedentItem> GetPrecedents(PrecedentItem item);

    /// <summary>
    /// Up to <paramref name="count"/> cells of a <see cref="PrecedentKind.Range"/> item in row-major order, starting
    /// at the 0-based cell index <paramref name="start"/>. Fewer means the range has no more cells.
    /// </summary>
    IReadOnlyList<PrecedentItem> GetRangeCells(PrecedentItem range, long start, int count);
}
