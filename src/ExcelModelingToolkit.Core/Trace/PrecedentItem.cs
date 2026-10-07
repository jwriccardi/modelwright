using System;

namespace ExcelModelingToolkit.Core.Trace;

/// <summary>
/// What a row of the Trace In tree stands for, as an <see cref="IPrecedentProvider"/> describes it: a cell, range,
/// name, table reference or (in evaluate mode) a function or group, with the text of its columns.
/// </summary>
public sealed class PrecedentItem
{
    /// <summary>Creates an item.</summary>
    /// <param name="kind">What the item is.</param>
    /// <param name="label">The Precedents column: <c>Sheet2!B5</c>, <c>Revenue</c>, <c>SUM(...)</c>.</param>
    /// <param name="workbook">The workbook holding the target, or null if it has no location (a function node).</param>
    /// <param name="sheet">The sheet holding the target, or null if it has none (a workbook-level name, a function).</param>
    /// <param name="address">
    /// The target's sheet-local address (<c>B5</c>, <c>A1:C10</c>), or null if it has none. Together with
    /// <paramref name="workbook"/> and <paramref name="sheet"/> it identifies the item for cycle detection.
    /// </param>
    /// <param name="cellCount">The number of cells for a range (at least 1).</param>
    /// <param name="valueText">The Value column (the first cell's value for a range), or null.</param>
    /// <param name="canExpand">False if the item has nothing below it (a constant cell, an error).</param>
    /// <param name="argument">The Argument column in evaluate mode (<c>[value_if_true]</c>), or null.</param>
    /// <exception cref="ArgumentNullException"><paramref name="label"/> is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="cellCount"/> is below 1.</exception>
    public PrecedentItem(PrecedentKind kind, string label, string? workbook = null, string? sheet = null,
        string? address = null, long cellCount = 1, string? valueText = null, bool canExpand = true, string? argument = null)
    {
        if (cellCount < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(cellCount), cellCount, "A cell count must be at least 1.");
        }

        Kind = kind;
        Label = label ?? throw new ArgumentNullException(nameof(label));
        Workbook = workbook;
        Sheet = sheet;
        Address = address;
        CellCount = cellCount;
        ValueText = valueText;
        CanExpand = canExpand;
        Argument = argument;
    }

    /// <summary>What the item is.</summary>
    public PrecedentKind Kind { get; }

    /// <summary>The Precedents column text.</summary>
    public string Label { get; }

    /// <summary>The target's workbook, or null.</summary>
    public string? Workbook { get; }

    /// <summary>The target's sheet, or null.</summary>
    public string? Sheet { get; }

    /// <summary>The target's sheet-local address, or null.</summary>
    public string? Address { get; }

    /// <summary>The number of cells (1 unless a range).</summary>
    public long CellCount { get; }

    /// <summary>The Value column text, or null.</summary>
    public string? ValueText { get; }

    /// <summary>False if the item has nothing below it.</summary>
    public bool CanExpand { get; }

    /// <summary>The Argument column text, or null.</summary>
    public string? Argument { get; }

    /// <summary>
    /// The canonical identity used for cycle detection, <c>WORKBOOK|SHEET|ADDRESS</c> in upper case with <c>$</c>
    /// removed from the address, so <c>$A$1</c> is <c>A1</c> (names use their label when they have no address); null for items that are part of a formula's structure (functions, groups)
    /// or added by the tree, which cannot form a cycle.
    /// </summary>
    public string? Id
    {
        get
        {
            switch (Kind)
            {
                case PrecedentKind.Function:
                case PrecedentKind.Group:
                case PrecedentKind.MoreCells:
                case PrecedentKind.Truncated:
                    return null;
                default:
                    return ((Workbook ?? string.Empty) + "|" + (Sheet ?? string.Empty) + "|" +
                        (Address?.Replace("$", string.Empty) ?? Label)).ToUpperInvariant();
            }
        }
    }

    /// <summary>
    /// Creates the item for a reference in a formula. The reference leaves out its formula's own workbook and sheet;
    /// the item fills them in from <paramref name="context"/>, so the same cell always has the same <see cref="Id"/>
    /// however it was written. A name or table keeps only the sheet written before it (an unqualified name may be
    /// sheet-scoped or workbook-level, which only the workbook knows); a 3-D reference's sheet is
    /// <c>First:Last</c>.
    /// </summary>
    /// <param name="reference">A reference from <see cref="ParsedFormula.References"/>.</param>
    /// <param name="context">The workbook and sheet of the formula the reference is in.</param>
    /// <param name="label">The Precedents column, or null for the reference as written.</param>
    /// <param name="valueText">The Value column, or null.</param>
    /// <param name="canExpand">False if the item has nothing below it; always false for <c>#REF!</c>.</param>
    /// <param name="argument">The Argument column in evaluate mode, or null.</param>
    /// <exception cref="ArgumentNullException"><paramref name="reference"/> or <paramref name="context"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="reference"/> is a LET or LAMBDA local name, which is not a precedent.
    /// </exception>
    public static PrecedentItem FromReference(FormulaReference reference, FormulaContext context, string? label = null,
        string? valueText = null, bool canExpand = true, string? argument = null)
    {
        if (reference is null)
        {
            throw new ArgumentNullException(nameof(reference));
        }

        if (context is null)
        {
            throw new ArgumentNullException(nameof(context));
        }

        PrecedentKind kind;
        switch (reference.Kind)
        {
            case FormulaReferenceKind.Cell:
                kind = PrecedentKind.Cell;
                break;
            case FormulaReferenceKind.Range:
            case FormulaReferenceKind.WholeColumn:
            case FormulaReferenceKind.WholeRow:
                kind = PrecedentKind.Range;
                break;
            case FormulaReferenceKind.Name:
                kind = PrecedentKind.Name;
                break;
            case FormulaReferenceKind.StructuredReference:
                kind = PrecedentKind.Table;
                break;
            case FormulaReferenceKind.RefError:
                kind = PrecedentKind.Error;
                canExpand = false;
                break;
            default:
                throw new ArgumentException("A LET or LAMBDA local name is not a precedent.", nameof(reference));
        }

        var sheet = reference.LastSheet is null ? reference.Sheet : reference.Sheet + ":" + reference.LastSheet;
        if (sheet is null && reference.WorkbookName is null && (kind == PrecedentKind.Cell || kind == PrecedentKind.Range))
        {
            sheet = context.SheetName;
        }

        return new PrecedentItem(kind, label ?? reference.Text, reference.WorkbookName ?? context.WorkbookName, sheet,
            reference.Address, reference.Area?.CellCount ?? 1, valueText, canExpand, argument);
    }

    /// <summary>The label.</summary>
    public override string ToString() => Label;
}
