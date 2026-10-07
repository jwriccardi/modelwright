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
    /// The canonical identity used for cycle detection, <c>WORKBOOK|SHEET|ADDRESS</c> in upper case (names use their
    /// label when they have no address); null for items that are part of a formula's structure (functions, groups)
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
                    return ((Workbook ?? string.Empty) + "|" + (Sheet ?? string.Empty) + "|" + (Address ?? Label))
                        .ToUpperInvariant();
            }
        }
    }

    /// <summary>The label.</summary>
    public override string ToString() => Label;
}
