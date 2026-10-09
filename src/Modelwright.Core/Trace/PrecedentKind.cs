namespace Modelwright.Core.Trace;

/// <summary>The kind of a <see cref="PrecedentItem"/>: what a row of the Trace In tree stands for.</summary>
public enum PrecedentKind
{
    /// <summary>One cell. Expanding it shows its formula's precedents.</summary>
    Cell,

    /// <summary>
    /// A range of more than one cell. Expanding it shows its cells: all of them up to
    /// <see cref="PrecedentTree.SummaryThreshold"/>, else a page at a time.
    /// </summary>
    Range,

    /// <summary>A defined name; expanding it shows what it refers to.</summary>
    Name,

    /// <summary>A structured (table) reference; expanding it shows the cells or range it refers to.</summary>
    Table,

    /// <summary>A reference computed at run time (INDIRECT, OFFSET, INDEX, CHOOSE); expanding it shows its target.</summary>
    DynamicReference,

    /// <summary>A function call in evaluate mode (v1.1); expanding it shows its arguments.</summary>
    Function,

    /// <summary>A parenthesized group or operator expression in evaluate mode (v1.1).</summary>
    Group,

    /// <summary>A reference that cannot be followed: <c>#REF!</c>, or a closed workbook that could not be opened.</summary>
    Error,

    /// <summary>Added by the tree: expanding it loads the next page of a large range's cells.</summary>
    MoreCells,

    /// <summary>Added by the tree: children were left out because a node or depth limit was reached.</summary>
    Truncated,
}
