using System.Collections.Generic;

namespace Modelwright.Core.Trace;

/// <summary>A row of the Trace In tree (<see cref="PrecedentTree"/>): an item, its place and its state.</summary>
public sealed class PrecedentNode
{
    private static readonly PrecedentNode[] NoChildren = new PrecedentNode[0];

    internal PrecedentNode(PrecedentItem item, PrecedentNode? parent, bool isCycle)
    {
        Item = item;
        Parent = parent;
        Depth = parent is null ? 0 : parent.Depth + 1;
        IsCycle = isCycle;
    }

    /// <summary>What the row stands for.</summary>
    public PrecedentItem Item { get; }

    /// <summary>The parent row, or null for the root (the audited cell).</summary>
    public PrecedentNode? Parent { get; }

    /// <summary>The level below the root (the root is 0).</summary>
    public int Depth { get; }

    /// <summary>
    /// True if the item is one of its own ancestors (a circular reference): it is shown with ↻ and cannot be
    /// expanded.
    /// </summary>
    public bool IsCycle { get; }

    /// <summary>True if the node's children have been read.</summary>
    public bool IsLoaded { get; internal set; }

    /// <summary>True if the node is expanded (its children are visible).</summary>
    public bool IsExpanded { get; internal set; }

    /// <summary>The children read so far (empty until loaded).</summary>
    public IReadOnlyList<PrecedentNode> Children => ChildList ?? (IReadOnlyList<PrecedentNode>)NoChildren;

    /// <summary>
    /// True if expanding the node can show something: not a cycle, not a truncation marker, and the item allows it.
    /// A loaded node with no children cannot be expanded further.
    /// </summary>
    public bool CanExpand =>
        !IsCycle && Item.CanExpand && Item.Kind != PrecedentKind.Truncated && (!IsLoaded || Children.Count > 0);

    /// <summary>
    /// True for a range of more than <see cref="PrecedentTree.SummaryThreshold"/> cells: the row summarizes it (the
    /// first cell's value and the cell count), and expanding it loads its cells a page at a time.
    /// </summary>
    public bool IsSummarizedRange => Item.Kind == PrecedentKind.Range && Item.CellCount > PrecedentTree.SummaryThreshold;

    internal List<PrecedentNode>? ChildList { get; set; }

    // For a MoreCells node: the range whose cells it loads, and the index of the next cell.
    internal PrecedentItem? PagedRange { get; set; }

    internal long NextCell { get; set; }

    /// <summary>The item's label.</summary>
    public override string ToString() => Item.Label;
}
