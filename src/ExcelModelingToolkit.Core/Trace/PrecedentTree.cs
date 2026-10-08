using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ExcelModelingToolkit.Core.Trace;

/// <summary>
/// The Trace In tree (docs/PLAN.md section 4.5, research/07): the audited cell at the root, children read lazily
/// through an <see cref="IPrecedentProvider"/> when a node is first expanded, a selection moved with the cursor
/// keys, and the flattened list of visible rows a list control draws.
/// </summary>
/// <remarks>
/// <para>
/// Limits, so a huge model cannot freeze Excel: a range of more than <see cref="SummaryThreshold"/> cells is one
/// node whose cells load <see cref="PageSize"/> at a time; the tree holds at most <see cref="MaxNodes"/> nodes and
/// is at most <see cref="MaxDepth"/> levels deep. Where a limit cuts children off, a
/// <see cref="PrecedentKind.Truncated"/> marker says so. An item that is its own ancestor is marked
/// <see cref="PrecedentNode.IsCycle"/> (↻) and not expanded.
/// </para>
/// <para>
/// If the provider throws, the tree is unchanged and the exception propagates: the node stays unloaded, so expanding
/// it again asks the provider again. A provider throws <see cref="PrecedentsUnavailableException"/> when Excel is
/// busy, rather than return an error row that would stay in the tree. Not thread-safe; the add-in uses it on Excel's
/// main thread.
/// </para>
/// </remarks>
public sealed class PrecedentTree
{
    /// <summary>A range with more cells than this is summarized as one node and its cells are paged.</summary>
    public const int SummaryThreshold = 50;

    /// <summary>The number of a large range's cells loaded per page.</summary>
    public const int PageSize = 100;

    /// <summary>Default <see cref="MaxNodes"/>.</summary>
    public const int DefaultMaxNodes = 5000;

    /// <summary>Default <see cref="MaxDepth"/>.</summary>
    public const int DefaultMaxDepth = 20;

    private readonly IPrecedentProvider _provider;
    private List<PrecedentNode>? _visible;

    /// <summary>Creates the tree for an audited cell and expands the root, so its precedents show at once.</summary>
    /// <param name="provider">Reads children.</param>
    /// <param name="root">The audited cell.</param>
    /// <param name="maxNodes">The most nodes the tree holds (markers not counted).</param>
    /// <param name="maxDepth">The deepest level a node can have (the root is level 0).</param>
    /// <exception cref="ArgumentNullException"><paramref name="provider"/> or <paramref name="root"/> is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">A limit is below 1.</exception>
    public PrecedentTree(IPrecedentProvider provider, PrecedentItem root, int maxNodes = DefaultMaxNodes, int maxDepth = DefaultMaxDepth)
    {
        if (maxNodes < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(maxNodes), maxNodes, "The node limit must be at least 1.");
        }

        if (maxDepth < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(maxDepth), maxDepth, "The depth limit must be at least 1.");
        }

        _provider = provider ?? throw new ArgumentNullException(nameof(provider));
        MaxNodes = maxNodes;
        MaxDepth = maxDepth;
        Root = new PrecedentNode(root ?? throw new ArgumentNullException(nameof(root)), null, false);
        NodeCount = 1;
        Selected = Root;
        Expand(Root);
    }

    /// <summary>The most nodes the tree holds.</summary>
    public int MaxNodes { get; }

    /// <summary>The deepest level a node can have.</summary>
    public int MaxDepth { get; }

    /// <summary>The audited cell.</summary>
    public PrecedentNode Root { get; }

    /// <summary>The nodes created so far, the root included and markers not.</summary>
    public int NodeCount { get; private set; }

    /// <summary>The selected node (initially the root).</summary>
    public PrecedentNode Selected { get; private set; }

    /// <summary>
    /// The rows a list control shows, top to bottom: the root, then each expanded node's children after it,
    /// depth-first.
    /// </summary>
    public IReadOnlyList<PrecedentNode> VisibleNodes
    {
        get
        {
            if (_visible is null)
            {
                _visible = new List<PrecedentNode>();
                AddVisible(Root, _visible);
            }

            return _visible;
        }
    }

    /// <summary>The position of <see cref="Selected"/> in <see cref="VisibleNodes"/>.</summary>
    public int SelectedIndex => IndexOf(Selected);

    /// <summary>Selects a visible node (a mouse click).</summary>
    /// <returns>True if the selection changed.</returns>
    /// <exception cref="ArgumentException"><paramref name="node"/> is not visible in this tree.</exception>
    public bool Select(PrecedentNode node)
    {
        if (node is null || IndexOf(node) < 0)
        {
            throw new ArgumentException("The node is not visible in this tree.", nameof(node));
        }

        var changed = !ReferenceEquals(node, Selected);
        Selected = node;
        return changed;
    }

    /// <summary>Up: selects the previous visible row.</summary>
    public TreeMove MoveUp()
    {
        var index = SelectedIndex;
        if (index <= 0)
        {
            return TreeMove.None;
        }

        Selected = VisibleNodes[index - 1];
        return TreeMove.Moved;
    }

    /// <summary>Down: selects the next visible row.</summary>
    public TreeMove MoveDown()
    {
        var index = SelectedIndex;
        if (index >= VisibleNodes.Count - 1)
        {
            return TreeMove.None;
        }

        Selected = VisibleNodes[index + 1];
        return TreeMove.Moved;
    }

    /// <summary>
    /// Right: expands the selected node, reading its children if needed; if it is already expanded, selects its
    /// first child. On a "more cells" row it loads the next page and selects that page's first cell.
    /// </summary>
    public TreeMove MoveRight()
    {
        var node = Selected;
        if (node.Item.Kind == PrecedentKind.MoreCells)
        {
            return Expand(node) && !ReferenceEquals(node, Selected) ? TreeMove.Moved : TreeMove.None;
        }

        if (!node.IsExpanded)
        {
            return Expand(node) ? TreeMove.Expanded : TreeMove.None;
        }

        if (node.Children.Count == 0)
        {
            return TreeMove.None;
        }

        Selected = node.Children[0];
        return TreeMove.Moved;
    }

    /// <summary>Left: collapses the selected node if it is expanded and has children, else selects its parent.</summary>
    public TreeMove MoveLeft()
    {
        var node = Selected;
        if (node.IsExpanded && node.Children.Count > 0)
        {
            Collapse(node);
            return TreeMove.Collapsed;
        }

        if (node.Parent is null)
        {
            return TreeMove.None;
        }

        Selected = node.Parent;
        return TreeMove.Moved;
    }

    /// <summary>
    /// Expands a node, reading its children the first time. Expanding a "more cells" row replaces it with the next
    /// page of cells (and a new "more cells" row if any remain); if it was selected, the page's first cell is.
    /// </summary>
    /// <returns>True if the tree changed; false if the node cannot be expanded or already is.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="node"/> is null.</exception>
    public bool Expand(PrecedentNode node)
    {
        if (node is null)
        {
            throw new ArgumentNullException(nameof(node));
        }

        if (node.Item.Kind == PrecedentKind.MoreCells)
        {
            return LoadNextPage(node);
        }

        if (node.IsExpanded || !node.CanExpand)
        {
            return false;
        }

        if (!node.IsLoaded)
        {
            Load(node);
        }

        node.IsExpanded = true;
        _visible = null;
        return true;
    }

    /// <summary>
    /// Collapses a node, keeping its children for the next expand. If the selection was below it, the node is
    /// selected.
    /// </summary>
    /// <returns>True if the node was expanded.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="node"/> is null.</exception>
    public bool Collapse(PrecedentNode node)
    {
        if (node is null)
        {
            throw new ArgumentNullException(nameof(node));
        }

        if (!node.IsExpanded)
        {
            return false;
        }

        for (var ancestor = Selected.Parent; ancestor is not null; ancestor = ancestor.Parent)
        {
            if (ReferenceEquals(ancestor, node))
            {
                Selected = node;
                break;
            }
        }

        node.IsExpanded = false;
        _visible = null;
        return true;
    }

    private void Load(PrecedentNode node)
    {
        var children = new List<PrecedentNode>();
        if (node.Depth >= MaxDepth)
        {
            children.Add(Marker(node, string.Format(CultureInfo.InvariantCulture,
                "Truncated: the depth limit ({0} levels) was reached", MaxDepth)));
        }
        else if (node.Item.Kind == PrecedentKind.Range && node.Item.CellCount > 1)
        {
            AddPage(node, node.Item, 0, children);
        }
        else if (NodeCount >= MaxNodes)
        {
            children.Add(NodeLimitMarker(node));
        }
        else
        {
            AddChildren(node, _provider.GetPrecedents(node.Item), children);
        }

        node.ChildList = children;
        node.IsLoaded = true;
    }

    // Replaces a MoreCells node with the next page of its range. A node already replaced (a stale row) does nothing.
    private bool LoadNextPage(PrecedentNode more)
    {
        var parent = more.Parent!;
        var siblings = parent.ChildList!;
        var index = siblings.IndexOf(more);
        if (index < 0)
        {
            return false;
        }

        var page = new List<PrecedentNode>();
        AddPage(parent, more.PagedRange!, more.NextCell, page);
        siblings.RemoveAt(index);
        siblings.InsertRange(index, page);
        if (ReferenceEquals(Selected, more))
        {
            Selected = page.Count > 0 ? page[0] : parent;
        }

        _visible = null;
        return true;
    }

    // Adds a range's cells from index start: all of a small range, else one page plus a MoreCells row if more remain.
    private void AddPage(PrecedentNode parent, PrecedentItem range, long start, List<PrecedentNode> into)
    {
        var remaining = range.CellCount - start;
        var wanted = range.CellCount <= SummaryThreshold ? (int)remaining : (int)Math.Min(PageSize, remaining);
        var room = MaxNodes - NodeCount;
        if (room <= 0)
        {
            into.Add(NodeLimitMarker(parent));
            return;
        }

        // Ask for one more than fits, to know whether the node limit cuts the page short.
        var request = Math.Min(wanted, room + 1);
        var cells = _provider.GetRangeCells(range, start, request);
        var truncated = AddChildren(parent, cells, into);
        var next = start + cells.Count;
        if (!truncated && cells.Count == request && request == wanted && next < range.CellCount)
        {
            var more = Marker(parent, string.Format(CultureInfo.InvariantCulture,
                "Next {0:N0} of {1:N0} cells ({2:N0}-{3:N0})",
                Math.Min(PageSize, range.CellCount - next), range.CellCount, next + 1, Math.Min(next + PageSize, range.CellCount)),
                PrecedentKind.MoreCells);
            more.PagedRange = range;
            more.NextCell = next;
            into.Add(more);
        }
    }

    // Adds child nodes up to the node limit, marking cycles; returns true if the limit cut them short.
    private bool AddChildren(PrecedentNode parent, IReadOnlyList<PrecedentItem> items, List<PrecedentNode> into)
    {
        foreach (var item in items)
        {
            if (NodeCount >= MaxNodes)
            {
                into.Add(NodeLimitMarker(parent));
                return true;
            }

            into.Add(new PrecedentNode(item, parent, IsAncestor(parent, item)));
            NodeCount++;
        }

        return false;
    }

    private static bool IsAncestor(PrecedentNode parent, PrecedentItem item)
    {
        var id = item.Id;
        if (id is null)
        {
            return false;
        }

        for (var node = parent; node is not null; node = node.Parent)
        {
            if (string.Equals(node.Item.Id, id, StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    private PrecedentNode NodeLimitMarker(PrecedentNode parent) =>
        Marker(parent, string.Format(CultureInfo.InvariantCulture,
            "Truncated: the node limit ({0:N0}) was reached", MaxNodes));

    private static PrecedentNode Marker(PrecedentNode parent, string label, PrecedentKind kind = PrecedentKind.Truncated) =>
        new PrecedentNode(new PrecedentItem(kind, label, canExpand: kind == PrecedentKind.MoreCells), parent, false);

    private int IndexOf(PrecedentNode node)
    {
        var visible = VisibleNodes;
        for (var index = 0; index < visible.Count; index++)
        {
            if (ReferenceEquals(visible[index], node))
            {
                return index;
            }
        }

        return -1;
    }

    // Depth-first, with a loop rather than recursion so no depth limit can overflow the stack.
    private static void AddVisible(PrecedentNode root, List<PrecedentNode> into)
    {
        var pending = new Stack<PrecedentNode>();
        pending.Push(root);
        while (pending.Count > 0)
        {
            var node = pending.Pop();
            into.Add(node);
            if (!node.IsExpanded)
            {
                continue;
            }

            for (var index = node.Children.Count - 1; index >= 0; index--)
            {
                pending.Push(node.Children[index]);
            }
        }
    }
}
