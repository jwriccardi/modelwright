using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelModelingToolkit.Core.Formatting;

/// <summary>
/// A formatting cycle: an ordered list of items that one command steps through (docs/PLAN.md section 4.3).
/// The constructor accepts any non-null contents; <see cref="Validate"/> reports what is wrong with them.
/// </summary>
public sealed class CycleDefinition
{
    /// <summary>Creates a cycle.</summary>
    /// <param name="id">Stable identifier; also the action id the cycle's command and shortcut use, e.g. <c>NumberCycle</c>.</param>
    /// <param name="displayName">Name shown to people, e.g. <c>Number</c>.</param>
    /// <param name="kind">The cell property the cycle changes.</param>
    /// <param name="items">The items, in cycle order.</param>
    /// <param name="provisional">
    /// True when the items are placeholders awaiting the owner's real list (they are shipped so the key works).
    /// </param>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="items"/> contains null.</exception>
    public CycleDefinition(string id, string displayName, CycleKind kind, IEnumerable<CycleItem> items, bool provisional = false)
    {
        Id = id ?? throw new ArgumentNullException(nameof(id));
        DisplayName = displayName ?? throw new ArgumentNullException(nameof(displayName));
        Kind = kind;
        Provisional = provisional;
        var list = (items ?? throw new ArgumentNullException(nameof(items))).ToArray();
        if (list.Any(item => item is null))
        {
            throw new ArgumentException("A cycle item cannot be null.", nameof(items));
        }

        Items = list;
    }

    /// <summary>Stable identifier, also the action id of the cycle's command, e.g. <c>NumberCycle</c>.</summary>
    public string Id { get; }

    /// <summary>Name shown to people, e.g. <c>Number</c>.</summary>
    public string DisplayName { get; }

    /// <summary>The cell property the cycle changes.</summary>
    public CycleKind Kind { get; }

    /// <summary>The items, in cycle order.</summary>
    public IReadOnlyList<CycleItem> Items { get; }

    /// <summary>
    /// True when the items are provisional placeholders, to be replaced by the owner's real Macabacus list.
    /// </summary>
    public bool Provisional { get; }

    /// <summary>True if <paramref name="value"/> is the kind of value this cycle reads and applies (unknown is always accepted).</summary>
    public bool Accepts(CycleValue value) =>
        value.IsUnknown || (Kind == CycleKind.NumberFormat ? value.IsNumberFormat : value.IsColor);

    /// <summary>
    /// Checks the cycle: a non-blank id and display name, at least one item, every item of the cycle's kind with a
    /// non-blank name, no blank number format code, and "no fill" only in a fill cycle. Returns the problems found
    /// (empty if none), each prefixed with the cycle id.
    /// </summary>
    public IReadOnlyList<string> Validate()
    {
        var problems = new List<string>();
        var label = string.IsNullOrWhiteSpace(Id) ? "cycle (no id)" : $"cycle '{Id}'";
        if (string.IsNullOrWhiteSpace(Id))
        {
            problems.Add($"{label}: the id is blank.");
        }

        if (string.IsNullOrWhiteSpace(DisplayName))
        {
            problems.Add($"{label}: the display name is blank.");
        }

        if (Items.Count == 0)
        {
            problems.Add($"{label}: it has no items; a cycle needs at least one.");
        }

        for (var i = 0; i < Items.Count; i++)
        {
            var item = Items[i];
            var itemLabel = $"{label}, item {i + 1}";
            if (string.IsNullOrWhiteSpace(item.Name))
            {
                problems.Add($"{itemLabel}: the name is blank.");
            }

            switch (item)
            {
                case NumberFormatItem format when Kind == CycleKind.NumberFormat:
                    if (string.IsNullOrWhiteSpace(format.Code))
                    {
                        problems.Add($"{itemLabel}: the number format code is blank.");
                    }

                    break;
                case ColorItem color when Kind == CycleKind.FontColor:
                    if (color.Color.IsNoFill)
                    {
                        problems.Add($"{itemLabel}: \"none\" (no fill) is only valid in a fill color cycle.");
                    }

                    break;
                case ColorItem when Kind == CycleKind.FillColor:
                    break;
                default:
                    problems.Add($"{itemLabel}: a {DescribeItem(item)} does not belong in a {DescribeKind(Kind)} cycle.");
                    break;
            }
        }

        return problems;
    }

    /// <inheritdoc />
    public override string ToString() => $"{Id} ({Kind}, {Items.Count} items)";

    private static string DescribeItem(CycleItem item) => item is NumberFormatItem ? "number format item" : "color item";

    private static string DescribeKind(CycleKind kind) => kind switch
    {
        CycleKind.NumberFormat => "number format",
        CycleKind.FontColor => "font color",
        CycleKind.FillColor => "fill color",
        _ => kind.ToString(),
    };
}
