using System;
using System.Collections.Generic;
using ExcelModelingToolkit.Core.Formatting;

namespace ExcelModelingToolkit.Core.Settings;

/// <summary>
/// One cycle of a <see cref="SettingsDraft"/>, open for editing: its items can be added, removed, moved, renamed
/// and given a new number format code or color. The id, display name and kind are fixed. Items are the immutable
/// <see cref="CycleItem"/>s; an edit replaces one. Invalid contents (a blank name or code, no items) are allowed
/// while editing and reported by <see cref="SettingsDraft.ToSettings"/>.
/// </summary>
/// <remarks>
/// <para>
/// <b>Provisional cycles.</b> A provisional cycle (<see cref="CycleDefinition.Provisional"/>) ships placeholder
/// items until the owner's real Macabacus list is captured. Any change to its items (adding, removing, moving,
/// renaming, or a new code or color) clears <see cref="Provisional"/>: the list is now the user's own, so it must
/// no longer be marked (and must not be replaced) as a placeholder. An edit that changes nothing (the same name,
/// code or color, or a move at the end of the list) leaves the flag alone.
/// </para>
/// </remarks>
public sealed class CycleDraft
{
    /// <summary>The name given to an added number format item.</summary>
    public const string NewFormatName = "New format";

    /// <summary>The code given to an added number format item.</summary>
    public const string NewFormatCode = "General";

    /// <summary>The name given to an added color item.</summary>
    public const string NewColorName = "New color";

    private readonly List<CycleItem> _items;

    internal CycleDraft(CycleDefinition cycle)
    {
        Id = cycle.Id;
        DisplayName = cycle.DisplayName;
        Kind = cycle.Kind;
        Provisional = cycle.Provisional;
        _items = new List<CycleItem>(cycle.Items);
    }

    /// <summary>The cycle's id, also the action id of its command (see <see cref="CycleDefinition.Id"/>).</summary>
    public string Id { get; }

    /// <summary>The name shown to people, e.g. <c>Number</c>.</summary>
    public string DisplayName { get; }

    /// <summary>The cell property the cycle changes.</summary>
    public CycleKind Kind { get; }

    /// <summary>
    /// True while the items are still the shipped placeholders; cleared by the first edit of the items (see the
    /// remarks on <see cref="CycleDraft"/>).
    /// </summary>
    public bool Provisional { get; private set; }

    /// <summary>The items, in cycle order.</summary>
    public IReadOnlyList<CycleItem> Items => _items;

    /// <summary>True for a fill color cycle, the only kind whose items may be <see cref="OleColor.NoFill"/>.</summary>
    public bool AllowsNoFill => Kind == CycleKind.FillColor;

    /// <summary>
    /// Inserts a new item at <paramref name="index"/> (0 to <c>Items.Count</c>) and returns that index. A number
    /// format item is <see cref="NewFormatName"/> with code <see cref="NewFormatCode"/>; a color item is
    /// <see cref="NewColorName"/>, black in a font cycle and white in a fill cycle.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is outside 0 to <c>Items.Count</c>.</exception>
    public int AddItem(int index)
    {
        if (index < 0 || index > _items.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(index), index, $"The index must be between 0 and {_items.Count}.");
        }

        CycleItem item = Kind switch
        {
            CycleKind.NumberFormat => new NumberFormatItem(NewFormatName, NewFormatCode),
            CycleKind.FillColor => new ColorItem(NewColorName, OleColor.FromRgb(255, 255, 255)),
            _ => new ColorItem(NewColorName, OleColor.FromRgb(0, 0, 0)),
        };
        _items.Insert(index, item);
        Edited();
        return index;
    }

    /// <summary>Removes the item at <paramref name="index"/>.</summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is not an item index.</exception>
    public void RemoveItem(int index)
    {
        CheckIndex(index);
        _items.RemoveAt(index);
        Edited();
    }

    /// <summary>
    /// Moves the item at <paramref name="index"/> one place earlier and returns its new index (unchanged for the
    /// first item).
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is not an item index.</exception>
    public int MoveItemUp(int index)
    {
        CheckIndex(index);
        if (index == 0)
        {
            return index;
        }

        Swap(index, index - 1);
        return index - 1;
    }

    /// <summary>
    /// Moves the item at <paramref name="index"/> one place later and returns its new index (unchanged for the last
    /// item).
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is not an item index.</exception>
    public int MoveItemDown(int index)
    {
        CheckIndex(index);
        if (index == _items.Count - 1)
        {
            return index;
        }

        Swap(index, index + 1);
        return index + 1;
    }

    /// <summary>Renames the item at <paramref name="index"/>. A blank name is reported by validation.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="name"/> is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is not an item index.</exception>
    public void RenameItem(int index, string name)
    {
        if (name is null)
        {
            throw new ArgumentNullException(nameof(name));
        }

        CheckIndex(index);
        var item = _items[index];
        if (string.Equals(item.Name, name, StringComparison.Ordinal))
        {
            return;
        }

        Replace(index, item switch
        {
            NumberFormatItem format => new NumberFormatItem(name, format.Code),
            ColorItem color => new ColorItem(name, color.Color),
            _ => throw new InvalidOperationException($"Unknown item type {item.GetType().Name}."),
        });
    }

    /// <summary>Sets the number format code of the item at <paramref name="index"/>. A blank or too long code is reported by validation.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="code"/> is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is not an item index.</exception>
    /// <exception cref="InvalidOperationException">This is not a number format cycle.</exception>
    public void SetCode(int index, string code)
    {
        if (code is null)
        {
            throw new ArgumentNullException(nameof(code));
        }

        CheckIndex(index);
        if (!(_items[index] is NumberFormatItem format) || Kind != CycleKind.NumberFormat)
        {
            throw new InvalidOperationException($"Cycle '{Id}' is a color cycle; its items have colors, not number format codes.");
        }

        if (!string.Equals(format.Code, code, StringComparison.Ordinal))
        {
            Replace(index, new NumberFormatItem(format.Name, code));
        }
    }

    /// <summary>Sets the color of the item at <paramref name="index"/>.</summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is not an item index.</exception>
    /// <exception cref="InvalidOperationException">This is not a color cycle.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="color"/> is <see cref="OleColor.NoFill"/> and this is not a fill cycle (<see cref="AllowsNoFill"/>).
    /// </exception>
    public void SetColor(int index, OleColor color)
    {
        CheckIndex(index);
        if (!(_items[index] is ColorItem item) || Kind == CycleKind.NumberFormat)
        {
            throw new InvalidOperationException($"Cycle '{Id}' is a number format cycle; its items have codes, not colors.");
        }

        if (color.IsNoFill && !AllowsNoFill)
        {
            throw new ArgumentException("\"No fill\" is only valid in a fill color cycle.", nameof(color));
        }

        if (item.Color != color)
        {
            Replace(index, new ColorItem(item.Name, color));
        }
    }

    /// <summary>The cycle as it now stands (not validated).</summary>
    public CycleDefinition ToDefinition() => new CycleDefinition(Id, DisplayName, Kind, _items, Provisional);

    private void Replace(int index, CycleItem item)
    {
        _items[index] = item;
        Edited();
    }

    private void Swap(int a, int b)
    {
        (_items[a], _items[b]) = (_items[b], _items[a]);
        Edited();
    }

    private void Edited() => Provisional = false;

    private void CheckIndex(int index)
    {
        if (index < 0 || index >= _items.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(index), index, $"Cycle '{Id}' has {_items.Count} items.");
        }
    }
}
