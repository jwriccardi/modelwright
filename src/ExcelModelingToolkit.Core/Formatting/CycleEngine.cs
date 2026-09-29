using System;
using System.Collections.Generic;

namespace ExcelModelingToolkit.Core.Formatting;

/// <summary>
/// Chooses the next item of a cycle (docs/PLAN.md section 4.3). Pure logic: the caller reads the active cell,
/// applies the chosen item, and keeps the returned <see cref="CycleState"/> for the next press.
/// </summary>
/// <remarks>
/// <para>
/// The engine also remembers number format aliases: what Excel reports back after applying an item can differ
/// from the item's code (Excel may normalize it). <see cref="RecordReadBack(CycleDefinition, CycleState, CycleValue)"/>
/// records that text so the cell is still recognized as that item next time. Aliases are kept per cycle id and
/// item code, for the engine's lifetime. When the cycle definitions change, call <see cref="RetainAliases"/> to
/// drop the aliases of codes that are gone.
/// </para>
/// <para>Not thread-safe; the add-in uses it on Excel's main thread only.</para>
/// </remarks>
public sealed class CycleEngine
{
    // cycle id -> item code -> texts Excel reported back for that code (other than the code itself).
    private readonly Dictionary<string, Dictionary<string, HashSet<string>>> _aliases =
        new Dictionary<string, Dictionary<string, HashSet<string>>>(StringComparer.Ordinal);

    /// <summary>
    /// Chooses the item to apply:
    /// <list type="number">
    /// <item>If <paramref name="previous"/> is for this cycle and <paramref name="selectionKey"/>, and the cell
    /// still holds what was applied (or its value is unknown, e.g. mixed), the next position after the last one.</item>
    /// <item>Otherwise, if <paramref name="current"/> matches item <c>k</c> (for number formats: the exact code,
    /// else a recorded alias), item <c>k+1</c>.</item>
    /// <item>Otherwise, the first item.</item>
    /// </list>
    /// Positions wrap from the last item to the first.
    /// </summary>
    /// <param name="cycle">The cycle; it must have at least one item.</param>
    /// <param name="current">The active cell's current value, or <see cref="CycleValue.Unknown"/>.</param>
    /// <param name="selectionKey">Identifies the current selection (compared ordinally).</param>
    /// <param name="previous">The state returned by the previous call for this cycle, or null.</param>
    /// <exception cref="ArgumentNullException"><paramref name="cycle"/> or <paramref name="selectionKey"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// The cycle has no items, or <paramref name="current"/> is a number format for a color cycle (or vice versa).
    /// </exception>
    public CycleStep Next(CycleDefinition cycle, CycleValue current, string selectionKey, CycleState? previous)
    {
        if (cycle is null)
        {
            throw new ArgumentNullException(nameof(cycle));
        }

        if (selectionKey is null)
        {
            throw new ArgumentNullException(nameof(selectionKey));
        }

        var count = cycle.Items.Count;
        if (count == 0)
        {
            throw new ArgumentException($"Cycle '{cycle.Id}' has no items.", nameof(cycle));
        }

        if (!cycle.Accepts(current))
        {
            throw new ArgumentException($"Cycle '{cycle.Id}' ({cycle.Kind}) cannot compare the value '{current}'.", nameof(current));
        }

        int index;
        CycleStepReason reason;
        if (IsContinuation(cycle, current, selectionKey, previous))
        {
            index = (previous!.LastIndex + 1) % count;
            reason = CycleStepReason.SameSelection;
        }
        else
        {
            var matched = IndexOf(cycle, current);
            index = matched < 0 ? 0 : (matched + 1) % count;
            reason = matched < 0 ? CycleStepReason.NoMatch : CycleStepReason.Matched;
        }

        var item = cycle.Items[index];
        return new CycleStep(index, item, reason, new CycleState(cycle.Id, selectionKey, index, item.Value));
    }

    /// <summary>
    /// Index of the first item whose value equals <paramref name="value"/>, or -1. Number formats match the exact
    /// code (ordinal) first; failing that, a recorded alias. Unknown values match nothing.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="cycle"/> is null.</exception>
    public int IndexOf(CycleDefinition cycle, CycleValue value)
    {
        if (cycle is null)
        {
            throw new ArgumentNullException(nameof(cycle));
        }

        if (value.IsUnknown)
        {
            return -1;
        }

        for (var i = 0; i < cycle.Items.Count; i++)
        {
            if (cycle.Items[i].Value == value)
            {
                return i;
            }
        }

        if (value.IsNumberFormat)
        {
            for (var i = 0; i < cycle.Items.Count; i++)
            {
                if (IsAlias(cycle, cycle.Items[i], value.NumberFormat!))
                {
                    return i;
                }
            }
        }

        return -1;
    }

    /// <summary>
    /// Records what Excel reported for the active cell's number format right after applying the item in
    /// <paramref name="state"/>; see <see cref="RecordReadBack(CycleDefinition, CycleState, CycleValue)"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// The cycle is not a number format cycle, the state is for another cycle, or its index is out of range.
    /// </exception>
    public CycleState RecordReadBack(CycleDefinition cycle, CycleState state, string readBack)
    {
        if (cycle is null)
        {
            throw new ArgumentNullException(nameof(cycle));
        }

        if (readBack is null)
        {
            throw new ArgumentNullException(nameof(readBack));
        }

        if (cycle.Kind != CycleKind.NumberFormat)
        {
            throw new ArgumentException($"Cycle '{cycle.Id}' is not a number format cycle.", nameof(cycle));
        }

        return RecordReadBack(cycle, state, CycleValue.FromNumberFormat(readBack));
    }

    /// <summary>
    /// Records what Excel reported for the active cell right after applying the item in <paramref name="state"/>.
    /// Returns the state with <see cref="CycleState.LastAppliedValue"/> set to <paramref name="readBack"/>, so the
    /// next press on the same selection recognizes the cell even if Excel changed the value (a normalized number
    /// format code, or a color mapped to the nearest palette entry). For a number format that differs from the
    /// item's code, the text also becomes an alias of that item. An unknown read-back returns
    /// <paramref name="state"/> unchanged.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="cycle"/> or <paramref name="state"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// The read-back is not the cycle's kind of value, the state is for another cycle, or its index is out of range.
    /// </exception>
    public CycleState RecordReadBack(CycleDefinition cycle, CycleState state, CycleValue readBack)
    {
        if (cycle is null)
        {
            throw new ArgumentNullException(nameof(cycle));
        }

        if (state is null)
        {
            throw new ArgumentNullException(nameof(state));
        }

        if (!cycle.Accepts(readBack))
        {
            throw new ArgumentException($"Cycle '{cycle.Id}' ({cycle.Kind}) cannot record the value '{readBack}'.", nameof(readBack));
        }

        if (!string.Equals(state.CycleId, cycle.Id, StringComparison.Ordinal) || state.LastIndex >= cycle.Items.Count)
        {
            throw new ArgumentException($"The state is not for an item of cycle '{cycle.Id}'.", nameof(state));
        }

        if (readBack.IsUnknown)
        {
            return state;
        }

        if (cycle.Items[state.LastIndex] is NumberFormatItem item &&
            readBack.IsNumberFormat &&
            !string.Equals(item.Code, readBack.NumberFormat, StringComparison.Ordinal))
        {
            if (!_aliases.TryGetValue(cycle.Id, out var byCode))
            {
                byCode = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
                _aliases[cycle.Id] = byCode;
            }

            if (!byCode.TryGetValue(item.Code, out var aliases))
            {
                aliases = new HashSet<string>(StringComparer.Ordinal);
                byCode[item.Code] = aliases;
            }

            aliases.Add(readBack.NumberFormat!);
        }

        return new CycleState(state.CycleId, state.SelectionKey, state.LastIndex, readBack);
    }

    /// <summary>
    /// Keeps only the aliases that still apply to <paramref name="cycles"/> (the new definitions after a settings
    /// reload): those of a number format cycle with the same id whose items still include the aliased code.
    /// Aliases of removed cycles, of cycles that are no longer number format cycles, and of removed codes are
    /// dropped.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="cycles"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="cycles"/> contains null.</exception>
    public void RetainAliases(IEnumerable<CycleDefinition> cycles)
    {
        var codesById = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        foreach (var cycle in cycles ?? throw new ArgumentNullException(nameof(cycles)))
        {
            if (cycle is null)
            {
                throw new ArgumentException("A cycle cannot be null.", nameof(cycles));
            }

            if (cycle.Kind != CycleKind.NumberFormat)
            {
                continue;
            }

            if (!codesById.TryGetValue(cycle.Id, out var codes))
            {
                codes = new HashSet<string>(StringComparer.Ordinal);
                codesById[cycle.Id] = codes;
            }

            foreach (var item in cycle.Items)
            {
                if (item is NumberFormatItem format)
                {
                    codes.Add(format.Code);
                }
            }
        }

        foreach (var cycleId in new List<string>(_aliases.Keys))
        {
            if (!codesById.TryGetValue(cycleId, out var codes))
            {
                _aliases.Remove(cycleId);
                continue;
            }

            var byCode = _aliases[cycleId];
            foreach (var code in new List<string>(byCode.Keys))
            {
                if (!codes.Contains(code))
                {
                    byCode.Remove(code);
                }
            }

            if (byCode.Count == 0)
            {
                _aliases.Remove(cycleId);
            }
        }
    }

    private bool IsContinuation(CycleDefinition cycle, CycleValue current, string selectionKey, CycleState? previous)
    {
        if (previous is null ||
            !string.Equals(previous.CycleId, cycle.Id, StringComparison.Ordinal) ||
            !string.Equals(previous.SelectionKey, selectionKey, StringComparison.Ordinal) ||
            previous.LastIndex >= cycle.Items.Count)
        {
            return false;
        }

        // Unknown (e.g. mixed) is no evidence of a manual change, so keep going.
        if (current.IsUnknown || current == previous.LastAppliedValue)
        {
            return true;
        }

        var lastItem = cycle.Items[previous.LastIndex];
        return current == lastItem.Value || (current.IsNumberFormat && IsAlias(cycle, lastItem, current.NumberFormat!));
    }

    private bool IsAlias(CycleDefinition cycle, CycleItem item, string numberFormat) =>
        item is NumberFormatItem format &&
        _aliases.TryGetValue(cycle.Id, out var byCode) &&
        byCode.TryGetValue(format.Code, out var aliases) &&
        aliases.Contains(numberFormat);
}
