using System;

namespace Modelwright.Core.Formatting;

/// <summary>What a cycle applied last time, so the next press can continue from it. Immutable.</summary>
public sealed class CycleState
{
    /// <summary>Creates a state.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="cycleId"/> or <paramref name="selectionKey"/> is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="lastIndex"/> is negative.</exception>
    public CycleState(string cycleId, string selectionKey, int lastIndex, CycleValue lastAppliedValue)
    {
        CycleId = cycleId ?? throw new ArgumentNullException(nameof(cycleId));
        SelectionKey = selectionKey ?? throw new ArgumentNullException(nameof(selectionKey));
        if (lastIndex < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(lastIndex), lastIndex, "The index cannot be negative.");
        }

        LastIndex = lastIndex;
        LastAppliedValue = lastAppliedValue;
    }

    /// <summary>The cycle that was applied.</summary>
    public string CycleId { get; }

    /// <summary>
    /// Identifies the selection it was applied to. The add-in uses <c>workbook|sheet|address</c>; the engine
    /// only compares it for equality.
    /// </summary>
    public string SelectionKey { get; }

    /// <summary>Index of the item that was applied.</summary>
    public int LastIndex { get; }

    /// <summary>
    /// The value that was applied: the item's value, or what Excel reported back after applying it (see
    /// <see cref="CycleEngine.RecordReadBack(CycleDefinition, CycleState, CycleValue)"/>).
    /// </summary>
    public CycleValue LastAppliedValue { get; }
}
