namespace ExcelModelingToolkit.Core.Formatting;

/// <summary>Why <see cref="CycleEngine.Next"/> chose its item.</summary>
public enum CycleStepReason
{
    /// <summary>The same cycle was just applied to the same selection and the cell still holds it: next position.</summary>
    SameSelection,

    /// <summary>The cell's value matched item <c>k</c>, so item <c>k+1</c> (wrapping) was chosen.</summary>
    Matched,

    /// <summary>The cell's value matched no item (or was unknown), so the first item was chosen.</summary>
    NoMatch,
}

/// <summary>The result of <see cref="CycleEngine.Next"/>: the item to apply and the state to keep for the next press.</summary>
public sealed class CycleStep
{
    internal CycleStep(int index, CycleItem item, CycleStepReason reason, CycleState state)
    {
        Index = index;
        Item = item;
        Reason = reason;
        State = state;
    }

    /// <summary>Zero-based index of <see cref="Item"/> in the cycle.</summary>
    public int Index { get; }

    /// <summary>The item to apply.</summary>
    public CycleItem Item { get; }

    /// <summary>Why this item was chosen.</summary>
    public CycleStepReason Reason { get; }

    /// <summary>The state to pass to the next <see cref="CycleEngine.Next"/> call for this cycle.</summary>
    public CycleState State { get; }
}
