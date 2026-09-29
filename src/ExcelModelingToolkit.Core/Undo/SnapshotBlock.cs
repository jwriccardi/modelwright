using System;
using ExcelModelingToolkit.Core.Formatting;

namespace ExcelModelingToolkit.Core.Undo;

/// <summary>
/// A rectangle of cells that held one value before a cycle (<see cref="Captured"/>) and the value the cycle
/// wrote there (<see cref="Applied"/>). Immutable.
/// </summary>
public sealed class SnapshotBlock
{
    /// <summary>Creates a block.</summary>
    /// <param name="range">The cells.</param>
    /// <param name="captured">Their value before the change; must be known.</param>
    /// <param name="applied">
    /// The value written; <see cref="CycleValue.Unknown"/> until the write has happened (see <see cref="WithApplied"/>).
    /// </param>
    /// <exception cref="ArgumentException"><paramref name="captured"/> is unknown.</exception>
    public SnapshotBlock(CellRect range, CycleValue captured, CycleValue applied)
    {
        if (captured.IsUnknown)
        {
            throw new ArgumentException("A block's captured value must be known.", nameof(captured));
        }

        Range = range;
        Captured = captured;
        Applied = applied;
    }

    /// <summary>The cells.</summary>
    public CellRect Range { get; }

    /// <summary>The sheet-local A1 address of <see cref="Range"/>.</summary>
    public string Address => Range.Address;

    /// <summary>The value the cells held before the change.</summary>
    public CycleValue Captured { get; }

    /// <summary>The value written to the cells (unknown until the write has happened).</summary>
    public CycleValue Applied { get; }

    /// <summary>This block with <see cref="Applied"/> set to <paramref name="applied"/>.</summary>
    public SnapshotBlock WithApplied(CycleValue applied) => new SnapshotBlock(Range, Captured, applied);

    /// <summary>
    /// What the cells must hold for <paramref name="key"/> to be safe: <see cref="Applied"/> before an undo,
    /// <see cref="Captured"/> before a redo.
    /// </summary>
    public CycleValue Expected(UndoKey key) => key == UndoKey.Undo ? Applied : Captured;

    /// <summary>What <paramref name="key"/> writes: <see cref="Captured"/> for an undo, <see cref="Applied"/> for a redo.</summary>
    public CycleValue Target(UndoKey key) => key == UndoKey.Undo ? Captured : Applied;

    /// <summary><c>address captured -> applied</c>.</summary>
    public override string ToString() => $"{Address} {Captured} -> {Applied}";
}
