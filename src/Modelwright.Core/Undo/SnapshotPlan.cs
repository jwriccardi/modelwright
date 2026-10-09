using System;
using System.Collections.Generic;

namespace Modelwright.Core.Undo;

/// <summary>
/// The result of <see cref="SnapshotPlanner.Plan"/>: the captured blocks (their applied values still unknown) and
/// the number of reads it took, or the reason the formats could not be captured.
/// </summary>
public sealed class SnapshotPlan
{
    private SnapshotPlan(IReadOnlyList<SnapshotBlock> blocks, int reads, string? unavailableReason)
    {
        Blocks = blocks;
        Reads = reads;
        UnavailableReason = unavailableReason;
        foreach (var block in blocks)
        {
            CellCount += block.Range.CellCount;
        }
    }

    /// <summary>The captured blocks (empty when unavailable).</summary>
    public IReadOnlyList<SnapshotBlock> Blocks { get; }

    /// <summary>How many <see cref="IFormatReader.ReadUniform"/> calls were made.</summary>
    public int Reads { get; }

    /// <summary>The number of cells the blocks cover.</summary>
    public long CellCount { get; }

    /// <summary>True if every cell was captured.</summary>
    public bool IsAvailable => UnavailableReason is null;

    /// <summary>Why the formats could not be captured, or null.</summary>
    public string? UnavailableReason { get; }

    /// <summary>A plan whose capture failed, with the reason and the reads made (for the log).</summary>
    public static SnapshotPlan Unavailable(string reason, int reads) =>
        new SnapshotPlan(Array.Empty<SnapshotBlock>(), reads, reason ?? throw new ArgumentNullException(nameof(reason)));

    internal static SnapshotPlan Captured(IReadOnlyList<SnapshotBlock> blocks, int reads) =>
        new SnapshotPlan(blocks, reads, null);
}
