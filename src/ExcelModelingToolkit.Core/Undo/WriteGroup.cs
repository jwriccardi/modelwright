using System.Collections.Generic;
using ExcelModelingToolkit.Core.Formatting;

namespace ExcelModelingToolkit.Core.Undo;

/// <summary>
/// Blocks of a <see cref="FormatSnapshot"/> that receive the same value, written with one property write to the
/// union address <see cref="Address"/> (comma-separated, at most <see cref="FormatSnapshot.MaxAddressLength"/>
/// characters, which is what <c>Worksheet.Range</c> accepts).
/// </summary>
public sealed class WriteGroup
{
    internal WriteGroup(string address, CycleValue value, IReadOnlyList<SnapshotBlock> blocks)
    {
        Address = address;
        Value = value;
        Blocks = blocks;
    }

    /// <summary>The blocks' addresses, comma-separated.</summary>
    public string Address { get; }

    /// <summary>The value to write.</summary>
    public CycleValue Value { get; }

    /// <summary>The blocks covered.</summary>
    public IReadOnlyList<SnapshotBlock> Blocks { get; }
}
