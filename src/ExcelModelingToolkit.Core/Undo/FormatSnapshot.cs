using System;
using System.Collections.Generic;
using System.Linq;
using ExcelModelingToolkit.Core.Formatting;

namespace ExcelModelingToolkit.Core.Undo;

/// <summary>
/// What one cycle changed, for undo and redo: the property (<see cref="Kind"/>), the workbook and sheet, and the
/// blocks of cells with their value before (<see cref="SnapshotBlock.Captured"/>) and after
/// (<see cref="SnapshotBlock.Applied"/>). Or, when the formats could not be captured, an unavailable snapshot
/// with the reason: a barrier on the undo stack, which records that a change was made there that cannot be undone
/// (so an undo never reaches past it to older changes without saying so). Immutable.
/// </summary>
public sealed class FormatSnapshot
{
    /// <summary>The longest address string <c>Worksheet.Range</c> accepts.</summary>
    public const int MaxAddressLength = 255;

    private static readonly char[] PathSeparators = { '\\', '/' };

    private FormatSnapshot(
        string label,
        CycleKind kind,
        string workbook,
        string sheet,
        IReadOnlyList<SnapshotBlock> blocks,
        string? unavailableReason)
    {
        Label = label ?? throw new ArgumentNullException(nameof(label));
        Kind = kind;
        Workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
        Sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
        Blocks = blocks;
        UnavailableReason = unavailableReason;
        CellCount = blocks.Sum(b => b.Range.CellCount);
    }

    /// <summary>The cycle's display name, e.g. <c>General Number</c>.</summary>
    public string Label { get; }

    /// <summary>The property the cycle changed.</summary>
    public CycleKind Kind { get; }

    /// <summary>
    /// The workbook's full name (<c>Workbook.FullName</c>: its path and file name, or just <c>Book1</c> if never
    /// saved). It identifies the workbook: one with the same file name from another folder is a different workbook.
    /// </summary>
    public string Workbook { get; }

    /// <summary>The workbook's file name (<c>Workbook.Name</c>): <see cref="Workbook"/> after its last <c>\</c> or <c>/</c>.</summary>
    public string WorkbookName => Workbook.Substring(Workbook.LastIndexOfAny(PathSeparators) + 1);

    /// <summary>The worksheet's name.</summary>
    public string Sheet { get; }

    /// <summary>The blocks (empty when unavailable).</summary>
    public IReadOnlyList<SnapshotBlock> Blocks { get; }

    /// <summary>The number of cells the blocks cover.</summary>
    public long CellCount { get; }

    /// <summary>True if the snapshot can be undone; false if the formats could not be captured.</summary>
    public bool IsAvailable => UnavailableReason is null;

    /// <summary>Why the formats could not be captured, or null.</summary>
    public string? UnavailableReason { get; }

    /// <summary>
    /// A snapshot of <paramref name="blocks"/>, which must all have known captured and applied values of
    /// <paramref name="kind"/> (for a font color, <see cref="CycleValue.Automatic"/> counts). <paramref name="workbook"/>
    /// is the workbook's full name (see <see cref="Workbook"/>).
    /// </summary>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// There are no blocks, a block is null, or a block's applied value is unknown or not of <paramref name="kind"/>.
    /// </exception>
    public static FormatSnapshot Create(string label, CycleKind kind, string workbook, string sheet, IEnumerable<SnapshotBlock> blocks)
    {
        var list = (blocks ?? throw new ArgumentNullException(nameof(blocks))).ToArray();
        if (list.Length == 0)
        {
            throw new ArgumentException("A snapshot needs at least one block.", nameof(blocks));
        }

        foreach (var block in list)
        {
            if (block is null)
            {
                throw new ArgumentException("A block cannot be null.", nameof(blocks));
            }

            if (!IsValueOf(kind, block.Captured) || !IsValueOf(kind, block.Applied))
            {
                throw new ArgumentException(
                    $"Block {block}: both values must be known {kind} values.",
                    nameof(blocks));
            }
        }

        return new FormatSnapshot(label, kind, workbook, sheet, list, null);
    }

    /// <summary>
    /// A snapshot that cannot be undone (a barrier), with the reason (e.g. the capture cap was reached).
    /// <paramref name="workbook"/> is the workbook's full name (see <see cref="Workbook"/>).
    /// </summary>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    public static FormatSnapshot Unavailable(string label, CycleKind kind, string workbook, string sheet, string reason) =>
        new FormatSnapshot(
            label,
            kind,
            workbook,
            sheet,
            Array.Empty<SnapshotBlock>(),
            reason ?? throw new ArgumentNullException(nameof(reason)));

    /// <summary>
    /// The writes for <paramref name="key"/>: blocks grouped by the value they receive
    /// (<see cref="SnapshotBlock.Target"/>). A group's union address stays within <see cref="MaxAddressLength"/>;
    /// a value whose blocks do not fit in one address gets several groups. Every block is in exactly one group.
    /// </summary>
    public IReadOnlyList<WriteGroup> WriteGroups(UndoKey key)
    {
        var groups = new List<WriteGroup>();
        var open = new Dictionary<CycleValue, (List<string> Addresses, List<SnapshotBlock> Blocks, int Length)>();
        var order = new List<CycleValue>();
        foreach (var block in Blocks)
        {
            var value = block.Target(key);
            var address = block.Address;
            if (open.TryGetValue(value, out var group) && group.Length + 1 + address.Length > MaxAddressLength)
            {
                groups.Add(Close(value, group.Addresses, group.Blocks));
                open.Remove(value);
            }

            if (!open.TryGetValue(value, out group))
            {
                group = (new List<string>(), new List<SnapshotBlock>(), -1);
                if (!order.Contains(value))
                {
                    order.Add(value);
                }
            }

            group.Addresses.Add(address);
            group.Blocks.Add(block);
            open[value] = (group.Addresses, group.Blocks, group.Length + 1 + address.Length);
        }

        foreach (var value in order)
        {
            if (open.TryGetValue(value, out var group))
            {
                groups.Add(Close(value, group.Addresses, group.Blocks));
            }
        }

        return groups;
    }

    /// <summary><c>label on workbook|sheet: n blocks, m cells</c>, or the unavailable reason.</summary>
    public override string ToString() =>
        IsAvailable
            ? $"{Label} on {Workbook}|{Sheet}: {Blocks.Count} blocks, {CellCount} cells"
            : $"{Label} on {Workbook}|{Sheet}: unavailable ({UnavailableReason})";

    private static WriteGroup Close(CycleValue value, List<string> addresses, List<SnapshotBlock> blocks) =>
        new WriteGroup(string.Join(",", addresses), value, blocks);

    private static bool IsValueOf(CycleKind kind, CycleValue value) =>
        kind == CycleKind.NumberFormat
            ? value.IsNumberFormat
            : kind == CycleKind.FontColor
                ? value.IsAutomatic || (value.IsColor && !value.Color!.Value.IsNoFill)
                : value.IsColor;
}
