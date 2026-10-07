using System;
using System.Collections.Generic;

namespace ExcelModelingToolkit.Core.Undo;

/// <summary>
/// Our own undo and redo stacks of <see cref="FormatSnapshot"/>s (docs/PLAN.md section 4.4), and the rule that
/// decides whether a Ctrl+Z or Ctrl+Y press is ours or Excel's (<see cref="Decide"/>).
/// </summary>
/// <remarks>
/// <para>
/// A cycle's COM write erases Excel's native undo and redo history (spike K2c), and so does our restore. So if
/// Excel has something to undo when Ctrl+Z is pressed, it must be newer than our last change, and the key goes to
/// Excel; only when Excel has nothing do we undo ours. Likewise, any native history seen while our redo stack is
/// non-empty is newer than our last restore, so our redo entries are stale and are dropped.
/// </para>
/// <para>
/// Every completed write is pushed, recorded or not: a change whose formats could not be captured goes on the
/// undo stack as a barrier (an unavailable <see cref="FormatSnapshot"/>), so an undo stops there instead of
/// reaching past it to older changes. Barriers never reach the redo stack.
/// </para>
/// <para>Not thread-safe; the add-in uses it on Excel's main thread only.</para>
/// </remarks>
public sealed class UndoManager
{
    /// <summary>Default <see cref="MaxDepth"/>.</summary>
    public const int DefaultMaxDepth = 100;

    /// <summary>Default <see cref="MaxBlocks"/>.</summary>
    public const int DefaultMaxBlocks = 200000;

    // The top of each stack is the last element.
    private readonly List<FormatSnapshot> _undo = new List<FormatSnapshot>();
    private readonly List<FormatSnapshot> _redo = new List<FormatSnapshot>();

    /// <summary>Creates empty stacks.</summary>
    /// <param name="maxDepth">The most snapshots the undo stack keeps; pushing more drops the oldest.</param>
    /// <param name="maxBlocks">
    /// The most blocks both stacks keep in all (a memory budget); pushing past it drops the oldest snapshots.
    /// </param>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="maxDepth"/> or <paramref name="maxBlocks"/> is below 1.</exception>
    public UndoManager(int maxDepth = DefaultMaxDepth, int maxBlocks = DefaultMaxBlocks)
    {
        if (maxDepth < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(maxDepth), maxDepth, "The depth must be at least 1.");
        }

        if (maxBlocks < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(maxBlocks), maxBlocks, "The block budget must be at least 1.");
        }

        MaxDepth = maxDepth;
        MaxBlocks = maxBlocks;
    }

    /// <summary>The most snapshots the undo stack keeps.</summary>
    public int MaxDepth { get; }

    /// <summary>The most blocks both stacks keep in all; the newest snapshot is always kept, however large.</summary>
    public int MaxBlocks { get; }

    /// <summary>Snapshots that can be undone, barriers included.</summary>
    public int UndoCount => _undo.Count;

    /// <summary>Snapshots that can be redone.</summary>
    public int RedoCount => _redo.Count;

    /// <summary>The blocks held on both stacks.</summary>
    public int BlockCount => Blocks(_undo) + Blocks(_redo);

    /// <summary>
    /// Decides a Ctrl+Z (<see cref="UndoKey.Undo"/>) or Ctrl+Y (<see cref="UndoKey.Redo"/>) press. A null native
    /// state is unknown (the query failed, a cell is being edited, or the keyboard focus is not on the grid).
    /// <para>Ctrl+Z:</para>
    /// <list type="bullet">
    /// <item>Excel can undo (<paramref name="nativeUndo"/> is true): its entries are newer than our last change, so
    /// pass; and if our redo stack is non-empty it is stale, so <see cref="UndoDecision.PassToExcelAndClearRedo"/>;</item>
    /// <item>Excel cannot undo and our undo stack has an entry (a barrier counts, so we never reach past it):
    /// <see cref="UndoDecision.HandleOurs"/>;</item>
    /// <item>otherwise (unknown, or ours is empty): pass, which is what Excel would have done without us.</item>
    /// </list>
    /// <para>Ctrl+Y:</para>
    /// <list type="bullet">
    /// <item>our redo stack is empty: pass;</item>
    /// <item>Excel can undo or redo: the user acted after our last restore, so our redo is stale:
    /// <see cref="UndoDecision.PassToExcelAndClearRedo"/>;</item>
    /// <item>Excel can do neither: <see cref="UndoDecision.HandleOurs"/>;</item>
    /// <item>otherwise (unknown): pass, keeping our redo.</item>
    /// </list>
    /// </summary>
    /// <param name="key">The key pressed.</param>
    /// <param name="ourUndoCount">Entries on our undo stack (<see cref="UndoCount"/>), barriers included.</param>
    /// <param name="ourRedoCount">Entries on our redo stack (<see cref="RedoCount"/>).</param>
    /// <param name="nativeUndo">Whether Excel's own Undo command is enabled, or null if unknown.</param>
    /// <param name="nativeRedo">Whether Excel's own Redo command is enabled, or null if unknown.</param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="key"/> is not a defined value, or a count is negative.
    /// </exception>
    public static UndoDecision Decide(UndoKey key, int ourUndoCount, int ourRedoCount, bool? nativeUndo, bool? nativeRedo)
    {
        if (ourUndoCount < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(ourUndoCount), ourUndoCount, "A count cannot be negative.");
        }

        if (ourRedoCount < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(ourRedoCount), ourRedoCount, "A count cannot be negative.");
        }

        switch (key)
        {
            case UndoKey.Undo:
                if (nativeUndo == true)
                {
                    return ourRedoCount > 0 ? UndoDecision.PassToExcelAndClearRedo : UndoDecision.PassToExcel;
                }

                return nativeUndo == false && ourUndoCount > 0 ? UndoDecision.HandleOurs : UndoDecision.PassToExcel;
            case UndoKey.Redo:
                if (ourRedoCount == 0)
                {
                    return UndoDecision.PassToExcel;
                }

                if (nativeUndo == true || nativeRedo == true)
                {
                    return UndoDecision.PassToExcelAndClearRedo;
                }

                return nativeUndo == false && nativeRedo == false ? UndoDecision.HandleOurs : UndoDecision.PassToExcel;
            default:
                throw new ArgumentOutOfRangeException(nameof(key), key, "Unknown key.");
        }
    }

    /// <summary>The number of snapshots on the stack <paramref name="key"/> takes from.</summary>
    public int Count(UndoKey key) => Stack(key).Count;

    /// <summary>
    /// Records a completed write: pushes <paramref name="snapshot"/> onto the undo stack (an unavailable one as a
    /// barrier) and clears the redo stack. If the undo stack is then deeper than <see cref="MaxDepth"/>, or both
    /// stacks hold more than <see cref="MaxBlocks"/> blocks, the oldest snapshots are dropped (never the new one).
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="snapshot"/> is null.</exception>
    public void Push(FormatSnapshot snapshot)
    {
        if (snapshot is null)
        {
            throw new ArgumentNullException(nameof(snapshot));
        }

        _redo.Clear();
        _undo.Add(snapshot);
        if (_undo.Count > MaxDepth)
        {
            _undo.RemoveAt(0);
        }

        // The redo stack is empty here, so the budget is the undo stack's.
        var blocks = Blocks(_undo);
        while (blocks > MaxBlocks && _undo.Count > 1)
        {
            blocks -= _undo[0].Blocks.Count;
            _undo.RemoveAt(0);
        }
    }

    /// <summary>The snapshot <paramref name="key"/> would take next (without taking it), or null if there is none.</summary>
    public FormatSnapshot? Peek(UndoKey key)
    {
        var stack = Stack(key);
        return stack.Count == 0 ? null : stack[stack.Count - 1];
    }

    /// <summary>
    /// Call after the top undo snapshot has been restored: moves it to the redo stack and returns it (null if the
    /// undo stack is empty).
    /// </summary>
    /// <exception cref="InvalidOperationException">The top is a barrier, which cannot be restored; <see cref="Discard"/> it.</exception>
    public FormatSnapshot? Undo()
    {
        if (Peek(UndoKey.Undo) is { IsAvailable: false } barrier)
        {
            throw new InvalidOperationException("The top of the undo stack is a barrier and cannot be undone: " + barrier);
        }

        return Move(_undo, _redo);
    }

    /// <summary>
    /// Call after the top redo snapshot has been applied again: moves it back to the undo stack and returns it
    /// (null if the redo stack is empty).
    /// </summary>
    public FormatSnapshot? Redo() => Move(_redo, _undo);

    /// <summary>
    /// Drops the snapshot <paramref name="key"/> would take next (it is a barrier, or can no longer be applied
    /// safely) and returns it, or null if there is none.
    /// </summary>
    public FormatSnapshot? Discard(UndoKey key)
    {
        var stack = Stack(key);
        if (stack.Count == 0)
        {
            return null;
        }

        var top = stack[stack.Count - 1];
        stack.RemoveAt(stack.Count - 1);
        return top;
    }

    /// <summary>Empties the redo stack (its entries are stale). Returns how many were dropped.</summary>
    public int ClearRedo()
    {
        var count = _redo.Count;
        _redo.Clear();
        return count;
    }

    /// <summary>
    /// Drops every snapshot, on both stacks, for the workbook whose full name is <paramref name="workbook"/>
    /// (<see cref="FormatSnapshot.Workbook"/>, compared ignoring case, like Excel). Returns how many were dropped.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="workbook"/> is null.</exception>
    public int InvalidateWorkbook(string workbook)
    {
        if (workbook is null)
        {
            throw new ArgumentNullException(nameof(workbook));
        }

        return RemoveAll(s => SameName(s.Workbook, workbook));
    }

    /// <summary>
    /// Drops every snapshot, on both stacks, for sheet <paramref name="sheet"/> of the workbook whose full name is
    /// <paramref name="workbook"/> (both compared ignoring case). Returns how many were dropped.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="workbook"/> or <paramref name="sheet"/> is null.</exception>
    public int InvalidateSheet(string workbook, string sheet)
    {
        if (workbook is null)
        {
            throw new ArgumentNullException(nameof(workbook));
        }

        if (sheet is null)
        {
            throw new ArgumentNullException(nameof(sheet));
        }

        return RemoveAll(s => SameName(s.Workbook, workbook) && SameName(s.Sheet, sheet));
    }

    /// <summary>Empties both stacks.</summary>
    public void Clear()
    {
        _undo.Clear();
        _redo.Clear();
    }

    private static int Blocks(List<FormatSnapshot> stack)
    {
        var total = 0;
        foreach (var snapshot in stack)
        {
            total += snapshot.Blocks.Count;
        }

        return total;
    }

    private static FormatSnapshot? Move(List<FormatSnapshot> from, List<FormatSnapshot> to)
    {
        if (from.Count == 0)
        {
            return null;
        }

        var top = from[from.Count - 1];
        from.RemoveAt(from.Count - 1);
        to.Add(top);
        return top;
    }

    private static bool SameName(string a, string b) => string.Equals(a, b, StringComparison.OrdinalIgnoreCase);

    private List<FormatSnapshot> Stack(UndoKey key) => key switch
    {
        UndoKey.Undo => _undo,
        UndoKey.Redo => _redo,
        _ => throw new ArgumentOutOfRangeException(nameof(key), key, "Unknown key."),
    };

    private int RemoveAll(Predicate<FormatSnapshot> match) => _undo.RemoveAll(match) + _redo.RemoveAll(match);
}
