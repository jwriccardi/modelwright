using System;
using System.Collections.Generic;

namespace ExcelModelingToolkit.Core.Undo;

/// <summary>
/// Our own undo and redo stacks of <see cref="FormatSnapshot"/>s (docs/PLAN.md section 4.4), and the rule that
/// decides whether a Ctrl+Z or Ctrl+Y press is ours or Excel's (<see cref="Decide"/>).
/// </summary>
/// <remarks>
/// A cycle's COM write erases Excel's native undo history (spike K2c). So if Excel has something to undo when
/// Ctrl+Z is pressed, it must be newer than our last change, and the key goes to Excel; only when Excel has
/// nothing do we undo ours. Not thread-safe; the add-in uses it on Excel's main thread only.
/// </remarks>
public sealed class UndoManager
{
    /// <summary>Default <see cref="MaxDepth"/>.</summary>
    public const int DefaultMaxDepth = 100;

    // The top of each stack is the last element.
    private readonly List<FormatSnapshot> _undo = new List<FormatSnapshot>();
    private readonly List<FormatSnapshot> _redo = new List<FormatSnapshot>();

    /// <summary>Creates empty stacks.</summary>
    /// <param name="maxDepth">The most snapshots the undo stack keeps; pushing more drops the oldest.</param>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="maxDepth"/> is below 1.</exception>
    public UndoManager(int maxDepth = DefaultMaxDepth)
    {
        if (maxDepth < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(maxDepth), maxDepth, "The depth must be at least 1.");
        }

        MaxDepth = maxDepth;
    }

    /// <summary>The most snapshots the undo stack keeps.</summary>
    public int MaxDepth { get; }

    /// <summary>Snapshots that can be undone.</summary>
    public int UndoCount => _undo.Count;

    /// <summary>Snapshots that can be redone.</summary>
    public int RedoCount => _redo.Count;

    /// <summary>
    /// Decides a Ctrl+Z (<see cref="UndoKey.Undo"/>) or Ctrl+Y (<see cref="UndoKey.Redo"/>) press:
    /// <list type="bullet">
    /// <item>our stack for that key is empty: <see cref="UndoDecision.PassToExcel"/>;</item>
    /// <item>Excel can undo (or redo) natively (<paramref name="nativeAvailable"/> is true): pass, because its
    /// entries are newer than our last change;</item>
    /// <item>Excel's state is unknown (null: the query failed, a cell is being edited, or the keyboard focus is not
    /// on the grid): pass, which is what Excel would have done without us;</item>
    /// <item>otherwise (Excel has nothing): <see cref="UndoDecision.HandleOurs"/>.</item>
    /// </list>
    /// </summary>
    /// <param name="key">The key pressed.</param>
    /// <param name="ourStackNonEmpty">True if our undo stack (for Ctrl+Z) or redo stack (for Ctrl+Y) has a snapshot.</param>
    /// <param name="nativeAvailable">
    /// Whether Excel's own Undo (for Ctrl+Z) or Redo (for Ctrl+Y) command is enabled, or null if unknown.
    /// </param>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="key"/> is not a defined value.</exception>
    public static UndoDecision Decide(UndoKey key, bool ourStackNonEmpty, bool? nativeAvailable)
    {
        if (key != UndoKey.Undo && key != UndoKey.Redo)
        {
            throw new ArgumentOutOfRangeException(nameof(key), key, "Unknown key.");
        }

        return ourStackNonEmpty && nativeAvailable == false ? UndoDecision.HandleOurs : UndoDecision.PassToExcel;
    }

    /// <summary>The number of snapshots on the stack <paramref name="key"/> takes from.</summary>
    public int Count(UndoKey key) => Stack(key).Count;

    /// <summary>
    /// Records a new change: pushes <paramref name="snapshot"/> onto the undo stack and clears the redo stack. If
    /// the undo stack is then deeper than <see cref="MaxDepth"/>, the oldest snapshot is dropped.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="snapshot"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="snapshot"/> is unavailable.</exception>
    public void Push(FormatSnapshot snapshot)
    {
        if (snapshot is null)
        {
            throw new ArgumentNullException(nameof(snapshot));
        }

        if (!snapshot.IsAvailable)
        {
            throw new ArgumentException("An unavailable snapshot cannot be undone: " + snapshot.UnavailableReason, nameof(snapshot));
        }

        _redo.Clear();
        _undo.Add(snapshot);
        if (_undo.Count > MaxDepth)
        {
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
    public FormatSnapshot? Undo() => Move(_undo, _redo);

    /// <summary>
    /// Call after the top redo snapshot has been applied again: moves it back to the undo stack and returns it
    /// (null if the redo stack is empty).
    /// </summary>
    public FormatSnapshot? Redo() => Move(_redo, _undo);

    /// <summary>
    /// Drops the snapshot <paramref name="key"/> would take next (it can no longer be applied safely) and returns
    /// it, or null if there is none.
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

    /// <summary>
    /// Drops every snapshot, on both stacks, for workbook <paramref name="workbook"/> (compared ignoring case, like
    /// Excel). Returns how many were dropped.
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
    /// Drops every snapshot, on both stacks, for sheet <paramref name="sheet"/> of workbook
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
