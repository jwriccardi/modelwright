using System;
using System.Collections.Generic;

namespace Modelwright.Core.Trace;

/// <summary>
/// The cells Trace In was opened on, newest last, for Last Audited Cell (Ctrl+Shift+\; docs/PLAN.md section 4.5):
/// each press returns to the previous audit. Holds at most <see cref="Capacity"/> audits and drops the oldest.
/// </summary>
/// <remarks>Not thread-safe; the add-in uses it on Excel's main thread.</remarks>
public sealed class AuditHistory
{
    /// <summary>Default <see cref="Capacity"/>.</summary>
    public const int DefaultCapacity = 20;

    // The newest audit is the last element.
    private readonly List<PrecedentItem> _audits = new List<PrecedentItem>();

    /// <summary>Creates an empty history.</summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="capacity"/> is below 1.</exception>
    public AuditHistory(int capacity = DefaultCapacity)
    {
        if (capacity < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(capacity), capacity, "The capacity must be at least 1.");
        }

        Capacity = capacity;
    }

    /// <summary>The most audits kept.</summary>
    public int Capacity { get; }

    /// <summary>The audits held.</summary>
    public int Count => _audits.Count;

    /// <summary>
    /// Records an audited cell. Auditing the cell that is already newest records nothing, so reopening Trace In on
    /// the same cell does not fill the history. When full, the oldest audit is dropped.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="audited"/> is null.</exception>
    public void Push(PrecedentItem audited)
    {
        if (audited is null)
        {
            throw new ArgumentNullException(nameof(audited));
        }

        if (_audits.Count > 0 && audited.Id is not null && audited.Id == _audits[_audits.Count - 1].Id)
        {
            return;
        }

        if (_audits.Count == Capacity)
        {
            _audits.RemoveAt(0);
        }

        _audits.Add(audited);
    }

    /// <summary>The newest audit, or null if the history is empty.</summary>
    public PrecedentItem? Peek() => _audits.Count == 0 ? null : _audits[_audits.Count - 1];

    /// <summary>Removes and returns the newest audit (where Last Audited Cell goes), or null if the history is empty.</summary>
    public PrecedentItem? Pop()
    {
        var newest = Peek();
        if (newest is not null)
        {
            _audits.RemoveAt(_audits.Count - 1);
        }

        return newest;
    }

    /// <summary>Forgets every audit.</summary>
    public void Clear() => _audits.Clear();
}
