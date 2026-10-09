namespace Modelwright.Core.Undo;

/// <summary>What to do with a Ctrl+Z or Ctrl+Y press; see <see cref="UndoManager.Decide"/>.</summary>
public enum UndoDecision
{
    /// <summary>Let the key through: Excel undoes or redoes natively.</summary>
    PassToExcel,

    /// <summary>Swallow the key and undo or redo our own last formatting change.</summary>
    HandleOurs,

    /// <summary>
    /// Let the key through, and first empty our redo stack (<see cref="UndoManager.ClearRedo"/>): Excel has history
    /// newer than our last restore, so our redo entries are stale.
    /// </summary>
    PassToExcelAndClearRedo,
}
