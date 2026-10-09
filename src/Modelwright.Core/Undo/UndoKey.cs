namespace Modelwright.Core.Undo;

/// <summary>Which way to go: undo (Ctrl+Z, or the Undo formatting button) or redo (Ctrl+Y, or Redo formatting).</summary>
public enum UndoKey
{
    /// <summary>Undo: put back the formats captured before a cycle.</summary>
    Undo,

    /// <summary>Redo: apply an undone cycle's formats again.</summary>
    Redo,
}
