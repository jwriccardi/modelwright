using System;
using ExcelDna.Integration;

namespace Modelwright.AddIn;

/// <summary>
/// Whether Excel's own Undo and Redo commands are enabled, from <c>CommandBars.GetEnabledMso</c>. Used by the
/// Ctrl+Z / Ctrl+Y hook (<see cref="UndoKeyHook"/>) and by <see cref="UndoCommand"/> (in macro context). Main
/// thread only. Never throws: a state that cannot be read is null (unknown).
/// </summary>
internal static class NativeUndoState
{
    /// <summary>
    /// Excel's Undo and Redo states, and, when <paramref name="withRepeat"/>, its Repeat state (for the diagnostics
    /// log only: it shows whether the <c>Redo</c> control reports Repeat when there is nothing to redo).
    /// </summary>
    public static (bool? Undo, bool? Redo, bool? Repeat) Query(bool withRepeat)
    {
        object commandBars;
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            commandBars = app.CommandBars;
        }
        catch (Exception)
        {
            return (null, null, null);
        }

        return (Enabled(commandBars, "Undo"), Enabled(commandBars, "Redo"), withRepeat ? Enabled(commandBars, "Repeat") : null);
    }

    /// <summary><c>true</c>, <c>false</c> or <c>unknown</c>, for the log.</summary>
    public static string Describe(bool? state) => state is bool value ? (value ? "true" : "false") : "unknown";

    private static bool? Enabled(object commandBars, string idMso)
    {
        try
        {
            dynamic bars = commandBars;
            object enabled = bars.GetEnabledMso(idMso);
            return enabled is bool value ? value : (bool?)null;
        }
        catch (Exception)
        {
            return null;
        }
    }
}
