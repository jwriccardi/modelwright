using System;

namespace Modelwright.Core.Settings;

/// <summary>
/// Running alongside Macabacus, which binds the same keyboard shortcuts (the defaults are its keymap): how the add-in
/// recognizes it among Excel's add-ins, and what it does when it finds it at startup (<see cref="Decide"/>).
/// </summary>
public static class Coexistence
{
    /// <summary>
    /// The environment variable that, set to <c>1</c>, stops the add-in showing notices (for automated runs; see
    /// <see cref="NoticesSuppressed"/>).
    /// </summary>
    public const string NoNoticesVariable = "MODELWRIGHT_NO_NOTICES";

    private const string MacabacusPrefix = "Macabacus";

    /// <summary>True if <paramref name="value"/>, the value of <see cref="NoNoticesVariable"/> (null when unset), is <c>1</c>.</summary>
    public static bool NoticesSuppressed(string? value) => string.Equals(value?.Trim(), "1", StringComparison.Ordinal);

    /// <summary>
    /// True for a COM add-in (Excel's <c>COMAddIns</c>) that is Macabacus and is loaded: its ProgId is "Macabacus"
    /// (any case; "Macabacus.Something" too) and it is connected. Macabacus binds its keys from this COM add-in; its
    /// workbook add-in (Macabacus.xlam) binds none, so it does not count (measured 2026-10-09).
    /// </summary>
    public static bool IsMacabacusComAddIn(string? progId, bool connected)
    {
        if (!connected || progId is null)
        {
            return false;
        }

        var id = progId.Trim();
        return string.Equals(id, MacabacusPrefix, StringComparison.OrdinalIgnoreCase) ||
            id.StartsWith(MacabacusPrefix + ".", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// What to do at startup, once the add-in has looked for Macabacus:
    /// <list type="table">
    /// <listheader><term>Macabacus found, shortcuts</term><description>Actions</description></listheader>
    /// <item><term>not found</term><description>none</description></item>
    /// <item><term>found, shortcuts off</term><description>none: Macabacus's keys stand, and the notice is not needed</description></item>
    /// <item><term>found, on, notice shown before</term><description><see cref="CoexistenceActions.Reregister"/></description></item>
    /// <item><term>found, on, notices suppressed</term><description><see cref="CoexistenceActions.Reregister"/></description></item>
    /// <item><term>found, on, notice not shown yet</term><description>Reregister and <see cref="CoexistenceActions.ShowNotice"/></description></item>
    /// </list>
    /// </summary>
    /// <param name="macabacusDetected">True if Macabacus is loaded.</param>
    /// <param name="useKeyboardShortcuts">The <see cref="ToolkitSettings.UseKeyboardShortcuts"/> setting.</param>
    /// <param name="noticeShown">True if the notice has been shown before (remembered in ui-state.json).</param>
    /// <param name="noticesSuppressed">True if notices are off for this run (<see cref="NoticesSuppressed"/>).</param>
    public static CoexistenceActions Decide(bool macabacusDetected, bool useKeyboardShortcuts, bool noticeShown, bool noticesSuppressed)
    {
        if (!macabacusDetected || !useKeyboardShortcuts)
        {
            return CoexistenceActions.None;
        }

        return noticeShown || noticesSuppressed
            ? CoexistenceActions.Reregister
            : CoexistenceActions.Reregister | CoexistenceActions.ShowNotice;
    }
}
