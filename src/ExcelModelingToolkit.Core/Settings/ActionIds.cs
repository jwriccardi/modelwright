using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelModelingToolkit.Core.Settings;

/// <summary>
/// The actions a keymap can bind. A cycle action's id is also the id of the cycle it runs
/// (<see cref="Formatting.CycleDefinition.Id"/>). Ids are case-sensitive.
/// </summary>
public static class ActionIds
{
    /// <summary>General Number cycle.</summary>
    public const string NumberCycle = "NumberCycle";

    /// <summary>Date cycle.</summary>
    public const string DateCycle = "DateCycle";

    /// <summary>Local Currency cycle.</summary>
    public const string CurrencyCycle = "CurrencyCycle";

    /// <summary>Percent cycle.</summary>
    public const string PercentCycle = "PercentCycle";

    /// <summary>Multiple cycle.</summary>
    public const string MultipleCycle = "MultipleCycle";

    /// <summary>Binary cycle (Yes/No, Y/N, On/Off, True/False).</summary>
    public const string BinaryCycle = "BinaryCycle";

    /// <summary>Ratio cycle (exchange ratio and fractions).</summary>
    public const string RatioCycle = "RatioCycle";

    /// <summary>Font Color cycle.</summary>
    public const string FontColorCycle = "FontColorCycle";

    /// <summary>Fill Color cycle.</summary>
    public const string FillColorCycle = "FillColorCycle";

    /// <summary>Blue-Black toggle (a two-item font color cycle).</summary>
    public const string BlueBlackToggle = "BlueBlackToggle";

    /// <summary>Trace In: the precedents window for the active cell (docs/PLAN.md section 4.5).</summary>
    public const string TraceIn = "TraceIn";

    /// <summary>Last Audited Cell: goes back to the cell Trace In was last opened on.</summary>
    public const string LastAuditedCell = "LastAuditedCell";

    /// <summary>The About box.</summary>
    public const string About = "About";

    /// <summary>The actions that run a cycle of the same id, in ribbon and keymap order.</summary>
    public static IReadOnlyList<string> CycleActions { get; } = new[]
    {
        NumberCycle,
        DateCycle,
        CurrencyCycle,
        PercentCycle,
        MultipleCycle,
        BinaryCycle,
        RatioCycle,
        FontColorCycle,
        FillColorCycle,
        BlueBlackToggle,
    };

    /// <summary>The auditing actions, in ribbon and keymap order.</summary>
    public static IReadOnlyList<string> TraceActions { get; } = new[] { TraceIn, LastAuditedCell };

    /// <summary>Every bindable action, in keymap order.</summary>
    public static IReadOnlyList<string> All { get; } = CycleActions.Concat(TraceActions).Concat(new[] { About }).ToArray();

    /// <summary>
    /// The actions added after settings files were first written (schema version 1 first shipped with the others).
    /// A settings file whose keymap does not mention one of these predates it, so loading the file gives it its
    /// default key and cycle (<see cref="ToolkitSettings.FromJson"/>); every file written since mentions them all.
    /// </summary>
    public static IReadOnlyList<string> AddedLater { get; } = new[] { BinaryCycle, RatioCycle, TraceIn, LastAuditedCell };

    /// <summary>True if <paramref name="actionId"/> is one of <see cref="All"/> (ordinal comparison).</summary>
    public static bool IsKnown(string actionId) => All.Contains(actionId, StringComparer.Ordinal);

    /// <summary>True if <paramref name="actionId"/> is one of <see cref="CycleActions"/> (ordinal comparison).</summary>
    public static bool IsCycle(string actionId) => CycleActions.Contains(actionId, StringComparer.Ordinal);

    /// <summary>
    /// The name people see for an action that is not a cycle (a cycle shows its own display name):
    /// <c>Trace In</c>, <c>Last Audited Cell</c> or <c>About</c>; any other id is returned unchanged.
    /// </summary>
    public static string DisplayName(string actionId)
    {
        switch (actionId)
        {
            case TraceIn:
                return "Trace In";
            case LastAuditedCell:
                return "Last Audited Cell";
            case About:
                return "About";
            default:
                return actionId;
        }
    }
}
