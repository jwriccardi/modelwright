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

    /// <summary>Font Color cycle.</summary>
    public const string FontColorCycle = "FontColorCycle";

    /// <summary>Fill Color cycle.</summary>
    public const string FillColorCycle = "FillColorCycle";

    /// <summary>Blue-Black toggle (a two-item font color cycle).</summary>
    public const string BlueBlackToggle = "BlueBlackToggle";

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
        FontColorCycle,
        FillColorCycle,
        BlueBlackToggle,
    };

    /// <summary>Every bindable action, in keymap order.</summary>
    public static IReadOnlyList<string> All { get; } = CycleActions.Concat(new[] { About }).ToArray();

    /// <summary>True if <paramref name="actionId"/> is one of <see cref="All"/> (ordinal comparison).</summary>
    public static bool IsKnown(string actionId) => All.Contains(actionId, StringComparer.Ordinal);

    /// <summary>True if <paramref name="actionId"/> is one of <see cref="CycleActions"/> (ordinal comparison).</summary>
    public static bool IsCycle(string actionId) => CycleActions.Contains(actionId, StringComparer.Ordinal);
}
