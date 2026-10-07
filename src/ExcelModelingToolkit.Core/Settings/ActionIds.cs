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

    /// <summary>Every bindable action, in keymap order.</summary>
    public static IReadOnlyList<string> All { get; } = CycleActions.Concat(new[] { About }).ToArray();

    /// <summary>
    /// The actions added after settings files were first written (schema version 1 first shipped with the others).
    /// A settings file whose keymap does not mention one of these predates it, so loading the file gives it its
    /// default key and cycle (<see cref="ToolkitSettings.FromJson"/>); every file written since mentions them all.
    /// </summary>
    public static IReadOnlyList<string> AddedLater { get; } = new[] { BinaryCycle, RatioCycle };

    /// <summary>True if <paramref name="actionId"/> is one of <see cref="All"/> (ordinal comparison).</summary>
    public static bool IsKnown(string actionId) => All.Contains(actionId, StringComparer.Ordinal);

    /// <summary>True if <paramref name="actionId"/> is one of <see cref="CycleActions"/> (ordinal comparison).</summary>
    public static bool IsCycle(string actionId) => CycleActions.Contains(actionId, StringComparer.Ordinal);
}
