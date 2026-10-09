using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using Modelwright.Core.Keys;
using Modelwright.Core.Settings;

namespace Modelwright.AddIn;

/// <summary>
/// Binds the settings keymap to the add-in's commands through <c>xlcOnKey</c> (spike K3), and restores Excel's
/// defaults for the keys we bound. Call in macro context (AutoOpen, AutoClose, commands) on the main thread.
/// </summary>
internal static class KeyBindings
{
    /// <summary>Action id to the <see cref="ExcelCommandAttribute"/> macro name that runs it.</summary>
    private static readonly Dictionary<string, string> Macros = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        [ActionIds.NumberCycle] = nameof(Commands.MwNumberCycle),
        [ActionIds.DateCycle] = nameof(Commands.MwDateCycle),
        [ActionIds.CurrencyCycle] = nameof(Commands.MwCurrencyCycle),
        [ActionIds.PercentCycle] = nameof(Commands.MwPercentCycle),
        [ActionIds.MultipleCycle] = nameof(Commands.MwMultipleCycle),
        [ActionIds.BinaryCycle] = nameof(Commands.MwBinaryCycle),
        [ActionIds.RatioCycle] = nameof(Commands.MwRatioCycle),
        [ActionIds.FontColorCycle] = nameof(Commands.MwFontColorCycle),
        [ActionIds.FillColorCycle] = nameof(Commands.MwFillColorCycle),
        [ActionIds.BlueBlackToggle] = nameof(Commands.MwBlueBlackToggle),
        [ActionIds.TraceIn] = nameof(Commands.MwTraceIn),
        [ActionIds.LastAuditedCell] = nameof(Commands.MwLastAuditedCell),
        [ActionIds.About] = nameof(Commands.MwAbout),
    };

    // What is currently registered: OnKey strings, and action id -> chord.
    private static readonly HashSet<string> Registered = new HashSet<string>(StringComparer.Ordinal);
    private static readonly Dictionary<string, KeyChord> ByAction = new Dictionary<string, KeyChord>(StringComparer.Ordinal);

    /// <summary>The number of keys currently bound.</summary>
    public static int Count => Registered.Count;

    /// <summary>
    /// Binds every non-empty keymap entry to its command (re-binding keys we already hold, which takes them back
    /// from any add-in that bound them since, like Macabacus's Override). Keys we held that the keymap no longer
    /// uses are restored to Excel's default. Returns one message per key that could not be bound.
    /// </summary>
    public static IReadOnlyList<string> Apply(IReadOnlyDictionary<string, string> keymap)
    {
        var failures = new List<string>();
        var wanted = new List<KeyValuePair<string, KeyChord>>();
        foreach (var actionId in ActionIds.All)
        {
            if (!keymap.TryGetValue(actionId, out var key) || key.Length == 0)
            {
                continue;
            }

            try
            {
                wanted.Add(new KeyValuePair<string, KeyChord>(actionId, KeyChord.Parse(key)));
            }
            catch (FormatException ex)
            {
                failures.Add($"{actionId}: {ex.Message}");
            }
        }

        var keep = new HashSet<string>(StringComparer.Ordinal);
        foreach (var entry in wanted)
        {
            keep.Add(entry.Value.ToOnKeyString());
        }

        foreach (var onKey in new List<string>(Registered))
        {
            if (!keep.Contains(onKey))
            {
                Restore(onKey);
                Registered.Remove(onKey);
            }
        }

        ByAction.Clear();
        foreach (var entry in wanted)
        {
            var onKey = entry.Value.ToOnKeyString();
            var error = Macros.TryGetValue(entry.Key, out var macro)
                ? Register(entry.Value, macro)
                : "the add-in has no command for this action";
            if (error is null)
            {
                Registered.Add(onKey);
                ByAction[entry.Key] = entry.Value;
            }
            else
            {
                // A key we held may still run whatever it was bound to before; hand it back to Excel.
                if (Registered.Contains(onKey))
                {
                    Restore(onKey);
                    Registered.Remove(onKey);
                }

                failures.Add($"{entry.Value} ({entry.Key}): {error}");
            }
        }

        DiagnosticsLog.Write("Keys", $"bound={Registered.Count}", $"failures={failures.Count}");
        foreach (var failure in failures)
        {
            DiagnosticsLog.Write("KeyFailure", failure);
        }

        return failures;
    }

    /// <summary>
    /// Self-check: the actions in <see cref="ActionIds.All"/> that have no command in the macro map (empty if the
    /// map is complete). <see cref="Apply"/> skips such actions and reports them as failures.
    /// </summary>
    public static IReadOnlyList<string> ActionsWithoutCommand()
    {
        var missing = new List<string>();
        foreach (var actionId in ActionIds.All)
        {
            if (!Macros.ContainsKey(actionId))
            {
                missing.Add(actionId);
            }
        }

        return missing;
    }

    /// <summary>Restores Excel's default for every key we bound (AutoClose). Never throws.</summary>
    public static void Clear()
    {
        // Known limitation: xlcOnKey with no macro restores Excel's default even if another add-in rebound the
        // key after us, silently unbinding theirs.
        foreach (var onKey in Registered)
        {
            Restore(onKey);
        }

        Registered.Clear();
        ByAction.Clear();
    }

    /// <summary>The key bound to <paramref name="actionId"/> for display (e.g. <c>Ctrl+Shift+1</c>), or <c>(no key)</c>.</summary>
    public static string KeyFor(string actionId) =>
        ByAction.TryGetValue(actionId, out var chord) ? chord.ToDisplayString() : "(no key)";

    /// <summary>Binds <paramref name="chord"/> to <paramref name="macro"/>. Returns null on success, else the reason.</summary>
    private static string? Register(KeyChord chord, string macro)
    {
        try
        {
            // The C API can "succeed" yet return FALSE or an error value for a rejected key string.
            var result = XlCall.Excel(XlCall.xlcOnKey, chord.ToOnKeyString(), macro);
            if (result is ExcelError || (result is bool ok && !ok))
            {
                return $"xlcOnKey returned {result}";
            }

            return null;
        }
        catch (Exception ex)
        {
            return ex.Message;
        }
    }

    private static void Restore(string onKey)
    {
        try
        {
            // xlcOnKey with no macro restores Excel's default behavior for the key.
            XlCall.Excel(XlCall.xlcOnKey, onKey);
        }
        catch (XlCallException)
        {
            // Excel is shutting down or the key was never registered; nothing to restore.
        }
    }
}
