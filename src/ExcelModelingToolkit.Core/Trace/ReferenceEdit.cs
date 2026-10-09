using System;
using System.Collections.Generic;
using ExcelModelingToolkit.Core.Undo;

namespace ExcelModelingToolkit.Core.Trace;

/// <summary>
/// F2 in the Trace In window (docs/PLAN.md section 4.5): which reference the selected row edits, and the keys that
/// put Excel in Point mode on it. Excel has no API for the edit caret, so the add-in goes to the cell whose formula
/// holds the reference and sends Excel these keys, as if typed: F2 (Edit mode, the caret at the end), the caret to the
/// reference, Shift+Right over it, and F2 again (Excel's Point mode: the arrow keys and sheet tabs now replace the
/// selected reference with the cell they point at). Point mode starts pointing at the edited cell, so when the
/// reference's target is known (<see cref="GoToText"/>) the keys then go there: for a target in another workbook first
/// Ctrl+Tab (<see cref="SwitchesWindow"/>: Point mode moves to Excel's previously active window, which the add-in has
/// made the target's), then F5 (Excel's Go To dialog, which opens in Point mode too), the target typed, Enter. Excel
/// puts the reference Go To points at in place of the selected text, and the arrow keys now move from the target.
/// </summary>
/// <remarks>
/// <para>
/// Go To into another workbook's window from Point mode is unreliable: it points the first time or two, then
/// intermittently only inserts the typed text without switching windows, whether the keys come from the add-in or from
/// outside (they demonstrably reach the dialog; found in Excel 2026-10-09). Go To within the active window has been
/// reliable throughout. Ctrl+Tab in Point mode reliably switches to the previously active workbook window, pointing at
/// that window's A1 (5 of 5 runs, also with a third workbook open). So for another workbook's target the add-in first
/// makes the target's window the previously active one (it goes to the target, then back to the edited cell), and the
/// keys switch windows with Ctrl+Tab and then go to the target with a Go To within that window (<c>'Sheet'!B3</c>, no
/// workbook).
/// </para>
/// <para>
/// Go To writes the reference in Point mode's own form: relative, or absolute into another workbook
/// (<c>[Book.xlsx]Sheet!$B$3</c>). So an anchored (<c>$</c>) or other-workbook reference that the user commits with
/// Enter without moving takes that form, not the original anchoring (any arrow move rewrites the reference in that
/// form anyway).
/// </para>
/// </remarks>
public static class ReferenceEdit
{
    /// <summary>
    /// The most key events (presses and releases) one F2 sends; a reference further into a longer formula is not
    /// selected (F2 then edits the cell as usual).
    /// </summary>
    public const int MaxKeyEvents = 4000;

    /// <summary>The longest reference typed into the Go To dialog (its Reference box holds 255 characters).</summary>
    public const int MaxGoToLength = 255;

    private const int VkTab = 0x09;
    private const int VkReturn = 0x0D;
    private const int VkShift = 0x10;
    private const int VkControl = 0x11;
    private const int VkEnd = 0x23;
    private const int VkHome = 0x24;
    private const int VkLeft = 0x25;
    private const int VkRight = 0x27;
    private const int VkF2 = 0x71;
    private const int VkF5 = 0x74;

    /// <summary>A plain F2 (press and release): Excel edits the active cell, as without the window.</summary>
    public static IReadOnlyList<SyntheticKey> PlainF2 { get; } = new[] { SyntheticKey.Down(VkF2), SyntheticKey.Up(VkF2) };

    /// <summary>Ctrl+Tab (Ctrl down, Tab press and release, Ctrl up): in Point mode, Excel's previously active window.</summary>
    public static IReadOnlyList<SyntheticKey> CtrlTab { get; } =
        new[] { SyntheticKey.Down(VkControl), SyntheticKey.Down(VkTab), SyntheticKey.Up(VkTab), SyntheticKey.Up(VkControl) };

    /// <summary>
    /// The reference F2 edits for the selected row <paramref name="node"/>: the nearest of the node and its ancestors
    /// that is written in a cell's formula (<see cref="PrecedentItem.Span"/>), going up only from a row that stands for
    /// its parent's reference (a cell of a range, a name's or table reference's target or the references in a name's
    /// formula). Null for the audited cell, an error, a "more cells" or "truncated" row, a computed reference
    /// (INDEX, OFFSET, ...) and what it points to, a function or group, and Excel's same-sheet precedents (a formula
    /// that could not be parsed): F2 then edits the active cell as usual.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="node"/> is null.</exception>
    public static ReferenceSpan? SpanOf(PrecedentNode node) => SpanNodeOf(node)?.Item.Span;

    /// <summary>
    /// The row whose item holds the reference F2 edits for the selected row <paramref name="node"/> (see
    /// <see cref="SpanOf"/>): the node itself, or the ancestor it stands for (the range a cell is in, say); null when
    /// F2 edits the active cell as usual. Its item is the reference's resolved target (<see cref="GoToText"/>).
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="node"/> is null.</exception>
    public static PrecedentNode? SpanNodeOf(PrecedentNode node)
    {
        if (node is null)
        {
            throw new ArgumentNullException(nameof(node));
        }

        switch (node.Item.Kind)
        {
            case PrecedentKind.Error:
            case PrecedentKind.MoreCells:
            case PrecedentKind.Truncated:
                return null;
        }

        for (var current = node; current is not null; current = current.Parent)
        {
            if (current.Item.Kind == PrecedentKind.Function || current.Item.Kind == PrecedentKind.Group)
            {
                return null;
            }

            if (current.Item.Span is not null)
            {
                return current;
            }

            if (current.Parent is null || !StandsForItsChildren(current.Parent.Item.Kind))
            {
                return null;
            }
        }

        return null;
    }

    /// <summary>
    /// What the Go To step types for the reference of <paramref name="span"/> whose resolved target is
    /// <paramref name="target"/> (the item that holds the span), so that Point mode moves from the target: whenever
    /// the target is a plain cell or range. Null for no Go To step where Go To would put back another reference: a
    /// name or table reference (Go To would replace the name with an address), a 3-D reference, a spill (<c>A1#</c>),
    /// or several areas. The target as Excel resolved it: its address on the formula's own sheet (<c>A2</c>,
    /// <c>B2:B5</c>); <c>'Sheet Name'!B4</c> on another sheet of the formula's workbook, and in another workbook too,
    /// with no workbook, since the keys switch to the target's window first (<see cref="SwitchesWindow"/>) and Go To
    /// stays in the active window (always quoted, <c>'</c> doubled: Go To accepts quotes where they are not needed). The
    /// reference Go To writes back has Point mode's anchoring, not necessarily the original's (see the remarks). The
    /// caller makes sure the target's workbook is open and its sheet visible (else Go To fails with a message box).
    /// </summary>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    public static string? GoToText(ReferenceSpan span, PrecedentItem target)
    {
        if (span is null)
        {
            throw new ArgumentNullException(nameof(span));
        }

        if (target is null)
        {
            throw new ArgumentNullException(nameof(target));
        }

        var workbook = target.Workbook;
        var sheetName = target.Sheet;
        var address = target.Address;

        // A spill (A1#) resolves to the spill range, which Go To would write in its place.
        if ((target.Kind != PrecedentKind.Cell && target.Kind != PrecedentKind.Range) || string.IsNullOrEmpty(workbook) ||
            string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(address) || sheetName!.IndexOf(':') >= 0 ||
            address!.IndexOfAny(new[] { ',', '$', '#', '!' }) >= 0 || workbook!.IndexOfAny(new[] { '[', ']' }) >= 0 ||
            TextAfterPrefix(span.Text).IndexOf('#') >= 0)
        {
            return null;
        }

        var text = string.Equals(workbook, span.Workbook, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(sheetName, span.Sheet, StringComparison.OrdinalIgnoreCase)
            ? address
            : "'" + sheetName.Replace("'", "''") + "'!" + address;
        return text.Length > MaxGoToLength ? null : text;
    }

    /// <summary>
    /// True when there is a Go To step for <paramref name="target"/> (<see cref="GoToText"/> is not null) and the
    /// target is in another workbook than the formula of <paramref name="span"/>: the keys then switch to the target's
    /// window with Ctrl+Tab before the Go To (see the remarks), and the caller makes that window Excel's previously
    /// active one first.
    /// </summary>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    public static bool SwitchesWindow(ReferenceSpan span, PrecedentItem target) =>
        GoToText(span, target) is not null && !string.Equals(target.Workbook, span.Workbook, StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// <paramref name="span"/> moved into <paramref name="formula"/>, the owner cell's formula now, when that differs
    /// from the traced one only in how its references are written, not in what they are (a closed workbook's path,
    /// which Excel leaves out once the trace has opened the workbook); the span itself if the formula is the same;
    /// null if the formula changed otherwise or either cannot be parsed.
    /// </summary>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    public static ReferenceSpan? Relocate(ReferenceSpan span, string formula)
    {
        if (span is null)
        {
            throw new ArgumentNullException(nameof(span));
        }

        if (formula is null)
        {
            throw new ArgumentNullException(nameof(formula));
        }

        if (string.Equals(formula, span.Formula, StringComparison.Ordinal))
        {
            return span;
        }

        var context = new FormulaContext(span.Workbook, span.Sheet);
        var before = FormulaParser.Parse(span.Formula, context);
        var after = FormulaParser.Parse(formula, context);
        if (!before.IsParsed || !after.IsParsed || before.References.Count != after.References.Count)
        {
            return null;
        }

        FormulaReference? moved = null;
        var endBefore = 0;
        var endAfter = 0;
        for (var i = 0; i < before.References.Count; i++)
        {
            var b = before.References[i];
            var a = after.References[i];
            if (b.Start < endBefore || a.Start < endAfter || !SameReference(b, a) ||
                !string.Equals(span.Formula.Substring(endBefore, b.Start - endBefore), formula.Substring(endAfter, a.Start - endAfter), StringComparison.Ordinal))
            {
                return null;
            }

            if (b.Start == span.Start && b.Length == span.Length)
            {
                moved = a;
            }

            endBefore = b.Start + b.Length;
            endAfter = a.Start + a.Length;
        }

        if (moved is null || !string.Equals(span.Formula.Substring(endBefore), formula.Substring(endAfter), StringComparison.Ordinal))
        {
            return null;
        }

        return new ReferenceSpan(span.Workbook, span.Sheet, span.Address, formula, moved.Start, moved.Length);
    }

    /// <summary>
    /// The keys that select <paramref name="span"/>'s reference in Excel's editor, switch to Point mode and, if
    /// <paramref name="goTo"/> is given, go to it, after Ctrl+Tab if <paramref name="switchWindow"/> (see
    /// <see cref="Keys(int, int, int, string?, bool)"/>), or null if that takes more than <see cref="MaxKeyEvents"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="span"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="goTo"/> is empty, or <paramref name="switchWindow"/> is true without <paramref name="goTo"/>.
    /// </exception>
    public static IReadOnlyList<SyntheticKey>? Keys(ReferenceSpan span, string? goTo = null, bool switchWindow = false)
    {
        if (span is null)
        {
            throw new ArgumentNullException(nameof(span));
        }

        return Keys(span.Formula.Length, span.Start, span.Length, goTo, switchWindow);
    }

    /// <summary>
    /// The keys that select the text from <paramref name="start"/> (0-based, the <c>=</c> being 0) for
    /// <paramref name="length"/> characters of a cell's formula of <paramref name="textLength"/> characters, as Excel's
    /// editor shows it, and switch to Point mode; null if that takes more than <see cref="MaxKeyEvents"/> key events.
    /// In order: F2 (Edit mode, the caret at the end); the caret to <paramref name="start"/> from the nearer end, so
    /// Ctrl+Home then Right <paramref name="start"/> times, or Ctrl+End then Left
    /// (<paramref name="textLength"/> - <paramref name="start"/>) times; Shift+Right <paramref name="length"/> times
    /// (the selection always runs forward, so the caret ends after the reference); F2 (Point mode). Then, if
    /// <paramref name="switchWindow"/>, <see cref="CtrlTab"/> (Excel's previously active window: see
    /// <see cref="SwitchesWindow"/>); then, if <paramref name="goTo"/> is given, the Go To step: F5 (Go To), each of
    /// its characters typed (<see cref="SyntheticKey.CharacterDown"/>), and Enter (see <see cref="GoToText"/>).
    /// Nothing follows the last Enter.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">
    /// The text is not inside the formula, or is empty; <paramref name="goTo"/> is empty.
    /// </exception>
    /// <exception cref="ArgumentException"><paramref name="switchWindow"/> is true without <paramref name="goTo"/>.</exception>
    public static IReadOnlyList<SyntheticKey>? Keys(int textLength, int start, int length, string? goTo = null, bool switchWindow = false)
    {
        if (start < 0 || length < 1 || start > textLength - length)
        {
            throw new ArgumentOutOfRangeException(nameof(start), start, "The text must lie inside the formula.");
        }

        if (goTo is not null && goTo.Length == 0)
        {
            throw new ArgumentOutOfRangeException(nameof(goTo), "The Go To text cannot be empty.");
        }

        if (switchWindow && goTo is null)
        {
            throw new ArgumentException("Switching windows needs the Go To in the target's window.", nameof(switchWindow));
        }

        var fromStart = start <= textLength - start;
        var moves = fromStart ? start : textLength - start;

        // F2, Ctrl+Home/End, Shift and F2: 10 events; each Right or Left: 2. Ctrl+Tab: 4. The Go To: F5 and Enter, 4
        // events; each character typed: 2.
        var events = 10 + (2L * moves) + (2L * length) + (switchWindow ? CtrlTab.Count : 0) +
            (goTo is null ? 0 : 4 + (2L * goTo.Length));
        if (events > MaxKeyEvents)
        {
            return null;
        }

        var keys = new List<SyntheticKey>((int)events);
        Press(keys, VkF2);
        keys.Add(SyntheticKey.Down(VkControl));
        Press(keys, fromStart ? VkHome : VkEnd);
        keys.Add(SyntheticKey.Up(VkControl));
        for (var i = 0; i < moves; i++)
        {
            Press(keys, fromStart ? VkRight : VkLeft);
        }

        keys.Add(SyntheticKey.Down(VkShift));
        for (var i = 0; i < length; i++)
        {
            Press(keys, VkRight);
        }

        keys.Add(SyntheticKey.Up(VkShift));
        Press(keys, VkF2);
        if (switchWindow)
        {
            keys.AddRange(CtrlTab);
        }

        if (goTo is not null)
        {
            GoTo(keys, goTo);
        }

        return keys;
    }

    /// <summary>
    /// <paramref name="keys"/> in the batches they are sent in: a batch ends after each F5 release (F5 press and
    /// release are in the same batch) and after each Ctrl+Tab (<see cref="EndsWithCtrlTab"/>) that is not the last key.
    /// So what is typed into a Go To dialog, and the Enter after it, is a batch of its own, sent once the dialog has the
    /// keyboard focus: characters that reach Excel before the dialog's box has it are lost, and Enter then goes to the
    /// dialog's previous reference (found in Excel 2026-10-09). And the F5 after Ctrl+Tab is a batch of its own, sent
    /// once the target's window is in front. Without F5 or Ctrl+Tab, one batch. The batches together are
    /// <paramref name="keys"/>, in order.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="keys"/> is null.</exception>
    public static IReadOnlyList<IReadOnlyList<SyntheticKey>> Batches(IReadOnlyList<SyntheticKey> keys)
    {
        if (keys is null)
        {
            throw new ArgumentNullException(nameof(keys));
        }

        var batches = new List<IReadOnlyList<SyntheticKey>>();
        var batch = new List<SyntheticKey>();
        foreach (var key in keys)
        {
            batch.Add(key);
            if (key.Equals(SyntheticKey.Up(VkF5)) || EndsWithCtrlTab(batch))
            {
                batches.Add(batch);
                batch = new List<SyntheticKey>();
            }
        }

        if (batch.Count > 0)
        {
            batches.Add(batch);
        }

        return batches;
    }

    /// <summary>True if <paramref name="keys"/> ends with <see cref="CtrlTab"/>'s Tab release and Ctrl release.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="keys"/> is null.</exception>
    public static bool EndsWithCtrlTab(IReadOnlyList<SyntheticKey> keys)
    {
        if (keys is null)
        {
            throw new ArgumentNullException(nameof(keys));
        }

        return keys.Count >= 2 && keys[keys.Count - 1].Equals(SyntheticKey.Up(VkControl)) &&
            keys[keys.Count - 2].Equals(SyntheticKey.Up(VkTab));
    }

    // A row under a range (a page of its cells), a name or a table reference is reached through that reference.
    private static bool StandsForItsChildren(PrecedentKind kind) =>
        kind == PrecedentKind.Range || kind == PrecedentKind.Name || kind == PrecedentKind.Table;

    // Two parses of the same reference, however it is written (a closed workbook's path or not).
    private static bool SameReference(FormulaReference a, FormulaReference b) =>
        a.Kind == b.Kind && a.IsSpill == b.IsSpill &&
        string.Equals(a.WorkbookName, b.WorkbookName, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(a.Sheet, b.Sheet, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(a.LastSheet, b.LastSheet, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(a.Address, b.Address, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(a.Name, b.Name, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(TextAfterPrefix(a.Text), TextAfterPrefix(b.Text), StringComparison.Ordinal);

    // The reference as written after its sheet or workbook prefix (A1, $B$2:$B$5, Name, Table[Col]).
    private static string TextAfterPrefix(string text)
    {
        var bang = text.LastIndexOf('!');
        return bang < 0 ? text : text.Substring(bang + 1);
    }

    // One Go To: F5, the text typed, Enter.
    private static void GoTo(List<SyntheticKey> keys, string text)
    {
        Press(keys, VkF5);
        foreach (var character in text)
        {
            keys.Add(SyntheticKey.CharacterDown(character));
            keys.Add(SyntheticKey.CharacterUp(character));
        }

        Press(keys, VkReturn);
    }

    private static void Press(List<SyntheticKey> keys, int virtualKey)
    {
        keys.Add(SyntheticKey.Down(virtualKey));
        keys.Add(SyntheticKey.Up(virtualKey));
    }
}
