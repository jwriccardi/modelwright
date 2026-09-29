using System;
using System.Diagnostics;
using System.Globalization;
using ExcelDna.Integration;
using ExcelModelingToolkit.Core.Undo;

namespace ExcelModelingToolkit.AddIn;

/// <summary>
/// Undoes or redoes our last formatting change from <see cref="Session.Undo"/>. Before writing, it checks that
/// every block still holds what our change left there (for a redo: what the undo put back); if anything differs
/// (the cells were edited, or rows or columns were inserted or deleted so other cells moved in), it refuses and
/// drops that step rather than write formats onto the wrong cells. Call in macro context on the main thread (the
/// status bar uses the C API). The restore's own COM writes clear Excel's native undo history, like the cycle's.
/// </summary>
internal static class UndoCommand
{
    /// <summary>
    /// Undoes (<see cref="UndoKey.Undo"/>) or redoes (<see cref="UndoKey.Redo"/>) the top snapshot, shows the result
    /// in the status bar and writes one diagnostics log line. <paramref name="source"/> is how it was invoked
    /// (<c>key</c> or <c>ribbon</c>). Never throws.
    /// </summary>
    public static void Run(UndoKey key, string source)
    {
        var stopwatch = Stopwatch.StartNew();
        string label = "-";
        var blocks = 0;
        string result;
        try
        {
            result = Execute(key, ref label, ref blocks);
        }
        catch (Exception ex)
        {
            result = "error: " + ex.Message;
            StatusBar.Show($"{ProductInfo.Name}: {Verb(key)} formatting failed: {ex.Message}");
        }

        DiagnosticsLog.Write(
            "UndoRestore",
            Verb(key),
            "source=" + source,
            "ms=" + stopwatch.Elapsed.TotalMilliseconds.ToString("0.0", CultureInfo.InvariantCulture),
            "cycle=" + label,
            "blocks=" + blocks.ToString(CultureInfo.InvariantCulture),
            "undoLeft=" + Session.Undo.UndoCount.ToString(CultureInfo.InvariantCulture),
            "redoLeft=" + Session.Undo.RedoCount.ToString(CultureInfo.InvariantCulture),
            result);
    }

    private static string Execute(UndoKey key, ref string label, ref int blocks)
    {
        var verb = Verb(key);
        var snapshot = Session.Undo.Peek(key);
        if (snapshot is null)
        {
            StatusBar.Show($"{ProductInfo.Name}: no formatting change to {verb}.");
            return "nothing to " + verb;
        }

        label = snapshot.Label;
        blocks = snapshot.Blocks.Count;

        dynamic app = ExcelDnaUtil.Application;
        object workbooks = app.Workbooks;
        var workbook = FindByName(workbooks, snapshot.Workbook);
        if (workbook is null)
        {
            var dropped = Session.Undo.InvalidateWorkbook(snapshot.Workbook);
            StatusBar.Show($"Can't {verb}: {snapshot.Workbook} is no longer open (under that name).");
            return $"refused: workbook missing (dropped {dropped})";
        }

        dynamic book = workbook;
        object worksheets = book.Worksheets;
        var worksheet = FindByName(worksheets, snapshot.Sheet);
        if (worksheet is null)
        {
            var dropped = Session.Undo.InvalidateSheet(snapshot.Workbook, snapshot.Sheet);
            StatusBar.Show($"Can't {verb}: sheet '{snapshot.Sheet}' was renamed or deleted.");
            return $"refused: sheet missing (dropped {dropped})";
        }

        if (CellFormats.FormattingIsProtected(worksheet))
        {
            // Kept: unprotecting the sheet and trying again works.
            StatusBar.Show($"Can't {verb}: sheet '{snapshot.Sheet}' is protected.");
            return "refused: sheet protected";
        }

        var reader = new SheetFormatReader(worksheet, snapshot.Kind);
        foreach (var block in snapshot.Blocks)
        {
            var found = reader.ReadUniform(block.Range);
            var expected = block.Expected(key);
            if (found != expected)
            {
                Session.Undo.Discard(key);
                StatusBar.Show($"Can't {verb}: the cells changed since (e.g. rows inserted or edited).");
                return $"refused: {block.Address} holds {found}, expected {expected}";
            }
        }

        try
        {
            dynamic sheet = worksheet;
            foreach (var group in snapshot.WriteGroups(key))
            {
                object range = sheet.Range(group.Address);
                CellFormats.Write(range, snapshot.Kind, group.Value);
            }
        }
        catch (Exception ex)
        {
            // Some groups may already be written; this step cannot be trusted any more.
            Session.Undo.Discard(key);
            StatusBar.Show($"Can't {verb}: Excel refused the change, so this step was dropped. ({ex.Message})");
            return "error while writing: " + ex.Message;
        }

        if (key == UndoKey.Undo)
        {
            Session.Undo.Undo();
        }
        else
        {
            Session.Undo.Redo();
        }

        var left = Session.Undo.Count(key).ToString(CultureInfo.InvariantCulture);
        StatusBar.Show($"{(key == UndoKey.Undo ? "Undo" : "Redo")}: {snapshot.Label} ({left} left)");
        return "ok";
    }

    /// <summary>The item of a COM collection (<c>Workbooks</c>, <c>Worksheets</c>) named <paramref name="name"/> (ignoring case, like Excel), or null.</summary>
    private static object? FindByName(object collection, string name)
    {
        dynamic items = collection;
        int count = items.Count;
        for (var i = 1; i <= count; i++)
        {
            object item = items.Item(i);
            dynamic named = item;
            string itemName = named.Name;
            if (string.Equals(itemName, name, StringComparison.OrdinalIgnoreCase))
            {
                return item;
            }
        }

        return null;
    }

    private static string Verb(UndoKey key) => key == UndoKey.Undo ? "undo" : "redo";
}
