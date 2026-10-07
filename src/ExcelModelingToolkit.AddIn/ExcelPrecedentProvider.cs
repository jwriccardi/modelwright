using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelModelingToolkit.Core.Trace;
using ExcelModelingToolkit.Core.Undo;

namespace ExcelModelingToolkit.AddIn;

/// <summary>
/// Reads the Trace In tree from Excel (docs/PLAN.md section 4.5): a thin COM adapter over the Core trace logic.
/// A cell's precedents come from its <c>Range.Formula</c> (en-US, A1) parsed by <see cref="FormulaParser"/>; each
/// reference is resolved to cells: A1 references through <c>Range</c>, names through <c>Names</c> (a name with no
/// cells shows its value), table references through <c>ListObjects</c> and <see cref="TableLayout"/>, and the
/// calls that compute a reference (INDEX, OFFSET, INDIRECT, CHOOSE) through <c>Worksheet.Evaluate</c> in the
/// formula's sheet. A formula the parser rejects falls back to Excel's <c>Range.DirectPrecedents</c> (same sheet
/// only, and labelled so).
/// </summary>
/// <remarks>
/// <para>
/// Main thread: in the Trace In macro, or from Excel's message loop while navigating (see
/// <see cref="TraceSession"/>). No member throws: a failure becomes a <see cref="PrecedentKind.Error"/> item that
/// says why. Nothing here writes to a cell, with one exception: a precedent in a closed workbook opens that workbook
/// (read-only, links not updated, macros disabled, no prompts), which is what Macabacus does; the workbook stays
/// open. Excel may treat the open as an action that clears its undo history.
/// </para>
/// <para>
/// Values come from <c>Range.Text</c> (as displayed) for a reference's first cell, and from one <c>Value2</c> read
/// per block for a page of a range's cells (so those show unformatted). COM objects are not released by hand, as
/// elsewhere in the add-in.
/// </para>
/// </remarks>
internal sealed class ExcelPrecedentProvider : IPrecedentProvider
{
    private const int XlSheetVisible = -1;
    private const int MsoAutomationSecurityForceDisable = 3;

    // Worksheet.Evaluate accepts at most 255 characters.
    private const int MaxEvaluateLength = 255;

    // A password no workbook has: opening a password-protected workbook then fails at once instead of prompting.
    // An unprotected workbook ignores it.
    private const string NotAPassword = "emt-not-a-password-7c1e";

    private static readonly IReadOnlyList<PrecedentItem> None = new PrecedentItem[0];

    // What each item stands for in Excel. PrecedentItem compares by reference, as wanted here.
    private readonly Dictionary<PrecedentItem, Target> _targets = new Dictionary<PrecedentItem, Target>();

    // "hidden workbook" / "hidden sheet" / "" (visible), by workbook|sheet, read once per trace.
    private readonly Dictionary<string, string> _sheetNotes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

    private readonly CultureInfo _culture = CultureInfo.CurrentCulture;

    /// <summary>The workbooks this trace opened (file names), for the status bar.</summary>
    public List<string> OpenedWorkbooks { get; } = new List<string>();

    /// <summary>Formulas that could not be parsed (their cells, with the reason), for the status bar and log.</summary>
    public List<string> ParseFallbacks { get; } = new List<string>();

    /// <summary>The item for the audited cell (the tree's root). Throws if the cell cannot be read.</summary>
    public PrecedentItem CreateRoot(object cell)
    {
        dynamic range = cell;
        object sheet = range.Worksheet;
        dynamic ws = sheet;
        object book = ws.Parent;
        dynamic wb = book;
        string workbookName = wb.Name;
        string sheetName = ws.Name;
        var local = LocalAddress(cell);
        var item = new PrecedentItem(PrecedentKind.Cell, sheetName + "!" + local, workbookName, sheetName, local, 1,
            CellValue(cell), canExpand: true, hiddenNote: HiddenNote(cell, RectOf(cell), book, sheet, workbookName, sheetName));
        _targets[item] = Target.ForRange(cell, new Area(cell, sheet, book, workbookName, sheetName, RectOf(cell)));
        return item;
    }

    /// <summary>
    /// The formula of a cell item (the root, or a cell found in the tree) and its context, or null if it is not a
    /// cell with a formula. Never throws.
    /// </summary>
    public string? ReadFormula(PrecedentItem item, out FormulaContext? context)
    {
        context = null;
        try
        {
            if (!_targets.TryGetValue(item, out var target) || target.Areas is null || target.Areas.Count == 0)
            {
                return null;
            }

            var area = target.Areas[0];
            dynamic cell = FirstCell(area.Range);
            object formula = cell.Formula;
            if (formula is string text && text.StartsWith("=", StringComparison.Ordinal))
            {
                context = new FormulaContext(area.WorkbookName, area.SheetName);
                return text;
            }

            return null;
        }
        catch (Exception)
        {
            return null;
        }
    }

    /// <summary>
    /// The range Excel goes to for <paramref name="item"/> (a name's, table reference's or computed reference's
    /// target, a merged cell's merge area), or null with the reason in <paramref name="reason"/> (null when the item is
    /// simply not a place, like a "more cells" row). Never throws.
    /// </summary>
    public object? NavigationRange(PrecedentItem item, out string? reason)
    {
        reason = null;
        switch (item.Kind)
        {
            case PrecedentKind.MoreCells:
            case PrecedentKind.Truncated:
            case PrecedentKind.Function:
            case PrecedentKind.Group:
                return null;
            case PrecedentKind.Error:
                reason = item.Label + ": " + (item.ValueText ?? "cannot be followed");
                return null;
        }

        try
        {
            if (_targets.TryGetValue(item, out var target) && target.GetRange() is object range)
            {
                return range;
            }
        }
        catch (Exception ex)
        {
            reason = item.Label + ": " + ex.Message;
            return null;
        }

        reason = item.Label + (item.ValueText is null ? " has no cells to go to." : " has no cells to go to (its value is " + item.ValueText + ").");
        return null;
    }

    /// <inheritdoc />
    public IReadOnlyList<PrecedentItem> GetPrecedents(PrecedentItem item)
    {
        try
        {
            if (!_targets.TryGetValue(item, out var target))
            {
                return None;
            }

            if (target.Child is not null)
            {
                return new[] { target.Child };
            }

            if (target.Areas is not null && target.Areas.Count == 1 &&
                (item.Kind == PrecedentKind.Cell || item.CellCount == 1))
            {
                return CellPrecedents(target.Areas[0]);
            }

            return None;
        }
        catch (Exception ex)
        {
            return new[] { ErrorItem(item.Label, "could not read its precedents: " + ex.Message) };
        }
    }

    /// <inheritdoc />
    public IReadOnlyList<PrecedentItem> GetRangeCells(PrecedentItem range, long start, int count)
    {
        try
        {
            if (!_targets.TryGetValue(range, out var target) || target.Areas is null)
            {
                return None;
            }

            var rects = new List<CellRect>(target.Areas.Count);
            var oneSheet = true;
            foreach (var area in target.Areas)
            {
                rects.Add(area.Rect);
                oneSheet &= string.Equals(area.SheetName, target.Areas[0].SheetName, StringComparison.OrdinalIgnoreCase);
            }

            var cells = new List<PrecedentItem>(count);
            foreach (var block in RangePaging.Plan(rects, start, count))
            {
                var area = target.Areas[block.Area];
                dynamic sheet = area.Sheet;
                object blockRange = sheet.Range(block.Block.Address);
                dynamic r = blockRange;
                object values = r.Value2;
                object formulas = r.Formula;
                var hidden = SheetNote(area.Book, area.Sheet, area.WorkbookName, area.SheetName);
                for (var i = 0; i < block.Take; i++)
                {
                    var (row, column) = block.Offset(i);
                    var local = CellRect.ColumnName(block.Block.Column + column) +
                        (block.Block.Row + row).ToString(CultureInfo.InvariantCulture);
                    var formula = Element(formulas, row, column) as string;
                    var item = new PrecedentItem(PrecedentKind.Cell, oneSheet ? local : area.SheetName + "!" + local,
                        area.WorkbookName, area.SheetName, local, 1,
                        TraceValueText.FromValue(Element(values, row, column), _culture),
                        canExpand: formula is not null && formula.StartsWith("=", StringComparison.Ordinal),
                        hiddenNote: hidden.Length == 0 ? null : hidden);
                    _targets[item] = Target.ForCell(area, local);
                    cells.Add(item);
                }
            }

            return cells;
        }
        catch (Exception ex)
        {
            return new[] { ErrorItem(range.Label, "could not read its cells: " + ex.Message) };
        }
    }

    // A cell's precedents: its formula's references in the order written, or Excel's same-sheet direct precedents
    // if the formula cannot be parsed.
    private IReadOnlyList<PrecedentItem> CellPrecedents(Area area)
    {
        object cell = FirstCell(area.Range);
        dynamic c = cell;
        object value = c.Formula;
        if (value is not string formula || !formula.StartsWith("=", StringComparison.Ordinal))
        {
            return None;
        }

        var context = new FormulaContext(area.WorkbookName, area.SheetName);
        var parsed = FormulaParser.Parse(formula, context);
        if (!parsed.IsParsed)
        {
            return DirectPrecedents(cell, area, parsed.Error ?? "unknown error");
        }

        var items = new List<PrecedentItem>();
        foreach (var row in ClassicPrecedents.Of(parsed))
        {
            items.Add(row.Dynamic is not null
                ? ResolveDynamic(row.Dynamic, cell, area)
                : ResolveReference(row.Reference!, context, cell, area));
        }

        return items;
    }

    private PrecedentItem ResolveReference(FormulaReference reference, FormulaContext context, object formulaCell, Area formulaArea)
    {
        try
        {
            switch (reference.Kind)
            {
                case FormulaReferenceKind.Cell:
                case FormulaReferenceKind.Range:
                case FormulaReferenceKind.WholeColumn:
                case FormulaReferenceKind.WholeRow:
                    return ResolveA1(reference, context, formulaArea);
                case FormulaReferenceKind.Name:
                    return ResolveName(reference, context, formulaCell, formulaArea);
                case FormulaReferenceKind.StructuredReference:
                    return ResolveTable(reference, reference.Name, reference.TableSpecifiers, reference.TableColumns, formulaCell, formulaArea);
                case FormulaReferenceKind.RefError:
                    return new PrecedentItem(PrecedentKind.Error, reference.Text, valueText: "#REF!", canExpand: false);
                default:
                    return ErrorItem(reference.Text, "not a reference that can be traced");
            }
        }
        catch (Exception ex)
        {
            return ErrorItem(reference.Text, ex.Message);
        }
    }

    // Cell, range, whole column or row, 3-D (Sheet1:Sheet3!A1) and spill (A1#) references, here or in another
    // workbook (opened if closed).
    private PrecedentItem ResolveA1(FormulaReference reference, FormulaContext context, Area formulaArea)
    {
        string? bookError = null;
        var book = reference.WorkbookName is null
            ? formulaArea.Book
            : FindWorkbook(reference.WorkbookName, reference.WorkbookPath, mayOpen: true, formulaArea.Book, out bookError);
        if (book is null)
        {
            return ErrorItem(reference.Text, bookError ?? reference.WorkbookName + " is not open");
        }

        dynamic wb = book;
        string workbookName = wb.Name;
        var sheetName = reference.Sheet ?? (reference.WorkbookName is null ? context.SheetName : null);
        if (sheetName is null)
        {
            return ErrorItem(reference.Text, "the reference names no sheet");
        }

        if (reference.Is3D)
        {
            return Resolve3D(reference, book, workbookName);
        }

        var sheet = FindSheet(book, sheetName);
        if (sheet is null)
        {
            return ErrorItem(reference.Text, $"there is no sheet '{sheetName}' in {workbookName}");
        }

        dynamic ws = sheet;
        object range = ws.Range(reference.Address);
        if (reference.IsSpill)
        {
            range = SpillRange(range);
        }

        return RangeItem(reference.Text, range);
    }

    // A1# is the anchor's spill range while it spills, else the anchor itself.
    private static object SpillRange(object anchor)
    {
        try
        {
            dynamic cell = anchor;
            object spill = cell.SpillingToRange;
            return spill ?? anchor;
        }
        catch (Exception)
        {
            return anchor;
        }
    }

    // Sheet1:Sheet3!A1: the same cells on every worksheet from the first sheet to the last, as one range item whose
    // areas are on different sheets (Excel has no single range for them; Goto goes to the first sheet's).
    private PrecedentItem Resolve3D(FormulaReference reference, object book, string workbookName)
    {
        var first = FindSheet(book, reference.Sheet!);
        var last = FindSheet(book, reference.LastSheet!);
        if (first is null || last is null)
        {
            return ErrorItem(reference.Text, $"there is no sheet '{(first is null ? reference.Sheet : reference.LastSheet)}' in {workbookName}");
        }

        dynamic firstSheet = first;
        dynamic lastSheet = last;
        int from = firstSheet.Index;
        int to = lastSheet.Index;
        dynamic wb = book;
        object sheets = wb.Sheets;
        dynamic all = sheets;
        var areas = new List<Area>();
        for (var index = Math.Min(from, to); index <= Math.Max(from, to); index++)
        {
            object sheet = all.Item(index);
            dynamic ws = sheet;
            object type = ws.Type;
            if (Convert.ToInt32(type, CultureInfo.InvariantCulture) != -4167)
            {
                continue; // a chart sheet (xlWorksheet is -4167)
            }

            object range = ws.Range(reference.Address);
            areas.Add(new Area(range, sheet, book, workbookName, (string)ws.Name, RectOf(range)));
        }

        if (areas.Count == 0)
        {
            return ErrorItem(reference.Text, "the sheets it spans hold no worksheets");
        }

        long cells = 0;
        foreach (var area in areas)
        {
            cells += area.Rect.CellCount;
        }

        var basis = PrecedentItem.TryFromReference(reference, new FormulaContext(workbookName, areas[0].SheetName))!;
        var item = new PrecedentItem(cells == 1 ? PrecedentKind.Cell : PrecedentKind.Range, reference.Text, workbookName,
            basis.Sheet, basis.Address, cells, TraceValueText.ForRange(CellValue(areas[0].Range), cells, _culture),
            hiddenNote: HiddenNote(areas[0].Range, areas[0].Rect, book, areas[0].Sheet, workbookName, areas[0].SheetName));
        _targets[item] = new Target(areas[0].Range, areas, null);
        return item;
    }

    private PrecedentItem ResolveName(FormulaReference reference, FormulaContext context, object formulaCell, Area formulaArea)
    {
        foreach (var lookup in NameLookup.Candidates(reference, context))
        {
            var book = lookup.WorkbookName is null
                ? formulaArea.Book
                : FindWorkbook(lookup.WorkbookName, reference.WorkbookPath, lookup.MayOpenWorkbook, formulaArea.Book, out _);
            if (book is null)
            {
                continue;
            }

            var name = lookup.SheetName is null ? WorkbookLevelName(book, lookup.Name) : SheetLevelName(book, lookup.SheetName, lookup.Name);
            if (name is not null)
            {
                return NameItem(reference, context, name, book, lookup);
            }
        }

        // A table's name written alone (=ROWS(Sales)) is its data, not a defined name.
        if (reference.Sheet is null && reference.Name is not null)
        {
            var book = reference.WorkbookName is null
                ? formulaArea.Book
                : FindWorkbook(reference.WorkbookName, reference.WorkbookPath, mayOpen: true, formulaArea.Book, out _);
            if (book is not null && FindTable(book, reference.Name) is not null)
            {
                return ResolveTable(reference, reference.Name, new string[0], new string[0], formulaCell, formulaArea);
            }
        }

        return ErrorItem(reference.Text, $"no name '{reference.Name}' is defined");
    }

    private PrecedentItem NameItem(FormulaReference reference, FormulaContext context, object name, object book, NameLookup lookup)
    {
        var basis = PrecedentItem.TryFromReference(reference, context)!;
        PrecedentItem? child = null;
        string? value;
        object? target = null;
        try
        {
            dynamic n = name;
            target = n.RefersToRange;
        }
        catch (Exception)
        {
            // A constant (=0.05) or a formula (=Rate*2, or one that computes a range): evaluate it below.
        }

        if (target is null)
        {
            var sheet = (lookup.SheetName is null ? null : FindSheet(book, lookup.SheetName)) ?? FirstWorksheet(book);
            dynamic ws = sheet!;
            object result = ws.Evaluate(lookup.Name);
            if (IsRange(result))
            {
                target = result;
            }
            else
            {
                value = TraceValueText.FromValue(result, _culture);
                var leaf = new PrecedentItem(PrecedentKind.Name, reference.Text, basis.Workbook, basis.Sheet,
                    valueText: value, canExpand: false);
                return leaf;
            }
        }

        child = RangeItem(DisplayAddress(target, context.WorkbookName), target);
        value = child.ValueText;
        var item = new PrecedentItem(PrecedentKind.Name, reference.Text, basis.Workbook, basis.Sheet,
            valueText: value, canExpand: true, hiddenNote: child.HiddenNote);
        _targets[item] = Indirect(child);
        return item;
    }

    // Table1[Col], Table1[[#Headers],[A]:[B]], [@Col] (the table holding the formula's cell), Table1 alone.
    private PrecedentItem ResolveTable(FormulaReference reference, string? tableName, IReadOnlyList<string> specifiers,
        IReadOnlyList<string> columns, object formulaCell, Area formulaArea)
    {
        object? table;
        if (tableName is null)
        {
            table = TableOf(formulaCell);
            if (table is null)
            {
                return ErrorItem(reference.Text, "the formula's cell is not in a table");
            }
        }
        else
        {
            var book = reference.WorkbookName is null
                ? formulaArea.Book
                : FindWorkbook(reference.WorkbookName, reference.WorkbookPath, mayOpen: true, formulaArea.Book, out _);
            if (book is null)
            {
                return ErrorItem(reference.Text, $"{reference.WorkbookName} is not open");
            }

            table = FindTable(book, tableName);
            if (table is null)
            {
                return ErrorItem(reference.Text, $"there is no table '{tableName}'");
            }
        }

        dynamic lo = table;
        object parent = lo.Parent;
        dynamic ws = parent;
        string sheetName = ws.Name;
        object book2 = ws.Parent;
        dynamic wb = book2;
        string workbookName = wb.Name;
        var layout = Layout(table);
        int? formulaRow = null;
        if (string.Equals(sheetName, formulaArea.SheetName, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(workbookName, formulaArea.WorkbookName, StringComparison.OrdinalIgnoreCase))
        {
            dynamic cell = formulaCell;
            formulaRow = Convert.ToInt32((object)cell.Row, CultureInfo.InvariantCulture);
        }

        var rect = layout.Resolve(specifiers, columns, formulaRow, out var error);
        if (rect is not CellRect cells)
        {
            return ErrorItem(reference.Text, error ?? "the table reference matches no cells");
        }

        object range = ws.Range(cells.Address);
        var child = RangeItem(DisplayAddress(range, formulaArea.WorkbookName), range);
        var item = new PrecedentItem(PrecedentKind.Table, reference.Text, workbookName, sheetName,
            valueText: child.ValueText, canExpand: true, hiddenNote: child.HiddenNote);
        _targets[item] = Indirect(child);
        return item;
    }

    // INDEX(...), OFFSET(...), INDIRECT(...), CHOOSE(...): evaluated in the formula's sheet; a range result can be
    // expanded and gone to, any other result is shown as the value.
    private PrecedentItem ResolveDynamic(DynamicReference dynamicReference, object formulaCell, Area formulaArea)
    {
        var text = dynamicReference.Text;
        try
        {
            if (text.Length > MaxEvaluateLength)
            {
                return new PrecedentItem(PrecedentKind.DynamicReference, text,
                    valueText: $"(not evaluated: longer than Excel's {MaxEvaluateLength}-character limit)", canExpand: false);
            }

            dynamic ws = formulaArea.Sheet;
            object result = ws.Evaluate(text);
            if (!IsRange(result))
            {
                return new PrecedentItem(PrecedentKind.DynamicReference, text,
                    valueText: TraceValueText.FromValue(result, _culture), canExpand: false);
            }

            var child = RangeItem(DisplayAddress(result, formulaArea.WorkbookName), result);
            var item = new PrecedentItem(PrecedentKind.DynamicReference, text, valueText: child.ValueText,
                canExpand: true, hiddenNote: child.HiddenNote);
            _targets[item] = Indirect(child);
            return item;
        }
        catch (Exception ex)
        {
            return ErrorItem(text, "could not evaluate it: " + ex.Message);
        }
    }

    // The parser rejected the formula: Excel's direct precedents, which only covers the formula's own sheet.
    private IReadOnlyList<PrecedentItem> DirectPrecedents(object cell, Area area, string parseError)
    {
        ParseFallbacks.Add(area.SheetName + "!" + LocalAddress(cell) + ": " + parseError);
        DiagnosticsLog.Write("TraceParseFallback", area.WorkbookName + "|" + area.SheetName + "|" + LocalAddress(cell), parseError);

        // Excel answers DirectPrecedents only for a cell on the active sheet (it is not activated here: that would
        // move the user).
        if (!IsActiveSheet(area.Sheet))
        {
            return new[] { ErrorItem("(formula not parsed)", parseError + "; go to the cell to list its precedents on its sheet") };
        }

        object precedents;
        try
        {
            dynamic c = cell;
            precedents = c.DirectPrecedents;
        }
        catch (Exception)
        {
            // Excel throws when there are none.
            return new[] { ErrorItem("(formula not parsed)", parseError + "; Excel lists no precedents on this sheet") };
        }

        var items = new List<PrecedentItem>();
        dynamic all = precedents;
        object areasObject = all.Areas;
        dynamic areas = areasObject;
        int count = areas.Count;
        for (var i = 1; i <= count; i++)
        {
            object range = areas.Item(i);
            items.Add(RangeItem(LocalAddress(range) + " (same sheet only)", range));
        }

        return items;
    }

    /// <summary>
    /// The item for a range: a cell (merged: its merge area) or a range of several cells or areas, labelled
    /// <paramref name="label"/>, with its first cell's value, whether a cell has a formula, and a hidden badge.
    /// </summary>
    private PrecedentItem RangeItem(string label, object range)
    {
        dynamic r = range;
        object sheet = r.Worksheet;
        dynamic ws = sheet;
        object book = ws.Parent;
        dynamic wb = book;
        string workbookName = wb.Name;
        string sheetName = ws.Name;
        object areasObject = r.Areas;
        dynamic areasCom = areasObject;
        int areaCount = areasCom.Count;
        var areas = new List<Area>(areaCount);
        long cells = 0;
        for (var i = 1; i <= areaCount; i++)
        {
            object area = areaCount == 1 ? range : areasCom.Item(i);
            var rect = RectOf(area);
            areas.Add(new Area(area, sheet, book, workbookName, sheetName, rect));
            cells += rect.CellCount;
        }

        var goTo = range;
        if (cells == 1)
        {
            object merged = r.MergeCells;
            if (merged is true)
            {
                goTo = r.MergeArea;
                label += " (merged " + LocalAddress(goTo) + ")";
                areas[0] = new Area(goTo, sheet, book, workbookName, sheetName, areas[0].Rect);
            }

            object hasFormula = ((dynamic)FirstCell(goTo)).HasFormula;
            var item = new PrecedentItem(PrecedentKind.Cell, label, workbookName, sheetName, LocalAddress(range), 1,
                CellValue(goTo), canExpand: hasFormula is true,
                hiddenNote: HiddenNote(range, areas[0].Rect, book, sheet, workbookName, sheetName));
            _targets[item] = new Target(goTo, areas, null);
            return item;
        }

        var rangeItem = new PrecedentItem(PrecedentKind.Range, label, workbookName, sheetName, LocalAddress(range), cells,
            TraceValueText.ForRange(CellValue(areas[0].Range), cells, _culture),
            hiddenNote: HiddenNote(range, areas[0].Rect, book, sheet, workbookName, sheetName));
        _targets[rangeItem] = new Target(range, areas, null);
        return rangeItem;
    }

    // A name's, table reference's or computed reference's target: its child item, and that child's range to go to.
    private Target Indirect(PrecedentItem child) =>
        new Target(_targets.TryGetValue(child, out var target) ? target.GetRange() : null, null, child);

    private PrecedentItem ErrorItem(string label, string reason) =>
        new PrecedentItem(PrecedentKind.Error, label, valueText: reason, canExpand: false);

    // The first cell's value as displayed (Range.Text), else its Value2.
    private string CellValue(object range)
    {
        dynamic cell = FirstCell(range);
        object? text = null;
        try
        {
            text = cell.Text;
        }
        catch (Exception)
        {
            // Text can fail on some cells (e.g. while Excel recalculates); Value2 still says something.
        }

        object value = cell.Value2;
        return TraceValueText.Display(text as string, value, _culture);
    }

    private string? HiddenNote(object range, CellRect rect, object book, object sheet, string workbookName, string sheetName)
    {
        var sheetNote = SheetNote(book, sheet, workbookName, sheetName);
        if (sheetNote.Length > 0)
        {
            return sheetNote;
        }

        bool? rows = false;
        bool? columns = false;
        try
        {
            dynamic r = range;
            if (rect.RowCount < CellRect.MaxRows)
            {
                object hidden = r.EntireRow.Hidden;
                rows = hidden is bool b ? b : (bool?)null;
            }

            if (rect.ColumnCount < CellRect.MaxColumns)
            {
                object hidden = r.EntireColumn.Hidden;
                columns = hidden is bool b ? b : (bool?)null;
            }
        }
        catch (Exception)
        {
            return null;
        }

        return TraceValueText.HiddenNote(false, false, rows, columns);
    }

    // "hidden workbook", "hidden sheet" or "" for a visible sheet; read once per sheet per trace.
    private string SheetNote(object book, object sheet, string workbookName, string sheetName)
    {
        var key = workbookName + "|" + sheetName;
        if (_sheetNotes.TryGetValue(key, out var note))
        {
            return note;
        }

        note = string.Empty;
        try
        {
            dynamic ws = sheet;
            object visible = ws.Visible;
            var workbookHidden = !WorkbookIsVisible(book);
            note = TraceValueText.HiddenNote(workbookHidden, Convert.ToInt32(visible, CultureInfo.InvariantCulture) != XlSheetVisible, false, false) ?? string.Empty;
        }
        catch (Exception)
        {
            // Unknown: no badge.
        }

        _sheetNotes[key] = note;
        return note;
    }

    /// <summary>True if any of the workbook's windows is visible (an add-in or hidden workbook has none). Never throws.</summary>
    internal static bool WorkbookIsVisible(object book)
    {
        try
        {
            dynamic wb = book;
            object windowsObject = wb.Windows;
            dynamic windows = windowsObject;
            int count = windows.Count;
            for (var i = 1; i <= count; i++)
            {
                object visible = windows.Item(i).Visible;
                if (visible is true)
                {
                    return true;
                }
            }

            return false;
        }
        catch (Exception)
        {
            return true;
        }
    }

    /// <summary>
    /// The open workbook named <paramref name="name"/>; if none and <paramref name="mayOpen"/>, the file is opened
    /// from <paramref name="folder"/> (as the formula wrote it) or else the folder of <paramref name="formulaBook"/>.
    /// Null with the reason in <paramref name="error"/> if it is not open and cannot be opened.
    /// </summary>
    private object? FindWorkbook(string name, string? folder, bool mayOpen, object formulaBook, out string? error)
    {
        error = null;
        dynamic app = ExcelDnaUtil.Application;
        try
        {
            object open = app.Workbooks.Item(name);
            return open;
        }
        catch (COMException)
        {
            // Not open under that name.
        }

        if (!mayOpen)
        {
            error = name + " is not open";
            return null;
        }

        string path;
        if (folder is not null)
        {
            path = folder + name;
        }
        else
        {
            dynamic wb = formulaBook;
            string formulaFolder = wb.Path;
            if (string.IsNullOrEmpty(formulaFolder))
            {
                error = name + " is not open, and the formula gives no folder to open it from";
                return null;
            }

            path = Path.Combine(formulaFolder, name);
        }

        return OpenWorkbook(path, out error);
    }

    /// <summary>
    /// Opens a closed workbook a precedent is in: read-only, without updating its links, with its macros disabled
    /// and no prompts (a password-protected workbook fails instead of asking). Excel's settings are put back and the
    /// window that was active is activated again, so Excel stays where the user was. Logs the open.
    /// </summary>
    private object? OpenWorkbook(string path, out string? error)
    {
        error = null;
        var stopwatch = Stopwatch.StartNew();
        var isUrl = path.StartsWith("http://", StringComparison.OrdinalIgnoreCase) || path.StartsWith("https://", StringComparison.OrdinalIgnoreCase);
        if (!isUrl && !File.Exists(path))
        {
            error = "the closed workbook was not found at " + path;
            DiagnosticsLog.Write("TraceWorkbookOpen", path, "ms=0", "not found");
            return null;
        }

        dynamic app = ExcelDnaUtil.Application;
        object? previousWindow = null;
        object? alerts = null;
        object? security = null;
        object? updating = null;
        try
        {
            previousWindow = app.ActiveWindow;
        }
        catch (Exception)
        {
            // No window is active (all hidden); nothing to go back to.
        }

        try
        {
            alerts = app.DisplayAlerts;
            security = app.AutomationSecurity;
            updating = app.ScreenUpdating;
            app.DisplayAlerts = false;
            app.AutomationSecurity = MsoAutomationSecurityForceDisable;
            app.ScreenUpdating = false;

            // Filename, UpdateLinks (0: never), ReadOnly, Format, Password, WriteResPassword,
            // IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify, Converter, AddToMru.
            object book = app.Workbooks.Open(path, 0, true, Type.Missing, NotAPassword, Type.Missing, true,
                Type.Missing, Type.Missing, false, false, Type.Missing, false);
            OpenedWorkbooks.Add(Path.GetFileName(path));
            DiagnosticsLog.Write("TraceWorkbookOpen", path, Ms(stopwatch), "opened read-only");
            return book;
        }
        catch (Exception ex)
        {
            error = "could not open " + path + ": " + ex.Message;
            DiagnosticsLog.Write("TraceWorkbookOpen", path, Ms(stopwatch), "failed: " + ex.Message);
            return null;
        }
        finally
        {
            // Only what was read (and so possibly changed) is put back.
            if (alerts is not null)
            {
                Restore(() => app.DisplayAlerts = alerts);
            }

            if (security is not null)
            {
                Restore(() => app.AutomationSecurity = security);
            }

            if (updating is not null)
            {
                Restore(() => app.ScreenUpdating = updating);
            }
            if (previousWindow is not null)
            {
                Restore(() => ((dynamic)previousWindow).Activate());
            }
        }
    }

    private static void Restore(Action action)
    {
        try
        {
            action();
        }
        catch (Exception ex)
        {
            DiagnosticsLog.Write("TraceWorkbookOpen", "could not restore a setting", ex.Message);
        }
    }

    // True if the sheet is Excel's active sheet. Never throws.
    private static bool IsActiveSheet(object sheet)
    {
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            object active = app.ActiveSheet;
            dynamic a = active;
            dynamic s = sheet;
            string activeName = a.Name;
            string sheetName = s.Name;
            string activeBook = a.Parent.Name;
            string sheetBook = s.Parent.Name;
            return string.Equals(activeName, sheetName, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(activeBook, sheetBook, StringComparison.OrdinalIgnoreCase);
        }
        catch (Exception)
        {
            return false;
        }
    }

    // A worksheet by name (Excel matches it ignoring case), or null.
    private static object? FindSheet(object book, string name)
    {
        try
        {
            dynamic wb = book;
            object sheet = wb.Worksheets.Item(name);
            return sheet;
        }
        catch (COMException)
        {
            return null;
        }
    }

    private static object? FirstWorksheet(object book)
    {
        dynamic wb = book;
        object sheets = wb.Worksheets;
        dynamic all = sheets;
        int count = all.Count;
        return count == 0 ? null : (object)all.Item(1);
    }

    // A workbook-level name. Names.Item can return a sheet-level name of the same name (Sheet1!Rate), so the
    // match is checked and, if it is not workbook-level, the collection is searched.
    private static object? WorkbookLevelName(object book, string name)
    {
        dynamic wb = book;
        object namesObject = wb.Names;
        dynamic names = namesObject;
        try
        {
            object found = names.Item(name);
            dynamic n = found;
            string foundName = n.Name;
            if (foundName.IndexOf('!') < 0)
            {
                return found;
            }
        }
        catch (COMException)
        {
            return null;
        }

        int count = names.Count;
        for (var i = 1; i <= count; i++)
        {
            object candidate = names.Item(i);
            dynamic n = candidate;
            string candidateName = n.Name;
            if (string.Equals(candidateName, name, StringComparison.OrdinalIgnoreCase))
            {
                return candidate;
            }
        }

        return null;
    }

    private static object? SheetLevelName(object book, string sheetName, string name)
    {
        var sheet = FindSheet(book, sheetName);
        if (sheet is null)
        {
            return null;
        }

        try
        {
            dynamic ws = sheet;
            object found = ws.Names.Item(name);
            return found;
        }
        catch (COMException)
        {
            // Not found by its short name: try the workbook's collection, where it is written Sheet!Name.
        }

        try
        {
            dynamic wb = book;
            var prefix = NeedsQuotes(sheetName) ? "'" + sheetName.Replace("'", "''") + "'" : sheetName;
            object found = wb.Names.Item(prefix + "!" + name);
            return found;
        }
        catch (COMException)
        {
            return null;
        }
    }

    // The table (ListObject) named name in any worksheet of the workbook, or null.
    private static object? FindTable(object book, string name)
    {
        dynamic wb = book;
        object sheetsObject = wb.Worksheets;
        dynamic sheets = sheetsObject;
        int sheetCount = sheets.Count;
        for (var s = 1; s <= sheetCount; s++)
        {
            object tablesObject = sheets.Item(s).ListObjects;
            dynamic tables = tablesObject;
            int count = tables.Count;
            for (var t = 1; t <= count; t++)
            {
                object table = tables.Item(t);
                dynamic lo = table;
                string tableName = lo.Name;
                if (string.Equals(tableName, name, StringComparison.OrdinalIgnoreCase))
                {
                    return table;
                }
            }
        }

        return null;
    }

    // The table holding the cell, or null.
    private static object? TableOf(object cell)
    {
        try
        {
            dynamic c = cell;
            object table = c.ListObject;
            return table;
        }
        catch (Exception)
        {
            return null;
        }
    }

    private static TableLayout Layout(object table)
    {
        dynamic lo = table;
        object rangeObject = lo.Range;
        var rect = RectOf(rangeObject);
        bool headers = lo.ShowHeaders;
        bool totals = lo.ShowTotals;
        object columnsObject = lo.ListColumns;
        dynamic listColumns = columnsObject;
        int count = listColumns.Count;
        var names = new List<string>(count);
        for (var i = 1; i <= count; i++)
        {
            string name = listColumns.Item(i).Name;
            names.Add(name);
        }

        return new TableLayout(rect, headers, totals, names);
    }

    private static bool IsRange(object? value)
    {
        if (value is null || !Marshal.IsComObject(value))
        {
            return false;
        }

        try
        {
            dynamic range = value;
            object address = range.Address;
            return address is string;
        }
        catch (Exception)
        {
            return false;
        }
    }

    private static object FirstCell(object range)
    {
        dynamic r = range;
        object cell = r.Cells.Item(1, 1);
        return cell;
    }

    /// <summary>The range's sheet-local address without <c>$</c>: <c>B5</c>, <c>A1:C10</c>, <c>A1:A5,C1:C5</c>.</summary>
    internal static string LocalAddress(object range) =>
        Convert.ToString(Get(range, "Address", false, false), CultureInfo.InvariantCulture) ?? string.Empty;

    /// <summary>
    /// The range's address for a label: <c>Sheet!B5</c>, or <c>[Book.xlsx]Sheet!B5</c> outside
    /// <paramref name="formulaWorkbook"/>.
    /// </summary>
    private static string DisplayAddress(object range, string formulaWorkbook)
    {
        dynamic r = range;
        object sheet = r.Worksheet;
        dynamic ws = sheet;
        string sheetName = ws.Name;
        string workbookName = ws.Parent.Name;
        var quoted = NeedsQuotes(sheetName) ? "'" + sheetName.Replace("'", "''") + "'" : sheetName;
        var local = LocalAddress(range);
        return string.Equals(workbookName, formulaWorkbook, StringComparison.OrdinalIgnoreCase)
            ? quoted + "!" + local
            : "[" + workbookName + "]" + quoted + "!" + local;
    }

    private static bool NeedsQuotes(string sheetName)
    {
        foreach (var c in sheetName)
        {
            if (!char.IsLetterOrDigit(c) && c != '_' && c != '.')
            {
                return true;
            }
        }

        return sheetName.Length > 0 && char.IsDigit(sheetName[0]);
    }

    private static CellRect RectOf(object range)
    {
        dynamic r = range;
        object row = r.Row;
        object column = r.Column;
        object rows = r.Rows.Count;
        object columns = r.Columns.Count;
        return new CellRect(
            Convert.ToInt32(row, CultureInfo.InvariantCulture),
            Convert.ToInt32(column, CultureInfo.InvariantCulture),
            Convert.ToInt32(rows, CultureInfo.InvariantCulture),
            Convert.ToInt32(columns, CultureInfo.InvariantCulture));
    }

    // One element of a block's Value2 or Formula: a 1-based 2-D array, or the value itself for one cell.
    private static object? Element(object? values, int row, int column) =>
        values is object[,] array
            ? array[array.GetLowerBound(0) + row, array.GetLowerBound(1) + column]
            : values;

    // A parameterized property read (Range.Address(RowAbsolute, ColumnAbsolute)).
    private static object Get(object target, string name, params object[] args) =>
        target.GetType().InvokeMember(name, BindingFlags.GetProperty, null, target, args, CultureInfo.InvariantCulture);

    private static string Ms(Stopwatch stopwatch) =>
        "ms=" + stopwatch.Elapsed.TotalMilliseconds.ToString("0.0", CultureInfo.InvariantCulture);

    /// <summary>One rectangle of a range, with its sheet and workbook.</summary>
    private sealed class Area
    {
        public Area(object range, object sheet, object book, string workbookName, string sheetName, CellRect rect)
        {
            Range = range;
            Sheet = sheet;
            Book = book;
            WorkbookName = workbookName;
            SheetName = sheetName;
            Rect = rect;
        }

        public object Range { get; }

        public object Sheet { get; }

        public object Book { get; }

        public string WorkbookName { get; }

        public string SheetName { get; }

        public CellRect Rect { get; }
    }

    /// <summary>
    /// What an item stands for: a range (with its areas, for paging and for a cell's formula), or, for a name,
    /// table reference or computed reference, the child item it resolves to. A paged cell's range is made only when
    /// it is first needed.
    /// </summary>
    private sealed class Target
    {
        private readonly Area? _cellArea;
        private readonly string? _cellAddress;
        private object? _range;
        private List<Area>? _areas;

        public Target(object? range, List<Area>? areas, PrecedentItem? child)
        {
            _range = range;
            _areas = areas;
            Child = child;
        }

        private Target(Area area, string cellAddress)
        {
            _cellArea = area;
            _cellAddress = cellAddress;
        }

        public PrecedentItem? Child { get; }

        public List<Area>? Areas
        {
            get
            {
                if (_areas is null && _cellArea is not null && GetRange() is object range)
                {
                    _areas = new List<Area>
                    {
                        new Area(range, _cellArea.Sheet, _cellArea.Book, _cellArea.WorkbookName, _cellArea.SheetName, RectOf(range)),
                    };
                }

                return _areas;
            }
        }

        public static Target ForRange(object range, Area area) => new Target(range, new List<Area> { area }, null);

        public static Target ForCell(Area area, string address) => new Target(area, address);

        public object? GetRange()
        {
            if (_range is null && _cellArea is not null)
            {
                dynamic ws = _cellArea.Sheet;
                _range = ws.Range(_cellAddress);
            }

            return _range;
        }
    }
}
