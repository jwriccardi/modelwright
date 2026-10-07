using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelModelingToolkit.Core.Trace;

namespace ExcelModelingToolkit.Core.Tests.Trace;

/// <summary>
/// A workbook's precedents as a lookup from an item's label to its children; a range's cells are generated
/// (<c>R1</c>, <c>R2</c>, ...) unless given. Records every call, to check that loading is lazy.
/// </summary>
internal sealed class FakePrecedentProvider : IPrecedentProvider
{
    private readonly Dictionary<string, PrecedentItem[]> _precedents = new Dictionary<string, PrecedentItem[]>(StringComparer.Ordinal);
    private readonly Dictionary<string, PrecedentItem[]> _rangeCells = new Dictionary<string, PrecedentItem[]>(StringComparer.Ordinal);

    public List<string> Calls { get; } = new List<string>();

    public Exception? ThrowNext { get; set; }

    public static PrecedentItem Cell(string address, string sheet = "Calc", bool canExpand = true) =>
        new PrecedentItem(PrecedentKind.Cell, sheet + "!" + address, "Model.xlsx", sheet, address, canExpand: canExpand);

    public static PrecedentItem Range(string address, long cells, string sheet = "Calc") =>
        new PrecedentItem(PrecedentKind.Range, sheet + "!" + address, "Model.xlsx", sheet, address, cells);

    public FakePrecedentProvider Has(PrecedentItem item, params PrecedentItem[] precedents)
    {
        _precedents[item.Label] = precedents;
        return this;
    }

    public FakePrecedentProvider RangeHas(PrecedentItem range, params PrecedentItem[] cells)
    {
        _rangeCells[range.Label] = cells;
        return this;
    }

    public IReadOnlyList<PrecedentItem> GetPrecedents(PrecedentItem item)
    {
        Calls.Add("precedents " + item.Label);
        ThrowIfAsked();
        return _precedents.TryGetValue(item.Label, out var found) ? found : new PrecedentItem[0];
    }

    public IReadOnlyList<PrecedentItem> GetRangeCells(PrecedentItem range, long start, int count)
    {
        Calls.Add(string.Format(CultureInfo.InvariantCulture, "cells {0} {1} {2}", range.Label, start, count));
        ThrowIfAsked();
        if (_rangeCells.TryGetValue(range.Label, out var given))
        {
            return given.Skip((int)start).Take(count).ToList();
        }

        var end = Math.Min(range.CellCount, start + count);
        var cells = new List<PrecedentItem>();
        for (var index = start; index < end; index++)
        {
            cells.Add(Cell("R" + (index + 1).ToString(CultureInfo.InvariantCulture), range.Sheet!));
        }

        return cells;
    }

    private void ThrowIfAsked()
    {
        if (ThrowNext is Exception exception)
        {
            ThrowNext = null;
            throw exception;
        }
    }
}
