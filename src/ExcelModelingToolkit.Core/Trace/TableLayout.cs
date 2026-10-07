using System;
using System.Collections.Generic;
using System.Linq;
using ExcelModelingToolkit.Core.Undo;

namespace ExcelModelingToolkit.Core.Trace;

/// <summary>
/// The shape of an Excel table (a <c>ListObject</c>): where it is, whether it shows its header and totals rows, and
/// its column names. <see cref="Resolve"/> finds the cells a structured reference to it means.
/// </summary>
public sealed class TableLayout
{
    /// <summary>Creates a layout.</summary>
    /// <param name="range">The whole table (<c>ListObject.Range</c>): header, data and totals rows.</param>
    /// <param name="hasHeaders">True if the header row is shown.</param>
    /// <param name="hasTotals">True if the totals row is shown.</param>
    /// <param name="columns">The column names, left to right: one per column of <paramref name="range"/>.</param>
    /// <exception cref="ArgumentNullException"><paramref name="columns"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// The column count differs from the range's, or the range has fewer rows than its header and totals.
    /// </exception>
    public TableLayout(CellRect range, bool hasHeaders, bool hasTotals, IReadOnlyList<string> columns)
    {
        if (columns is null)
        {
            throw new ArgumentNullException(nameof(columns));
        }

        if (columns.Count != range.ColumnCount)
        {
            throw new ArgumentException("There must be one column name per column of the table.", nameof(columns));
        }

        if (range.RowCount < (hasHeaders ? 1 : 0) + (hasTotals ? 1 : 0))
        {
            throw new ArgumentException("The table has fewer rows than its header and totals rows.", nameof(range));
        }

        Range = range;
        HasHeaders = hasHeaders;
        HasTotals = hasTotals;
        Columns = columns.ToArray();
    }

    /// <summary>The whole table.</summary>
    public CellRect Range { get; }

    /// <summary>True if the header row is shown.</summary>
    public bool HasHeaders { get; }

    /// <summary>True if the totals row is shown.</summary>
    public bool HasTotals { get; }

    /// <summary>The column names.</summary>
    public IReadOnlyList<string> Columns { get; }

    /// <summary>The first data row.</summary>
    public int FirstDataRow => Range.Row + (HasHeaders ? 1 : 0);

    /// <summary>The last data row (below <see cref="FirstDataRow"/> if the table has no data rows).</summary>
    public int LastDataRow => Range.LastRow - (HasTotals ? 1 : 0);

    /// <summary>
    /// The cells a structured reference to this table means, or null with the reason in <paramref name="error"/>.
    /// </summary>
    /// <param name="specifiers">
    /// The reference's item specifiers (<see cref="FormulaReference.TableSpecifiers"/>): <c>#All</c>,
    /// <c>#Data</c>, <c>#Headers</c>, <c>#Totals</c>, <c>#This Row</c> or <c>@</c>, in any case; none means
    /// <c>#Data</c>. <c>#Headers</c> with <c>#Data</c>, and <c>#Data</c> with <c>#Totals</c>, combine.
    /// </param>
    /// <param name="columns">
    /// The reference's columns (<see cref="FormulaReference.TableColumns"/>), matched ignoring case: none for every
    /// column, one, or two for the span between them (<c>[[Q1]:[Q4]]</c>).
    /// </param>
    /// <param name="formulaRow">
    /// The row of the formula's cell, for <c>@</c> and <c>#This Row</c>; null if unknown.
    /// </param>
    /// <param name="error">Why there are no cells, or null.</param>
    /// <exception cref="ArgumentNullException"><paramref name="specifiers"/> or <paramref name="columns"/> is null.</exception>
    public CellRect? Resolve(IReadOnlyList<string> specifiers, IReadOnlyList<string> columns, int? formulaRow, out string? error)
    {
        if (specifiers is null)
        {
            throw new ArgumentNullException(nameof(specifiers));
        }

        if (columns is null)
        {
            throw new ArgumentNullException(nameof(columns));
        }

        error = null;
        if (!ResolveColumns(columns, out var firstColumn, out var lastColumn, ref error) ||
            !ResolveRows(specifiers, formulaRow, out var firstRow, out var lastRow, ref error))
        {
            return null;
        }

        return new CellRect(firstRow, firstColumn, lastRow - firstRow + 1, lastColumn - firstColumn + 1);
    }

    private bool ResolveColumns(IReadOnlyList<string> columns, out int first, out int last, ref string? error)
    {
        first = Range.Column;
        last = Range.LastColumn;
        if (columns.Count == 0)
        {
            return true;
        }

        if (columns.Count > 2)
        {
            error = "a structured reference can name at most two columns (a span)";
            return false;
        }

        var indexes = new int[columns.Count];
        for (var i = 0; i < columns.Count; i++)
        {
            indexes[i] = IndexOfColumn(columns[i]);
            if (indexes[i] < 0)
            {
                error = "the table has no column '" + columns[i] + "'";
                return false;
            }
        }

        first = Range.Column + indexes.Min();
        last = Range.Column + indexes.Max();
        return true;
    }

    private bool ResolveRows(IReadOnlyList<string> specifiers, int? formulaRow, out int first, out int last, ref string? error)
    {
        first = last = 0;
        bool headers = false, data = false, totals = false, all = false, thisRow = false;
        foreach (var specifier in specifiers)
        {
            switch (specifier.Trim().ToUpperInvariant())
            {
                case "#ALL":
                    all = true;
                    break;
                case "#DATA":
                    data = true;
                    break;
                case "#HEADERS":
                    headers = true;
                    break;
                case "#TOTALS":
                    totals = true;
                    break;
                case "#THIS ROW":
                case "@":
                    thisRow = true;
                    break;
                default:
                    error = "'" + specifier + "' is not a table item specifier";
                    return false;
            }
        }

        if (all)
        {
            first = Range.Row;
            last = Range.LastRow;
            return true;
        }

        if (thisRow)
        {
            if (headers || data || totals)
            {
                error = "this row cannot be combined with other item specifiers";
                return false;
            }

            if (formulaRow is not int row || row < FirstDataRow || row > LastDataRow)
            {
                error = "this row is outside the table's data rows";
                return false;
            }

            first = last = row;
            return true;
        }

        if (!headers && !totals)
        {
            data = true;
        }

        if (headers && totals && !data)
        {
            error = "the header and totals rows are not next to each other";
            return false;
        }

        if ((headers && !HasHeaders) || (totals && !HasTotals))
        {
            error = headers && !HasHeaders ? "the table's header row is hidden" : "the table's totals row is hidden";
            return false;
        }

        if (data && LastDataRow < FirstDataRow)
        {
            error = "the table has no data rows";
            return false;
        }

        // The rows run from the topmost part named to the bottommost: header, data, totals.
        first = headers ? Range.Row : data ? FirstDataRow : Range.LastRow;
        last = totals ? Range.LastRow : data ? LastDataRow : Range.Row;
        return true;
    }

    private int IndexOfColumn(string name)
    {
        for (var i = 0; i < Columns.Count; i++)
        {
            if (string.Equals(Columns[i], name, StringComparison.OrdinalIgnoreCase))
            {
                return i;
            }
        }

        return -1;
    }
}
