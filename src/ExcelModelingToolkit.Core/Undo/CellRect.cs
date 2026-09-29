using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcelModelingToolkit.Core.Undo;

/// <summary>
/// A rectangle of cells on one worksheet: 1-based first row and column, and a size of at least one cell. Stays
/// within Excel's sheet limits (<see cref="MaxRows"/> by <see cref="MaxColumns"/>).
/// </summary>
/// <remarks><c>default(CellRect)</c> is not a valid rectangle; use the constructor.</remarks>
public readonly struct CellRect : IEquatable<CellRect>
{
    /// <summary>Rows in a worksheet (Excel 2007 and later).</summary>
    public const int MaxRows = 1048576;

    /// <summary>Columns in a worksheet (Excel 2007 and later), <c>A</c> to <c>XFD</c>.</summary>
    public const int MaxColumns = 16384;

    /// <summary>Creates a rectangle.</summary>
    /// <exception cref="ArgumentOutOfRangeException">
    /// A position or size is below 1, or the rectangle extends past the last row or column of a sheet.
    /// </exception>
    public CellRect(int row, int column, int rowCount, int columnCount)
    {
        Check(row, rowCount, MaxRows, nameof(row), nameof(rowCount));
        Check(column, columnCount, MaxColumns, nameof(column), nameof(columnCount));
        Row = row;
        Column = column;
        RowCount = rowCount;
        ColumnCount = columnCount;
    }

    /// <summary>First row, 1-based.</summary>
    public int Row { get; }

    /// <summary>First column, 1-based (<c>A</c> is 1).</summary>
    public int Column { get; }

    /// <summary>Number of rows (at least 1).</summary>
    public int RowCount { get; }

    /// <summary>Number of columns (at least 1).</summary>
    public int ColumnCount { get; }

    /// <summary>Last row, 1-based.</summary>
    public int LastRow => Row + RowCount - 1;

    /// <summary>Last column, 1-based.</summary>
    public int LastColumn => Column + ColumnCount - 1;

    /// <summary>Number of cells (up to 17,179,869,184 for a whole sheet).</summary>
    public long CellCount => (long)RowCount * ColumnCount;

    /// <summary>True for a single cell.</summary>
    public bool IsSingleCell => RowCount == 1 && ColumnCount == 1;

    /// <summary>
    /// The sheet-local absolute A1 address: <c>$B$3</c> for one cell, else <c>$B$3:$D$10</c>. Whole columns and
    /// rows are spelled out (<c>$A$1:$A$1048576</c>), which <c>Range</c> accepts.
    /// </summary>
    public string Address
    {
        get
        {
            var first = CellAddress(Row, Column);
            return IsSingleCell ? first : first + ":" + CellAddress(LastRow, LastColumn);
        }
    }

    /// <summary>The column letters for a 1-based column number: 1 is <c>A</c>, 27 is <c>AA</c>, 16384 is <c>XFD</c>.</summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="column"/> is outside 1 to <see cref="MaxColumns"/>.</exception>
    public static string ColumnName(int column)
    {
        if (column < 1 || column > MaxColumns)
        {
            throw new ArgumentOutOfRangeException(nameof(column), column, $"A column must be between 1 and {MaxColumns}.");
        }

        var letters = new StringBuilder(3);
        while (column > 0)
        {
            var remainder = (column - 1) % 26;
            letters.Insert(0, (char)('A' + remainder));
            column = (column - 1) / 26;
        }

        return letters.ToString();
    }

    /// <summary>The cells in both rectangles, or null if they do not overlap.</summary>
    public CellRect? Intersect(CellRect other)
    {
        var row = Math.Max(Row, other.Row);
        var column = Math.Max(Column, other.Column);
        var lastRow = Math.Min(LastRow, other.LastRow);
        var lastColumn = Math.Min(LastColumn, other.LastColumn);
        return row > lastRow || column > lastColumn
            ? (CellRect?)null
            : new CellRect(row, column, lastRow - row + 1, lastColumn - column + 1);
    }

    /// <summary>
    /// The cells of this rectangle outside <paramref name="other"/>, as up to four non-overlapping bands: above
    /// and below it (full width), then left and right of it (only the rows they share). Just this rectangle if
    /// they do not overlap; empty if <paramref name="other"/> covers it.
    /// </summary>
    public IReadOnlyList<CellRect> Subtract(CellRect other)
    {
        if (Intersect(other) is not CellRect inner)
        {
            return new[] { this };
        }

        var bands = new List<CellRect>(4);
        if (inner.Row > Row)
        {
            bands.Add(new CellRect(Row, Column, inner.Row - Row, ColumnCount));
        }

        if (inner.LastRow < LastRow)
        {
            bands.Add(new CellRect(inner.LastRow + 1, Column, LastRow - inner.LastRow, ColumnCount));
        }

        if (inner.Column > Column)
        {
            bands.Add(new CellRect(inner.Row, Column, inner.RowCount, inner.Column - Column));
        }

        if (inner.LastColumn < LastColumn)
        {
            bands.Add(new CellRect(inner.Row, inner.LastColumn + 1, inner.RowCount, LastColumn - inner.LastColumn));
        }

        return bands;
    }

    /// <inheritdoc />
    public bool Equals(CellRect other) =>
        Row == other.Row && Column == other.Column && RowCount == other.RowCount && ColumnCount == other.ColumnCount;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is CellRect other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode()
    {
        unchecked
        {
            var hash = Row;
            hash = (hash * 397) ^ Column;
            hash = (hash * 397) ^ RowCount;
            return (hash * 397) ^ ColumnCount;
        }
    }

    /// <summary>Value equality.</summary>
    public static bool operator ==(CellRect left, CellRect right) => left.Equals(right);

    /// <summary>Value inequality.</summary>
    public static bool operator !=(CellRect left, CellRect right) => !left.Equals(right);

    /// <summary>The <see cref="Address"/>.</summary>
    public override string ToString() => Address;

    private static string CellAddress(int row, int column) =>
        "$" + ColumnName(column) + "$" + row.ToString(CultureInfo.InvariantCulture);

    private static void Check(int start, int count, int max, string startName, string countName)
    {
        if (start < 1 || start > max)
        {
            throw new ArgumentOutOfRangeException(startName, start, $"Must be between 1 and {max}.");
        }

        if (count < 1 || count > max - start + 1)
        {
            throw new ArgumentOutOfRangeException(countName, count, $"Must be at least 1 and end by {max}.");
        }
    }
}
