using ExcelModelingToolkit.Core.Formatting;
using ExcelModelingToolkit.Core.Undo;

namespace ExcelModelingToolkit.AddIn;

/// <summary>Reads one format property of rectangles on one worksheet over COM (<see cref="CellFormats.Read"/>).</summary>
internal sealed class SheetFormatReader : IFormatReader
{
    private readonly object _worksheet;
    private readonly CycleKind _kind;

    /// <summary>Creates a reader for <paramref name="kind"/> on <paramref name="worksheet"/> (a COM <c>Worksheet</c>).</summary>
    public SheetFormatReader(object worksheet, CycleKind kind)
    {
        _worksheet = worksheet;
        _kind = kind;
    }

    /// <inheritdoc />
    public CycleValue ReadUniform(CellRect range)
    {
        dynamic sheet = _worksheet;
        object cells = sheet.Range(range.Address);
        return CellFormats.Read(cells, _kind);
    }
}
