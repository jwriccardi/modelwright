using ExcelModelingToolkit.Core.Formatting;

namespace ExcelModelingToolkit.Core.Undo;

/// <summary>Reads one format property of a rectangle of cells on one sheet (the add-in implements it over COM).</summary>
public interface IFormatReader
{
    /// <summary>
    /// The value every cell of <paramref name="range"/> holds, or <see cref="CycleValue.Unknown"/> if they differ
    /// (mixed). A single cell can also be mixed, e.g. rich text in several font colors.
    /// </summary>
    CycleValue ReadUniform(CellRect range);
}
