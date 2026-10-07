using System.Collections.Generic;
using ExcelModelingToolkit.Core.Undo;

namespace ExcelModelingToolkit.Core.Trace;

/// <summary>
/// One reference in a formula, as written: its kind, where it sits in the formula text (for coloring), and the
/// target it names, relative to the formula's <see cref="FormulaContext"/>.
/// </summary>
public sealed class FormulaReference
{
    private static readonly string[] None = new string[0];

    // The whole formula, as in FormulaNode: a merged bounding range (A1:A2:...:A1500) spans most of it.
    private readonly string _formula;

    internal FormulaReference(FormulaReferenceKind kind, string formula, int start, int length)
    {
        Kind = kind;
        _formula = formula;
        Start = start;
        Length = length;
    }

    /// <summary>What the reference points to.</summary>
    public FormulaReferenceKind Kind { get; internal set; }

    /// <summary>
    /// The reference exactly as written in the formula, including any sheet or workbook prefix. Read from the
    /// formula on each access, not stored.
    /// </summary>
    public string Text => _formula.Substring(Start, Length);

    /// <summary>The 0-based position of <see cref="Text"/> in the formula string (which starts with <c>=</c>).</summary>
    public int Start { get; }

    /// <summary>The length of <see cref="Text"/>.</summary>
    public int Length { get; internal set; }

    /// <summary>
    /// The external workbook's file name (<c>Book.xlsx</c>), or null for the formula's own workbook. Whether that
    /// workbook is open is not known here.
    /// </summary>
    public string? WorkbookName { get; internal set; }

    /// <summary>
    /// The folder or URL written before an external workbook's name (<c>C:\dir\</c>), with its trailing separator;
    /// null when none is written (Excel writes the path only when the workbook is closed).
    /// </summary>
    public string? WorkbookPath { get; internal set; }

    /// <summary>True if the reference points into another workbook.</summary>
    public bool IsExternal => WorkbookName is not null;

    /// <summary>
    /// The sheet named by the reference (the first sheet of a 3-D reference), or null if it names none or names the
    /// formula's own sheet in the formula's own workbook.
    /// </summary>
    public string? Sheet { get; internal set; }

    /// <summary>The last sheet of a 3-D reference (<c>Sheet1:Sheet3!A1</c>), else null.</summary>
    public string? LastSheet { get; internal set; }

    /// <summary>True for a 3-D reference spanning sheets <see cref="Sheet"/> to <see cref="LastSheet"/>.</summary>
    public bool Is3D => LastSheet is not null;

    /// <summary>
    /// The normalized sheet-local A1 address, without <c>$</c> and in upper case, corners ordered: <c>A1</c>,
    /// <c>A1:B5</c>, <c>A:C</c>, <c>3:5</c>. Null for names, structured references and <c>#REF!</c>.
    /// </summary>
    public string? Address { get; internal set; }

    /// <summary>The cells of <see cref="Address"/>, or null when it has none.</summary>
    public CellRect? Area { get; internal set; }

    /// <summary>
    /// The defined or local name, or the table name of a structured reference (null for an unqualified one such as
    /// <c>[@Col]</c>).
    /// </summary>
    public string? Name { get; internal set; }

    /// <summary>The column names of a structured reference (<c>Revenue</c>, <c>Cost</c>), unescaped.</summary>
    public IReadOnlyList<string> TableColumns { get; internal set; } = None;

    /// <summary>The item specifiers of a structured reference (<c>#This Row</c>, <c>#Totals</c>, <c>@</c>).</summary>
    public IReadOnlyList<string> TableSpecifiers { get; internal set; } = None;

    /// <summary>
    /// True for a structured reference written without a table name (<c>[@Col]</c>, <c>[Col]</c>): it refers to the
    /// table containing the formula cell, which only the caller can find.
    /// </summary>
    public bool NeedsTableContext => Kind == FormulaReferenceKind.StructuredReference && Name is null;

    /// <summary>True for a spilled-range reference (<c>A1#</c>); <see cref="Text"/> includes the <c>#</c>.</summary>
    public bool IsSpill { get; internal set; }

    /// <summary>
    /// The innermost function enclosing this reference that can return a reference itself (INDEX, OFFSET,
    /// INDIRECT or CHOOSE; see <see cref="ParsedFormula.DynamicReferences"/>), or null. For example, in
    /// <c>INDEX(A1:C9,2,3)</c> the precedent is one cell of <c>A1:C9</c>, not the whole range.
    /// </summary>
    public string? EnclosingReferenceFunction { get; internal set; }

    /// <summary>The reference as written.</summary>
    public override string ToString() => Text;
}
