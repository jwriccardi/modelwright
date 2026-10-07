namespace ExcelModelingToolkit.Core.Trace;

/// <summary>What a <see cref="FormulaReference"/> points to.</summary>
public enum FormulaReferenceKind
{
    /// <summary>One cell: <c>A1</c>, <c>Sheet2!$B$5</c>, <c>[Book.xlsx]Data!C3</c>.</summary>
    Cell,

    /// <summary>A rectangle of cells: <c>A1:B5</c>.</summary>
    Range,

    /// <summary>Whole columns: <c>A:A</c>, <c>$B:$D</c>.</summary>
    WholeColumn,

    /// <summary>Whole rows: <c>1:1</c>, <c>$3:$5</c>.</summary>
    WholeRow,

    /// <summary>A defined name: workbook-level (<c>Revenue</c>) or sheet-scoped (<c>Sheet1!Revenue</c>).</summary>
    Name,

    /// <summary>
    /// A name declared by <c>LET</c> or <c>LAMBDA</c> inside the formula (the declaration or a use of it). It is not
    /// a defined name and has no cells to navigate to.
    /// </summary>
    LocalName,

    /// <summary>A structured (table) reference: <c>Table1[Col]</c>, <c>Table1[[#This Row],[Col]]</c>, <c>[@Col]</c>.</summary>
    StructuredReference,

    /// <summary>A deleted reference: <c>#REF!</c> or <c>Sheet1!#REF!</c>.</summary>
    RefError,
}
