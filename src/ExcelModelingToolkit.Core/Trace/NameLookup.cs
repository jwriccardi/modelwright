using System;
using System.Collections.Generic;

namespace ExcelModelingToolkit.Core.Trace;

/// <summary>
/// One place to look for a defined name a formula refers to: a workbook, and the sheet whose scope holds it (null
/// for a workbook-level name). <see cref="Candidates"/> lists them in the order the add-in tries them.
/// </summary>
public sealed class NameLookup
{
    private const string ReservedPrefix = "_xlnm.";

    private NameLookup(string? workbookName, string? sheetName, string name, bool mayOpenWorkbook)
    {
        WorkbookName = workbookName;
        SheetName = sheetName;
        Name = name;
        MayOpenWorkbook = mayOpenWorkbook;
    }

    /// <summary>The workbook to look in, or null for the formula's own workbook.</summary>
    public string? WorkbookName { get; }

    /// <summary>The sheet whose scope holds the name, or null for a workbook-level name.</summary>
    public string? SheetName { get; }

    /// <summary>
    /// The name as Excel's <c>Names</c> collection holds it: without the <c>_xlnm.</c> prefix of a built-in name
    /// (<c>_xlnm.Print_Area</c> is <c>Print_Area</c>).
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// True if <see cref="WorkbookName"/> was written as a workbook in the formula, so a closed workbook of that
    /// name may be opened to look in it; false if only an open workbook can match (the <c>Book2</c> of
    /// <c>Book2!Rate</c>, which may instead be a sheet).
    /// </summary>
    public bool MayOpenWorkbook { get; }

    /// <summary>
    /// Where to look for the name <paramref name="reference"/> refers to, in order; the first that has it wins:
    /// <list type="bullet">
    /// <item><c>Rate</c>: the formula's sheet's scope, then the formula's workbook.</item>
    /// <item><c>Inputs!Rate</c>: the scope of sheet <c>Inputs</c> in the formula's workbook; then, since an unsaved
    /// workbook has no extension to tell it from a sheet (docs/PLAN.md, notes for 4b), the workbook level of an open
    /// workbook named <c>Inputs</c>.</item>
    /// <item><c>Book.xlsx!Rate</c>: the workbook level of <c>Book.xlsx</c>, also when it is the formula's own
    /// workbook (a sheet-level <c>Rate</c> on the formula's sheet does not hide it).</item>
    /// <item><c>[Book.xlsx]Inputs!Rate</c>: the scope of sheet <c>Inputs</c> in <c>Book.xlsx</c>.</item>
    /// </list>
    /// </summary>
    /// <param name="reference">A <see cref="FormulaReferenceKind.Name"/> reference.</param>
    /// <param name="context">The workbook and sheet of the formula it is in.</param>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="reference"/> is not a defined name.</exception>
    public static IReadOnlyList<NameLookup> Candidates(FormulaReference reference, FormulaContext context)
    {
        if (reference is null)
        {
            throw new ArgumentNullException(nameof(reference));
        }

        if (context is null)
        {
            throw new ArgumentNullException(nameof(context));
        }

        if (reference.Kind != FormulaReferenceKind.Name || reference.Name is null)
        {
            throw new ArgumentException("The reference is not a defined name.", nameof(reference));
        }

        var name = reference.Name.StartsWith(ReservedPrefix, StringComparison.OrdinalIgnoreCase)
            ? reference.Name.Substring(ReservedPrefix.Length)
            : reference.Name;
        if (reference.WorkbookName is not null)
        {
            return new[] { new NameLookup(reference.WorkbookName, reference.Sheet, name, mayOpenWorkbook: true) };
        }

        if (reference.IsWorkbookLevelQualified)
        {
            // Model.xlsx!Rate in Model.xlsx itself: the workbook-level name, even where a sheet has its own Rate.
            return new[] { new NameLookup(null, null, name, mayOpenWorkbook: false) };
        }

        if (reference.Sheet is not null)
        {
            return new[]
            {
                new NameLookup(null, reference.Sheet, name, mayOpenWorkbook: false),
                new NameLookup(reference.Sheet, null, name, mayOpenWorkbook: false),
            };
        }

        return new[]
        {
            new NameLookup(null, context.SheetName, name, mayOpenWorkbook: false),
            new NameLookup(null, null, name, mayOpenWorkbook: false),
        };
    }

    /// <summary>For the log: <c>[Book]Sheet!Name</c> with the parts that are set.</summary>
    public override string ToString() =>
        (WorkbookName is null ? string.Empty : "[" + WorkbookName + "]") + (SheetName is null ? string.Empty : SheetName + "!") + Name;
}
