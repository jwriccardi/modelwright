using System;

namespace Modelwright.Core.Trace;

/// <summary>Where a formula lives: the workbook and sheet that unqualified references in it point to.</summary>
public sealed class FormulaContext
{
    /// <summary>Creates a context.</summary>
    /// <param name="workbookName">The workbook's file name as Excel shows it in references (<c>Model.xlsx</c>).</param>
    /// <param name="sheetName">The worksheet holding the formula cell.</param>
    /// <exception cref="ArgumentException">A name is null or empty.</exception>
    public FormulaContext(string workbookName, string sheetName)
    {
        if (string.IsNullOrEmpty(workbookName))
        {
            throw new ArgumentException("A workbook name is required.", nameof(workbookName));
        }

        if (string.IsNullOrEmpty(sheetName))
        {
            throw new ArgumentException("A sheet name is required.", nameof(sheetName));
        }

        WorkbookName = workbookName;
        SheetName = sheetName;
    }

    /// <summary>The formula's workbook file name.</summary>
    public string WorkbookName { get; }

    /// <summary>The formula's worksheet name.</summary>
    public string SheetName { get; }
}
