namespace ExcelModelingToolkit.Core.Trace;

/// <summary>Where an operator sits relative to its operands.</summary>
public enum OperatorFixity
{
    /// <summary>Not an operator node.</summary>
    None,

    /// <summary>Before its operand: <c>-A1</c>, <c>@A1:A9</c>.</summary>
    Prefix,

    /// <summary>Between its operands: <c>A1+B1</c>, <c>A1:B2</c>.</summary>
    Infix,

    /// <summary>After its operand: <c>A1%</c>.</summary>
    Postfix,
}
