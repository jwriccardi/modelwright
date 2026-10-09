namespace Modelwright.Core.Trace;

/// <summary>The kind of a <see cref="FormulaNode"/>.</summary>
public enum FormulaNodeKind
{
    /// <summary>A parenthesized expression, <c>(B2+C2/D2)</c>; its one child is the inner expression.</summary>
    Group,

    /// <summary>A function call, <c>SUM(...)</c>; its children are the arguments, in order.</summary>
    Function,

    /// <summary>A reference (<see cref="FormulaNode.Reference"/>); a leaf.</summary>
    Reference,

    /// <summary>
    /// An operator applied to its operands (<see cref="FormulaNode.Operator"/>): arithmetic, comparison, <c>&amp;</c>,
    /// <c>%</c>, <c>@</c>, and the reference operators range (<c>:</c>), intersection (a space) and union (<c>,</c>).
    /// </summary>
    Operator,

    /// <summary>A number, string, boolean or error constant; a leaf.</summary>
    Constant,

    /// <summary>An array constant, <c>{1,2;3,4}</c>; a leaf (it can hold no references).</summary>
    ArrayConstant,

    /// <summary>An omitted function argument, as in <c>IF(A1,,0)</c>; a leaf with no text.</summary>
    MissingArgument,

    /// <summary>A construct with no dedicated kind (for example a DDE link); its children are its parts.</summary>
    Other,
}
