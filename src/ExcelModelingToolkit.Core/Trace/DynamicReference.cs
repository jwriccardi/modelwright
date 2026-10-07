namespace ExcelModelingToolkit.Core.Trace;

/// <summary>
/// A call in a formula to a function that can return a reference computed at run time (INDEX, OFFSET, INDIRECT,
/// CHOOSE). Parsing cannot say which cells it points to; the caller evaluates it (docs/PLAN.md section 4.5).
/// </summary>
public sealed class DynamicReference
{
    internal DynamicReference(string functionName, string text, int start)
    {
        FunctionName = functionName;
        Text = text;
        Start = start;
    }

    /// <summary>The function's name in upper case (<c>INDIRECT</c>).</summary>
    public string FunctionName { get; }

    /// <summary>The call's source text, from the name to the closing parenthesis.</summary>
    public string Text { get; }

    /// <summary>The 0-based position of <see cref="Text"/> in the formula string.</summary>
    public int Start { get; }

    /// <summary>The length of <see cref="Text"/>.</summary>
    public int Length => Text.Length;

    /// <summary>The call as written.</summary>
    public override string ToString() => Text;
}
