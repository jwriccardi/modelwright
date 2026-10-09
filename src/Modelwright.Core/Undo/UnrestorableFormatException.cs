using System;

namespace Modelwright.Core.Undo;

/// <summary>
/// Thrown by an <see cref="IFormatReader"/> when a cell holds a format that undo could not put back exactly (for
/// example, a pattern or gradient fill). <see cref="SnapshotPlanner"/> turns it into an unavailable plan whose
/// reason is <see cref="Exception.Message"/>.
/// </summary>
public sealed class UnrestorableFormatException : Exception
{
    /// <summary>Creates the exception; <paramref name="reason"/> is shown to the user, e.g. <c>pattern or gradient fill</c>.</summary>
    public UnrestorableFormatException(string reason)
        : base(reason)
    {
    }
}
