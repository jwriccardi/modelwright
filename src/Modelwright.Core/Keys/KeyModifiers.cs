using System;

namespace Modelwright.Core.Keys;

/// <summary>Modifier keys of a <see cref="KeyChord"/>.</summary>
[Flags]
public enum KeyModifiers
{
    /// <summary>No modifier.</summary>
    None = 0,

    /// <summary>Ctrl (OnKey <c>^</c>).</summary>
    Ctrl = 1,

    /// <summary>Alt (OnKey <c>%</c>).</summary>
    Alt = 2,

    /// <summary>Shift (OnKey <c>+</c>).</summary>
    Shift = 4,
}
