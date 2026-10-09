using System;

namespace Modelwright.Core.Settings;

/// <summary>What to do once the add-in has looked for Macabacus at startup; see <see cref="Coexistence.Decide"/>.</summary>
[Flags]
public enum CoexistenceActions
{
    /// <summary>Nothing.</summary>
    None = 0,

    /// <summary>
    /// Bind our keys again, so that ours are the most recent bindings (in Excel the most recent binding of a key wins)
    /// over those Macabacus made while it loaded.
    /// </summary>
    Reregister = 1,

    /// <summary>
    /// Show the one-time notice that Macabacus is also loaded (and remember that it was shown), which offers to
    /// switch our shortcuts off.
    /// </summary>
    ShowNotice = 2,
}
