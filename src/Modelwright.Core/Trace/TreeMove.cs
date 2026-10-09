namespace Modelwright.Core.Trace;

/// <summary>What a cursor key did to a <see cref="PrecedentTree"/>.</summary>
public enum TreeMove
{
    /// <summary>Nothing: already at the edge, or nothing to expand.</summary>
    None,

    /// <summary>The selection moved to another node; the add-in navigates Excel to it.</summary>
    Moved,

    /// <summary>The selected node was expanded; the selection did not move.</summary>
    Expanded,

    /// <summary>The selected node was collapsed; the selection did not move.</summary>
    Collapsed,
}
