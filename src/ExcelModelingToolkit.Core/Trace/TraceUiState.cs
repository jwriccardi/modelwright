using System;
using System.Globalization;
using ExcelModelingToolkit.Core.Json;

namespace ExcelModelingToolkit.Core.Trace;

/// <summary>
/// What the Trace In window remembers between sessions: its bounds and whether the formula header wraps. Kept in
/// its own best-effort file (<c>ui-state.json</c>), not in the settings file: it changes every time the window
/// moves, and a lost or damaged copy only costs the default position. Reading never fails: anything that cannot be
/// used is ignored.
/// </summary>
/// <example>
/// <code>
/// {
///   "schemaVersion": 1,
///   "traceWindow": { "left": 1200, "top": 300, "width": 620, "height": 380, "wrapFormula": false }
/// }
/// </code>
/// </example>
public sealed class TraceUiState
{
    /// <summary>The <c>schemaVersion</c> written.</summary>
    public const int SchemaVersion = 1;

    // A remembered size beyond this is not a real window (and keeps the arithmetic far from overflow).
    private const int MaxCoordinate = 1000000;

    /// <summary>Creates a state.</summary>
    /// <param name="bounds">The window's last bounds, or null to use the default place.</param>
    /// <param name="wrapFormula">True if the formula header wraps long formulas.</param>
    public TraceUiState(WindowRect? bounds = null, bool wrapFormula = false)
    {
        Bounds = bounds;
        WrapFormula = wrapFormula;
    }

    /// <summary>The window's last bounds in screen pixels, or null if not known.</summary>
    public WindowRect? Bounds { get; }

    /// <summary>True if the formula header wraps.</summary>
    public bool WrapFormula { get; }

    /// <summary>This state with other bounds.</summary>
    public TraceUiState WithBounds(WindowRect? bounds) => new TraceUiState(bounds, WrapFormula);

    /// <summary>This state with another wrap setting.</summary>
    public TraceUiState WithWrapFormula(bool wrapFormula) => new TraceUiState(Bounds, wrapFormula);

    /// <summary>
    /// Reads a state written by <see cref="ToJson"/>. Never throws: null, text that is not JSON, or values of the
    /// wrong type or out of range give the defaults for what they affect (bounds need all four numbers, a positive
    /// size, and every coordinate within a million pixels). Unknown properties are ignored, so a newer file still
    /// reads.
    /// </summary>
    public static TraceUiState FromJson(string? json)
    {
        if (json is null)
        {
            return new TraceUiState();
        }

        try
        {
            if (JsonParser.Parse(json) is not JsonObject root ||
                !root.TryGet("traceWindow", out var value) || value is not JsonObject window)
            {
                return new TraceUiState();
            }

            WindowRect? bounds = null;
            if (TryInt(window, "left", out var left) && TryInt(window, "top", out var top) &&
                TryInt(window, "width", out var width) && TryInt(window, "height", out var height) &&
                width > 0 && height > 0)
            {
                bounds = new WindowRect(left, top, width, height);
            }

            var wrap = window.TryGet("wrapFormula", out var wrapValue) && wrapValue is true;
            return new TraceUiState(bounds, wrap);
        }
        catch (FormatException)
        {
            return new TraceUiState();
        }
    }

    /// <summary>Writes the state as indented JSON (bounds left out when unknown).</summary>
    public string ToJson()
    {
        var window = new JsonObject();
        if (Bounds is WindowRect bounds)
        {
            window.Add("left", bounds.Left);
            window.Add("top", bounds.Top);
            window.Add("width", bounds.Width);
            window.Add("height", bounds.Height);
        }

        window.Add("wrapFormula", WrapFormula);
        var root = new JsonObject();
        root.Add("schemaVersion", SchemaVersion);
        root.Add("traceWindow", window);
        return JsonWriter.Write(root);
    }

    private static bool TryInt(JsonObject obj, string name, out int value)
    {
        value = 0;
        return obj.TryGet(name, out var raw) && raw is JsonNumber number &&
            int.TryParse(number.Text, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out value) &&
            value >= -MaxCoordinate && value <= MaxCoordinate;
    }
}
