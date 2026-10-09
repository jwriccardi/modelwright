using System;
using System.Globalization;
using Modelwright.Core.Json;

namespace Modelwright.Core.Trace;

/// <summary>
/// What the Trace In window remembers between sessions: its bounds, whether the formula header wraps, and whether
/// the tree shows "Evaluate functions &amp; groups" (Ctrl+E) rather than the classic view. Kept in
/// its own best-effort file (<c>ui-state.json</c>), not in the settings file: it changes every time the window
/// moves, and a lost or damaged copy only costs the default position. Reading never fails: anything that cannot be
/// used is ignored. The file also remembers that the one-time Macabacus notice was shown
/// (<see cref="MacabacusNoticeShown"/>, written only when true).
/// </summary>
/// <example>
/// <code>
/// {
///   "schemaVersion": 1,
///   "traceWindow": { "left": 1200, "top": 300, "width": 620, "height": 380, "wrapFormula": false, "evaluateFunctions": false },
///   "macabacusNoticeShown": true
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
    /// <param name="evaluateFunctions">True if the tree follows the formula's structure (evaluate mode).</param>
    /// <param name="macabacusNoticeShown">True once the one-time Macabacus notice has been shown.</param>
    public TraceUiState(
        WindowRect? bounds = null,
        bool wrapFormula = false,
        bool evaluateFunctions = false,
        bool macabacusNoticeShown = false)
    {
        Bounds = bounds;
        WrapFormula = wrapFormula;
        EvaluateFunctions = evaluateFunctions;
        MacabacusNoticeShown = macabacusNoticeShown;
    }

    /// <summary>The window's last bounds in screen pixels, or null if not known.</summary>
    public WindowRect? Bounds { get; }

    /// <summary>True if the formula header wraps.</summary>
    public bool WrapFormula { get; }

    /// <summary>
    /// True if the tree shows "Evaluate functions &amp; groups" (<see cref="EvaluatePrecedents"/>); false (the default)
    /// for the classic view, one row per reference.
    /// </summary>
    public bool EvaluateFunctions { get; }

    /// <summary>
    /// True once the one-time notice that Macabacus is also loaded has been shown (<see cref="Modelwright.Core.Settings.Coexistence"/>).
    /// </summary>
    public bool MacabacusNoticeShown { get; }

    /// <summary>This state with other bounds.</summary>
    public TraceUiState WithBounds(WindowRect? bounds) =>
        new TraceUiState(bounds, WrapFormula, EvaluateFunctions, MacabacusNoticeShown);

    /// <summary>This state with another wrap setting.</summary>
    public TraceUiState WithWrapFormula(bool wrapFormula) =>
        new TraceUiState(Bounds, wrapFormula, EvaluateFunctions, MacabacusNoticeShown);

    /// <summary>This state with evaluate mode on or off.</summary>
    public TraceUiState WithEvaluateFunctions(bool evaluateFunctions) =>
        new TraceUiState(Bounds, WrapFormula, evaluateFunctions, MacabacusNoticeShown);

    /// <summary>This state with the Macabacus notice marked as shown or not.</summary>
    public TraceUiState WithMacabacusNoticeShown(bool macabacusNoticeShown) =>
        new TraceUiState(Bounds, WrapFormula, EvaluateFunctions, macabacusNoticeShown);

    /// <summary>
    /// Reads a state written by <see cref="ToJson"/>. Never throws: null, text that is not JSON, or values of the
    /// wrong type or out of range give the defaults for what they affect (bounds need all four numbers, a positive
    /// size, and every coordinate within a million pixels). Unknown properties are ignored, so a newer file still
    /// reads, and a file from before evaluate mode reads with it off. <c>macabacusNoticeShown</c> is read from the
    /// top level, with or without a usable <c>traceWindow</c>; anything but <c>true</c> reads as false.
    /// </summary>
    public static TraceUiState FromJson(string? json)
    {
        if (json is null)
        {
            return new TraceUiState();
        }

        try
        {
            if (JsonParser.Parse(json) is not JsonObject root)
            {
                return new TraceUiState();
            }

            var noticeShown = root.TryGet("macabacusNoticeShown", out var noticeValue) && noticeValue is true;
            if (!root.TryGet("traceWindow", out var value) || value is not JsonObject window)
            {
                return new TraceUiState(macabacusNoticeShown: noticeShown);
            }

            WindowRect? bounds = null;
            if (TryInt(window, "left", out var left) && TryInt(window, "top", out var top) &&
                TryInt(window, "width", out var width) && TryInt(window, "height", out var height) &&
                width > 0 && height > 0)
            {
                bounds = new WindowRect(left, top, width, height);
            }

            var wrap = window.TryGet("wrapFormula", out var wrapValue) && wrapValue is true;
            var evaluate = window.TryGet("evaluateFunctions", out var evaluateValue) && evaluateValue is true;
            return new TraceUiState(bounds, wrap, evaluate, noticeShown);
        }
        catch (FormatException)
        {
            return new TraceUiState();
        }
    }

    /// <summary>Writes the state as indented JSON (bounds left out when unknown, the notice flag when false).</summary>
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
        window.Add("evaluateFunctions", EvaluateFunctions);
        var root = new JsonObject();
        root.Add("schemaVersion", SchemaVersion);
        root.Add("traceWindow", window);
        if (MacabacusNoticeShown)
        {
            root.Add("macabacusNoticeShown", true);
        }

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
