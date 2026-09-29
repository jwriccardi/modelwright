using System;
using System.Reflection;
using System.Threading;
using ExcelDna.Integration;

namespace ExcelModelingToolkit.AddIn;

/// <summary>
/// Status-bar messages through the C API (<c>xlcMessage</c>). Each message is cleared after a timeout so a
/// stale message cannot pass for fresh feedback (spike K3 lesson). Call <see cref="Show"/> in macro context.
/// </summary>
internal static class StatusBar
{
    /// <summary>Longest text Excel's status bar accepts; longer messages are truncated.</summary>
    private const int MaxLength = 255;

    private static readonly TimeSpan ClearAfter = TimeSpan.FromSeconds(10);
    private static Timer? _clearTimer;
    private static int _generation;

    // The message we are showing and have not cleared yet (null = no clear pending). Main thread only.
    private static string? _shownMessage;

    /// <summary>Shows <paramref name="message"/> and schedules it to be cleared. Never throws.</summary>
    public static void Show(string message)
    {
        var text = Truncate(message);
        try
        {
            XlCall.Excel(XlCall.xlcMessage, true, text);
        }
        catch (XlCallException)
        {
            return;
        }

        _shownMessage = text;
        var generation = ++_generation;
        _clearTimer?.Dispose();
        _clearTimer = new Timer(_ => QueueClear(generation), null, ClearAfter, Timeout.InfiniteTimeSpan);
    }

    /// <summary>Stops the timer and clears a message that is still pending (call from AutoClose). Never throws.</summary>
    public static void Shutdown()
    {
        _clearTimer?.Dispose();
        _clearTimer = null;
        _generation++; // A clear already queued as a macro becomes a no-op.

        if (_shownMessage == null)
        {
            return;
        }

        _shownMessage = null;
        try
        {
            XlCall.Excel(XlCall.xlcMessage, false);
        }
        catch (XlCallException)
        {
            // Excel is shutting down; the status bar goes with it.
        }
    }

    private static void QueueClear(int generation)
    {
        try
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                // Only clear our own latest message; a newer Show has rescheduled.
                if (generation != _generation || _shownMessage == null)
                {
                    return;
                }

                var shown = _shownMessage;
                _shownMessage = null;

                // Another add-in or Excel may have replaced our text since; leave theirs alone.
                if (!IsShowing(shown))
                {
                    return;
                }

                try
                {
                    XlCall.Excel(XlCall.xlcMessage, false);
                }
                catch (XlCallException)
                {
                    // Nothing useful to do if Excel refuses.
                }
            });
        }
        catch (Exception)
        {
            // The add-in may be unloading; the message simply stays until Excel replaces it.
        }
    }

    /// <summary>
    /// True if <c>Application.StatusBar</c> still shows <paramref name="message"/> (it returns FALSE when Excel
    /// owns the status bar). Main thread only. False if the value cannot be read.
    /// </summary>
    private static bool IsShowing(string message)
    {
        try
        {
            var application = ExcelDnaUtil.Application;
            var current = application.GetType().InvokeMember("StatusBar", BindingFlags.GetProperty, null, application, null);
            return current is string text && string.Equals(text, message, StringComparison.Ordinal);
        }
        catch (Exception)
        {
            return false;
        }
    }

    private static string Truncate(string message)
    {
        if (message.Length <= MaxLength)
        {
            return message;
        }

        // Do not split a surrogate pair.
        var length = char.IsHighSurrogate(message[MaxLength - 1]) ? MaxLength - 1 : MaxLength;
        return message.Substring(0, length);
    }
}
