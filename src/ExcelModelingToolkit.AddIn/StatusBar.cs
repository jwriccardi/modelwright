using System;
using System.Threading;
using ExcelDna.Integration;

namespace ExcelModelingToolkit.AddIn;

/// <summary>
/// Status-bar messages through the C API (<c>xlcMessage</c>). Each message is cleared after a timeout so a
/// stale message cannot pass for fresh feedback (spike K3 lesson). Call <see cref="Show"/> in macro context.
/// </summary>
internal static class StatusBar
{
    private static readonly TimeSpan ClearAfter = TimeSpan.FromSeconds(10);
    private static Timer? _clearTimer;
    private static int _generation;

    /// <summary>Shows <paramref name="message"/> and schedules it to be cleared. Never throws.</summary>
    public static void Show(string message)
    {
        try
        {
            XlCall.Excel(XlCall.xlcMessage, true, message);
        }
        catch (XlCallException)
        {
            return;
        }

        var generation = ++_generation;
        _clearTimer?.Dispose();
        _clearTimer = new Timer(_ => QueueClear(generation), null, ClearAfter, Timeout.InfiniteTimeSpan);
    }

    /// <summary>Stops any pending clear (call from AutoClose).</summary>
    public static void Shutdown()
    {
        _clearTimer?.Dispose();
        _clearTimer = null;
    }

    private static void QueueClear(int generation)
    {
        try
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                // Only clear our own latest message; a newer Show has rescheduled.
                if (generation != _generation)
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
}
