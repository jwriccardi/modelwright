using System;
using System.Collections.Generic;

namespace ExcelModelingToolkit.Core.Trace;

/// <summary>
/// One first-in, first-out queue for everything the Trace In window does (keys from the hook, clicks, OK and Cancel),
/// run one command at a time on Excel's main thread, in the order they arrived (docs/PLAN.md section 4.5).
/// </summary>
/// <remarks>
/// <para>
/// A command can pump messages while it runs (opening a closed workbook does), so another command can arrive, or a
/// drain be dispatched, in the middle of it. Such a command waits in the queue and the running drain runs it next; a
/// nested drain does nothing. So commands never interleave or overtake each other, and nothing is posted again or
/// spins while one runs.
/// </para>
/// <para>Not thread-safe: one thread (the add-in uses Excel's main thread).</para>
/// </remarks>
public sealed class SerialCommandQueue
{
    private readonly Queue<Action> _pending = new Queue<Action>();
    private readonly Func<Action, bool> _schedule;
    private readonly Action<Exception> _onError;

    // True while a drain has been scheduled and has not started yet.
    private bool _scheduled;

    /// <summary>Creates a queue.</summary>
    /// <param name="schedule">
    /// Arranges for its argument (<see cref="Drain"/>) to run later on this thread, after the caller returns (the add-in
    /// posts it to a window); returns false if it cannot. It must not run it before returning.
    /// </param>
    /// <param name="onError">Told about an exception a command threw; the queue then runs the next command.</param>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    public SerialCommandQueue(Func<Action, bool> schedule, Action<Exception> onError)
    {
        _schedule = schedule ?? throw new ArgumentNullException(nameof(schedule));
        _onError = onError ?? throw new ArgumentNullException(nameof(onError));
    }

    /// <summary>The commands waiting to run.</summary>
    public int Count => _pending.Count;

    /// <summary>True while a command is running (in <see cref="Drain"/>).</summary>
    public bool IsRunning { get; private set; }

    /// <summary>
    /// Adds a command after those already queued and makes sure a drain will run it (one scheduled drain serves any
    /// number of commands). Returns false, without queuing it, if no drain could be scheduled.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="command"/> is null.</exception>
    public bool Enqueue(Action command)
    {
        if (command is null)
        {
            throw new ArgumentNullException(nameof(command));
        }

        if (!IsRunning && !_scheduled)
        {
            if (!_schedule(Drain))
            {
                return false;
            }

            _scheduled = true;
        }

        _pending.Enqueue(command);
        return true;
    }

    /// <summary>
    /// Runs the queued commands in order, including any queued while they run, until none is left. Does nothing if a
    /// drain is already running further up the stack (it will run them).
    /// </summary>
    public void Drain()
    {
        _scheduled = false;
        if (IsRunning)
        {
            return;
        }

        IsRunning = true;
        try
        {
            while (_pending.Count > 0)
            {
                var command = _pending.Dequeue();
                try
                {
                    command();
                }
                catch (Exception ex)
                {
                    Report(ex);
                }
            }
        }
        finally
        {
            IsRunning = false;
        }
    }

    /// <summary>Drops the commands not yet run.</summary>
    public void Clear() => _pending.Clear();

    private void Report(Exception exception)
    {
        try
        {
            _onError(exception);
        }
        catch (Exception)
        {
            // The handler's own failure must not stop the queue.
        }
    }
}
