using System;
using System.Collections.Generic;
using Modelwright.Core.Trace;
using Xunit;

namespace Modelwright.Core.Tests.Trace;

public class SerialCommandQueueTests
{
    // Stands in for the window the add-in posts drains to: they run when the test pumps "messages".
    private sealed class Loop
    {
        private readonly Queue<Action> _posted = new Queue<Action>();

        public bool Accepts { get; set; } = true;

        public int Posted { get; private set; }

        public bool Post(Action work)
        {
            if (!Accepts)
            {
                return false;
            }

            Posted++;
            _posted.Enqueue(work);
            return true;
        }

        public void Pump()
        {
            while (_posted.Count > 0)
            {
                _posted.Dequeue()();
            }
        }
    }

    [Fact]
    public void Commands_run_later_in_the_order_they_arrived_from_one_drain()
    {
        var loop = new Loop();
        var queue = new SerialCommandQueue(loop.Post, ex => throw ex);
        var ran = new List<string>();

        Assert.True(queue.Enqueue(() => ran.Add("key Down")));
        Assert.True(queue.Enqueue(() => ran.Add("click row 3")));
        Assert.True(queue.Enqueue(() => ran.Add("OK")));

        Assert.Empty(ran);
        Assert.Equal(1, loop.Posted);
        loop.Pump();
        Assert.Equal(new[] { "key Down", "click row 3", "OK" }, ran);
        Assert.Equal(0, queue.Count);
    }

    [Fact]
    public void A_command_that_pumps_messages_is_not_overtaken_and_nothing_is_posted_again()
    {
        var loop = new Loop();
        var ran = new List<string>();
        var queue = new SerialCommandQueue(loop.Post, ex => throw ex);
        queue.Enqueue(() =>
        {
            ran.Add("expand (opens a workbook)");

            // While it pumps, a key and a click arrive, and a stray drain is dispatched.
            queue.Enqueue(() => ran.Add("key Esc"));
            queue.Enqueue(() => ran.Add("click"));
            queue.Drain();
            ran.Add("expand done");
        });

        loop.Pump();

        Assert.Equal(new[] { "expand (opens a workbook)", "expand done", "key Esc", "click" }, ran);
        Assert.Equal(1, loop.Posted);
        Assert.False(queue.IsRunning);
    }

    [Fact]
    public void After_a_drain_the_next_command_schedules_a_new_one()
    {
        var loop = new Loop();
        var queue = new SerialCommandQueue(loop.Post, ex => throw ex);
        var ran = new List<int>();
        queue.Enqueue(() => ran.Add(1));
        loop.Pump();

        queue.Enqueue(() => ran.Add(2));
        loop.Pump();

        Assert.Equal(new[] { 1, 2 }, ran);
        Assert.Equal(2, loop.Posted);
    }

    [Fact]
    public void A_failing_command_is_reported_and_the_rest_still_run()
    {
        var loop = new Loop();
        var errors = new List<string>();
        var queue = new SerialCommandQueue(loop.Post, ex => errors.Add(ex.Message));
        var ran = new List<int>();
        queue.Enqueue(() => throw new InvalidOperationException("boom"));
        queue.Enqueue(() => ran.Add(2));

        loop.Pump();

        Assert.Equal(new[] { "boom" }, errors);
        Assert.Equal(new[] { 2 }, ran);
    }

    [Fact]
    public void A_failing_error_handler_does_not_stop_the_queue()
    {
        var loop = new Loop();
        var queue = new SerialCommandQueue(loop.Post, ex => throw new InvalidOperationException("handler"));
        var ran = new List<int>();
        queue.Enqueue(() => throw new InvalidOperationException("boom"));
        queue.Enqueue(() => ran.Add(2));

        loop.Pump();

        Assert.Equal(new[] { 2 }, ran);
    }

    [Fact]
    public void A_command_is_not_queued_when_no_drain_can_be_scheduled()
    {
        var loop = new Loop { Accepts = false };
        var queue = new SerialCommandQueue(loop.Post, ex => throw ex);

        Assert.False(queue.Enqueue(() => { }));
        Assert.Equal(0, queue.Count);

        loop.Accepts = true;
        Assert.True(queue.Enqueue(() => { }));
        Assert.Equal(1, queue.Count);
    }

    [Fact]
    public void Clear_drops_what_has_not_run()
    {
        var loop = new Loop();
        var queue = new SerialCommandQueue(loop.Post, ex => throw ex);
        var ran = new List<int>();
        queue.Enqueue(() => ran.Add(1));
        queue.Clear();

        loop.Pump();

        Assert.Empty(ran);
    }

    [Fact]
    public void Arguments_are_checked()
    {
        Assert.Throws<ArgumentNullException>(() => new SerialCommandQueue(null!, ex => { }));
        Assert.Throws<ArgumentNullException>(() => new SerialCommandQueue(work => true, null!));
        Assert.Throws<ArgumentNullException>(() => new SerialCommandQueue(work => true, ex => { }).Enqueue(null!));
    }
}
