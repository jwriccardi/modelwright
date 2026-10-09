using System;
using System.Collections.Generic;
using System.Linq;
using Modelwright.Core.Formatting;
using Modelwright.Core.Undo;
using Xunit;

namespace Modelwright.Core.Tests.Undo;

public class UndoManagerTests
{
    private static FormatSnapshot Snapshot(string label, string workbook = "Book1.xlsx", string sheet = "Sheet1", int blocks = 1) =>
        FormatSnapshot.Create(label, CycleKind.NumberFormat, workbook, sheet, Enumerable.Range(1, blocks).Select(row =>
            new SnapshotBlock(new CellRect(row, 1, 1, 1), CycleValue.FromNumberFormat("General"), CycleValue.FromNumberFormat("0.0"))));

    private static FormatSnapshot Barrier(string label, string reason = "too many formats") =>
        FormatSnapshot.Unavailable(label, CycleKind.NumberFormat, "Book1.xlsx", "Sheet1", reason);

    [Theory]
    // Ctrl+Z: Excel's newer history goes first; our stale redo is dropped when it does.
    [InlineData(UndoKey.Undo, 0, 0, null, null, UndoDecision.PassToExcel)]
    [InlineData(UndoKey.Undo, 0, 0, false, false, UndoDecision.PassToExcel)]
    [InlineData(UndoKey.Undo, 0, 0, true, false, UndoDecision.PassToExcel)]
    [InlineData(UndoKey.Undo, 1, 0, false, null, UndoDecision.HandleOurs)]
    [InlineData(UndoKey.Undo, 1, 0, false, true, UndoDecision.HandleOurs)]
    [InlineData(UndoKey.Undo, 1, 0, true, false, UndoDecision.PassToExcel)]
    [InlineData(UndoKey.Undo, 1, 0, null, false, UndoDecision.PassToExcel)]
    [InlineData(UndoKey.Undo, 1, 2, true, false, UndoDecision.PassToExcelAndClearRedo)]
    [InlineData(UndoKey.Undo, 0, 2, true, null, UndoDecision.PassToExcelAndClearRedo)]
    [InlineData(UndoKey.Undo, 0, 2, false, false, UndoDecision.PassToExcel)]
    [InlineData(UndoKey.Undo, 1, 2, false, false, UndoDecision.HandleOurs)]
    [InlineData(UndoKey.Undo, 1, 2, null, true, UndoDecision.PassToExcel)]
    // Ctrl+Y: ours only when Excel has no history at all; any history makes our redo stale.
    [InlineData(UndoKey.Redo, 0, 0, false, false, UndoDecision.PassToExcel)]
    [InlineData(UndoKey.Redo, 3, 0, true, true, UndoDecision.PassToExcel)]
    [InlineData(UndoKey.Redo, 0, 1, false, false, UndoDecision.HandleOurs)]
    [InlineData(UndoKey.Redo, 2, 1, false, false, UndoDecision.HandleOurs)]
    [InlineData(UndoKey.Redo, 0, 1, true, false, UndoDecision.PassToExcelAndClearRedo)]
    [InlineData(UndoKey.Redo, 0, 1, false, true, UndoDecision.PassToExcelAndClearRedo)]
    [InlineData(UndoKey.Redo, 0, 1, null, true, UndoDecision.PassToExcelAndClearRedo)]
    [InlineData(UndoKey.Redo, 0, 1, true, null, UndoDecision.PassToExcelAndClearRedo)]
    [InlineData(UndoKey.Redo, 0, 1, null, false, UndoDecision.PassToExcel)]
    [InlineData(UndoKey.Redo, 0, 1, false, null, UndoDecision.PassToExcel)]
    [InlineData(UndoKey.Redo, 0, 1, null, null, UndoDecision.PassToExcel)]
    public void Decide_follows_the_ordering_and_staleness_rules(
        UndoKey key, int ourUndo, int ourRedo, bool? nativeUndo, bool? nativeRedo, UndoDecision expected)
    {
        Assert.Equal(expected, UndoManager.Decide(key, ourUndo, ourRedo, nativeUndo, nativeRedo));
    }

    [Fact]
    public void Decide_rejects_an_undefined_key_and_negative_counts()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => UndoManager.Decide((UndoKey)7, 1, 0, false, false));
        Assert.Throws<ArgumentOutOfRangeException>(() => UndoManager.Decide(UndoKey.Undo, -1, 0, false, false));
        Assert.Throws<ArgumentOutOfRangeException>(() => UndoManager.Decide(UndoKey.Redo, 0, -1, false, false));
    }

    [Fact]
    public void New_manager_is_empty()
    {
        var manager = new UndoManager();

        Assert.Equal(UndoManager.DefaultMaxDepth, manager.MaxDepth);
        Assert.Equal(100, manager.MaxDepth);
        Assert.Equal(UndoManager.DefaultMaxBlocks, manager.MaxBlocks);
        Assert.Equal(200000, manager.MaxBlocks);
        Assert.Equal(0, manager.UndoCount);
        Assert.Equal(0, manager.RedoCount);
        Assert.Equal(0, manager.BlockCount);
        Assert.Null(manager.Peek(UndoKey.Undo));
        Assert.Null(manager.Peek(UndoKey.Redo));
        Assert.Null(manager.Undo());
        Assert.Null(manager.Redo());
        Assert.Null(manager.Discard(UndoKey.Undo));
        Assert.Equal(0, manager.ClearRedo());
    }

    [Fact]
    public void Undo_moves_the_newest_to_redo_and_redo_moves_it_back()
    {
        var manager = new UndoManager();
        var first = Snapshot("first");
        var second = Snapshot("second");
        manager.Push(first);
        manager.Push(second);

        Assert.Same(second, manager.Peek(UndoKey.Undo));
        Assert.Same(second, manager.Undo());
        Assert.Equal(1, manager.UndoCount);
        Assert.Equal(1, manager.RedoCount);
        Assert.Same(first, manager.Peek(UndoKey.Undo));
        Assert.Same(second, manager.Peek(UndoKey.Redo));

        Assert.Same(first, manager.Undo());
        Assert.Same(first, manager.Peek(UndoKey.Redo));
        Assert.Equal(2, manager.Count(UndoKey.Redo));

        Assert.Same(first, manager.Redo());
        Assert.Same(second, manager.Redo());
        Assert.Equal(2, manager.Count(UndoKey.Undo));
        Assert.Equal(0, manager.Count(UndoKey.Redo));
        Assert.Same(second, manager.Peek(UndoKey.Undo));
    }

    [Fact]
    public void Push_clears_the_redo_stack()
    {
        var manager = new UndoManager();
        manager.Push(Snapshot("a"));
        manager.Undo();
        Assert.Equal(1, manager.RedoCount);

        manager.Push(Snapshot("b"));

        Assert.Equal(0, manager.RedoCount);
        Assert.Equal(1, manager.UndoCount);
    }

    [Fact]
    public void Push_beyond_max_depth_drops_the_oldest()
    {
        var manager = new UndoManager(maxDepth: 3);
        var snapshots = new List<FormatSnapshot>();
        for (var i = 0; i < 5; i++)
        {
            snapshots.Add(Snapshot("s" + i));
            manager.Push(snapshots[i]);
        }

        Assert.Equal(3, manager.UndoCount);
        Assert.Same(snapshots[4], manager.Undo());
        Assert.Same(snapshots[3], manager.Undo());
        Assert.Same(snapshots[2], manager.Undo());
        Assert.Null(manager.Undo());
    }

    [Fact]
    public void Default_depth_keeps_one_hundred()
    {
        var manager = new UndoManager();
        for (var i = 0; i < 150; i++)
        {
            manager.Push(Snapshot("s" + i));
        }

        Assert.Equal(100, manager.UndoCount);
    }

    [Fact]
    public void Push_beyond_the_block_budget_evicts_the_oldest_snapshots()
    {
        var manager = new UndoManager(maxBlocks: 5);
        var a = Snapshot("a", blocks: 2);
        var b = Snapshot("b", blocks: 2);
        var c = Snapshot("c", blocks: 2);
        manager.Push(a);
        manager.Push(b);
        Assert.Equal(4, manager.BlockCount);

        manager.Push(c);

        Assert.Equal(2, manager.UndoCount);
        Assert.Equal(4, manager.BlockCount);
        Assert.Same(c, manager.Undo());
        Assert.Same(b, manager.Undo());
        Assert.Null(manager.Undo());
        Assert.Equal(4, manager.BlockCount); // on the redo stack now
    }

    [Fact]
    public void The_newest_snapshot_is_kept_even_if_it_alone_exceeds_the_budget()
    {
        var manager = new UndoManager(maxBlocks: 5);
        manager.Push(Snapshot("small", blocks: 1));
        manager.Push(Barrier("barrier"));
        var big = Snapshot("big", blocks: 7);

        manager.Push(big);

        Assert.Equal(1, manager.UndoCount);
        Assert.Same(big, manager.Peek(UndoKey.Undo));
        Assert.Equal(7, manager.BlockCount);
    }

    [Fact]
    public void Barriers_cost_no_blocks_and_stay_within_the_budget()
    {
        var manager = new UndoManager(maxBlocks: 2);
        manager.Push(Snapshot("a", blocks: 2));
        manager.Push(Barrier("b"));

        Assert.Equal(2, manager.UndoCount);
        Assert.Equal(2, manager.BlockCount);
    }

    [Fact]
    public void Push_accepts_a_barrier_and_clears_the_redo_stack()
    {
        var manager = new UndoManager();
        manager.Push(Snapshot("a"));
        manager.Undo();
        var barrier = Barrier("b");

        manager.Push(barrier);

        Assert.Equal(0, manager.RedoCount);
        Assert.Same(barrier, manager.Peek(UndoKey.Undo));
        Assert.Throws<ArgumentNullException>(() => manager.Push(null!));
    }

    [Fact]
    public void A_barrier_cannot_be_undone_only_discarded()
    {
        var manager = new UndoManager();
        var a = Snapshot("a");
        manager.Push(a);
        var barrier = Barrier("b");
        manager.Push(barrier);

        var error = Assert.Throws<InvalidOperationException>(() => manager.Undo());
        Assert.Contains("barrier", error.Message);
        Assert.Equal(2, manager.UndoCount);
        Assert.Equal(0, manager.RedoCount);

        Assert.Same(barrier, manager.Discard(UndoKey.Undo));
        Assert.Same(a, manager.Undo());
    }

    [Fact]
    public void ClearRedo_empties_only_the_redo_stack()
    {
        var manager = new UndoManager();
        manager.Push(Snapshot("a"));
        manager.Push(Snapshot("b"));
        manager.Push(Snapshot("c"));
        manager.Undo();
        manager.Undo();

        Assert.Equal(2, manager.ClearRedo());

        Assert.Equal(0, manager.RedoCount);
        Assert.Equal(1, manager.UndoCount);
    }

    [Fact]
    public void Constructor_rejects_a_depth_or_budget_below_one()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new UndoManager(0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new UndoManager(maxBlocks: 0));
    }

    [Fact]
    public void Discard_drops_only_the_top_of_that_stack()
    {
        var manager = new UndoManager();
        var a = Snapshot("a");
        var b = Snapshot("b");
        var c = Snapshot("c");
        manager.Push(a);
        manager.Push(b);
        manager.Push(c);
        manager.Undo();

        Assert.Same(b, manager.Discard(UndoKey.Undo));
        Assert.Same(a, manager.Peek(UndoKey.Undo));
        Assert.Same(c, manager.Peek(UndoKey.Redo));

        Assert.Same(c, manager.Discard(UndoKey.Redo));
        Assert.Equal(0, manager.RedoCount);
        Assert.Equal(1, manager.UndoCount);
    }

    [Fact]
    public void InvalidateWorkbook_drops_its_snapshots_from_both_stacks_ignoring_case()
    {
        var manager = new UndoManager();
        manager.Push(Snapshot("a1", "Model.xlsx"));
        manager.Push(Snapshot("b1", "Other.xlsx"));
        manager.Push(Snapshot("a2", "Model.xlsx", "Inputs"));
        manager.Undo(); // a2 on redo

        Assert.Equal(2, manager.InvalidateWorkbook("MODEL.XLSX"));

        Assert.Equal(1, manager.UndoCount);
        Assert.Equal(0, manager.RedoCount);
        Assert.Equal("b1", manager.Peek(UndoKey.Undo)!.Label);
    }

    [Fact]
    public void InvalidateWorkbook_matches_the_full_name_so_a_same_name_file_elsewhere_is_kept()
    {
        var manager = new UndoManager();
        manager.Push(Snapshot("here", @"C:\Deals\Model.xlsx"));
        manager.Push(Snapshot("there", @"C:\Archive\Model.xlsx"));
        manager.Push(Barrier("barrier"));

        Assert.Equal(1, manager.InvalidateWorkbook(@"c:\deals\MODEL.xlsx"));

        Assert.Equal(new[] { "barrier", "there" }, new[] { manager.Discard(UndoKey.Undo)!.Label, manager.Discard(UndoKey.Undo)!.Label });
        Assert.Equal(0, manager.UndoCount);
    }

    [Fact]
    public void InvalidateSheet_drops_only_that_sheet()
    {
        var manager = new UndoManager();
        manager.Push(Snapshot("keep", "Model.xlsx", "Inputs"));
        manager.Push(Snapshot("drop1", "Model.xlsx", "Calc"));
        manager.Push(Snapshot("other book", "Other.xlsx", "Calc"));
        manager.Push(Snapshot("drop2", "Model.xlsx", "calc"));

        Assert.Equal(2, manager.InvalidateSheet("model.xlsx", "CALC"));

        Assert.Equal("other book", manager.Undo()!.Label);
        Assert.Equal("keep", manager.Undo()!.Label);
        Assert.Null(manager.Undo());
    }

    [Fact]
    public void Invalidate_rejects_null_names()
    {
        var manager = new UndoManager();

        Assert.Throws<ArgumentNullException>(() => manager.InvalidateWorkbook(null!));
        Assert.Throws<ArgumentNullException>(() => manager.InvalidateSheet(null!, "Sheet1"));
        Assert.Throws<ArgumentNullException>(() => manager.InvalidateSheet("Book1.xlsx", null!));
    }

    [Fact]
    public void Clear_empties_both_stacks()
    {
        var manager = new UndoManager();
        manager.Push(Snapshot("a"));
        manager.Push(Snapshot("b"));
        manager.Undo();

        manager.Clear();

        Assert.Equal(0, manager.UndoCount);
        Assert.Equal(0, manager.RedoCount);
    }

    [Fact]
    public void Type_cycle_type_then_three_ctrl_z_undoes_the_typing_then_our_cycle_then_nothing()
    {
        // docs/PLAN.md section 4.4: our cycle wiped Excel's history (the first typing is gone), and the second
        // typing is newer than our cycle, so Excel must undo it first.
        var excel = new SimulatedExcel();
        excel.Type();
        excel.Cycle(Snapshot("Number"));
        excel.Type();

        Assert.Equal("Excel undid typing", excel.Press(UndoKey.Undo));
        Assert.Equal("we undid Number", excel.Press(UndoKey.Undo));
        Assert.Equal("Excel had nothing to undo", excel.Press(UndoKey.Undo));
        Assert.Equal(1, excel.Manager.RedoCount);
    }

    [Fact]
    public void Three_cycles_then_three_ctrl_z_then_three_ctrl_y()
    {
        var excel = new SimulatedExcel();
        excel.Cycle(Snapshot("one"));
        excel.Cycle(Snapshot("two"));
        excel.Cycle(Snapshot("three"));

        Assert.Equal("we undid three", excel.Press(UndoKey.Undo));
        Assert.Equal("we undid two", excel.Press(UndoKey.Undo));
        Assert.Equal("we undid one", excel.Press(UndoKey.Undo));
        Assert.Equal("Excel had nothing to undo", excel.Press(UndoKey.Undo));

        Assert.Equal("we redid one", excel.Press(UndoKey.Redo));
        Assert.Equal("we redid two", excel.Press(UndoKey.Redo));
        Assert.Equal("we redid three", excel.Press(UndoKey.Redo));
        Assert.Equal("Excel had nothing to redo", excel.Press(UndoKey.Redo));
        Assert.Equal(3, excel.Manager.UndoCount);
    }

    [Fact]
    public void Excels_redo_goes_first_when_it_has_one()
    {
        var excel = new SimulatedExcel();
        excel.Cycle(Snapshot("Number"));
        excel.Type();

        Assert.Equal("Excel undid typing", excel.Press(UndoKey.Undo));
        Assert.Equal("Excel redid typing", excel.Press(UndoKey.Redo));
        Assert.Equal("Excel undid typing", excel.Press(UndoKey.Undo));
        Assert.Equal("we undid Number", excel.Press(UndoKey.Undo));

        // Our restore is a COM write too: it wiped Excel's redo of the typing.
        Assert.Equal("we redid Number", excel.Press(UndoKey.Redo));
        Assert.Equal("Excel had nothing to redo", excel.Press(UndoKey.Redo));
    }

    [Fact]
    public void Unknown_native_state_always_passes_the_key()
    {
        var excel = new SimulatedExcel();
        excel.Cycle(Snapshot("Number"));
        excel.Unknown = true; // e.g. a cell is being edited

        Assert.Equal("Excel had nothing to undo", excel.Press(UndoKey.Undo));
        Assert.Equal(1, excel.Manager.UndoCount);
    }

    [Fact]
    public void A_new_cycle_after_an_undo_clears_our_redo()
    {
        var excel = new SimulatedExcel();
        excel.Cycle(Snapshot("one"));
        Assert.Equal("we undid one", excel.Press(UndoKey.Undo));
        excel.Cycle(Snapshot("two"));

        Assert.Equal("Excel had nothing to redo", excel.Press(UndoKey.Redo));
        Assert.Equal("we undid two", excel.Press(UndoKey.Undo));
    }

    [Fact]
    public void Ctrl_z_stops_at_an_unrecorded_change_and_does_not_touch_the_older_one()
    {
        var excel = new SimulatedExcel();
        excel.Cycle(Snapshot("A"));
        excel.Unrecorded("B", "too many formats");

        Assert.Equal("Can't undo: B could not be recorded (too many formats)", excel.Press(UndoKey.Undo));
        Assert.Equal(1, excel.Manager.UndoCount);
        Assert.Equal("A", excel.Manager.Peek(UndoKey.Undo)!.Label);
        Assert.Equal(0, excel.Manager.RedoCount);

        Assert.Equal("we undid A", excel.Press(UndoKey.Undo));
        Assert.Equal("we redid A", excel.Press(UndoKey.Redo));
    }

    [Fact]
    public void An_unrecorded_change_clears_our_redo()
    {
        var excel = new SimulatedExcel();
        excel.Cycle(Snapshot("A"));
        Assert.Equal("we undid A", excel.Press(UndoKey.Undo));
        excel.Unrecorded("B", "pattern or gradient fill");

        Assert.Equal(0, excel.Manager.RedoCount);
        Assert.Equal("Excel had nothing to redo", excel.Press(UndoKey.Redo));
        Assert.Equal("Can't undo: B could not be recorded (pattern or gradient fill)", excel.Press(UndoKey.Undo));
        Assert.Equal("Excel had nothing to undo", excel.Press(UndoKey.Undo));
    }

    [Fact]
    public void Our_undo_then_typing_then_ctrl_y_drops_our_stale_redo_and_passes()
    {
        var excel = new SimulatedExcel();
        excel.Cycle(Snapshot("A"));
        Assert.Equal("we undid A", excel.Press(UndoKey.Undo));
        excel.Type();

        Assert.Equal("Excel had nothing to redo", excel.Press(UndoKey.Redo));
        Assert.Equal(0, excel.Manager.RedoCount);

        // Our redo of A must not come back once the typing is undone.
        Assert.Equal("Excel undid typing", excel.Press(UndoKey.Undo));
        Assert.Equal("Excel redid typing", excel.Press(UndoKey.Redo));
        Assert.Equal("Excel had nothing to redo", excel.Press(UndoKey.Redo));
    }

    [Fact]
    public void Our_undo_then_typing_then_native_ctrl_z_then_two_ctrl_y()
    {
        var excel = new SimulatedExcel();
        excel.Cycle(Snapshot("A"));
        excel.Cycle(Snapshot("B"));
        Assert.Equal("we undid B", excel.Press(UndoKey.Undo));
        excel.Type();

        // Excel's undo is newer than our restore: it goes first, and our redo of B is stale.
        Assert.Equal("Excel undid typing", excel.Press(UndoKey.Undo));
        Assert.Equal(0, excel.Manager.RedoCount);

        Assert.Equal("Excel redid typing", excel.Press(UndoKey.Redo));
        Assert.Equal("Excel had nothing to redo", excel.Press(UndoKey.Redo));

        // Older history is intact.
        Assert.Equal("Excel undid typing", excel.Press(UndoKey.Undo));
        Assert.Equal("we undid A", excel.Press(UndoKey.Undo));
    }

    /// <summary>
    /// Excel's native undo and redo as counts, per spike K2c: every COM write (our cycle, recorded or not, and our
    /// restore) wipes both; typing adds an undo entry and clears redo. Keys go through
    /// <see cref="UndoManager.Decide"/>, and a barrier is handled like the add-in's UndoCommand does.
    /// </summary>
    private sealed class SimulatedExcel
    {
        private int _nativeUndo;
        private int _nativeRedo;

        public UndoManager Manager { get; } = new UndoManager();

        public bool Unknown { get; set; }

        public void Type()
        {
            _nativeUndo++;
            _nativeRedo = 0;
        }

        public void Cycle(FormatSnapshot snapshot)
        {
            Manager.Push(snapshot);
            WipeNative();
        }

        public void Unrecorded(string label, string reason) => Cycle(Barrier(label, reason));

        public string Press(UndoKey key)
        {
            var nativeUndo = Unknown ? (bool?)null : _nativeUndo > 0;
            var nativeRedo = Unknown ? (bool?)null : _nativeRedo > 0;
            var decision = UndoManager.Decide(key, Manager.UndoCount, Manager.RedoCount, nativeUndo, nativeRedo);
            if (decision == UndoDecision.PassToExcelAndClearRedo)
            {
                Manager.ClearRedo();
            }

            if (decision == UndoDecision.HandleOurs)
            {
                var top = Manager.Peek(key)!;
                if (!top.IsAvailable)
                {
                    Manager.Discard(key);
                    return $"Can't undo: {top.Label} could not be recorded ({top.UnavailableReason})";
                }

                var snapshot = key == UndoKey.Undo ? Manager.Undo() : Manager.Redo();
                WipeNative();
                return (key == UndoKey.Undo ? "we undid " : "we redid ") + snapshot!.Label;
            }

            if (key == UndoKey.Undo)
            {
                if (Unknown || _nativeUndo == 0)
                {
                    return "Excel had nothing to undo";
                }

                _nativeUndo--;
                _nativeRedo++;
                return "Excel undid typing";
            }

            if (Unknown || _nativeRedo == 0)
            {
                return "Excel had nothing to redo";
            }

            _nativeRedo--;
            _nativeUndo++;
            return "Excel redid typing";
        }

        private void WipeNative()
        {
            _nativeUndo = 0;
            _nativeRedo = 0;
        }
    }
}
