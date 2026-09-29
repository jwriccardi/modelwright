using System;
using System.Collections.Generic;
using ExcelModelingToolkit.Core.Formatting;
using ExcelModelingToolkit.Core.Undo;
using Xunit;

namespace ExcelModelingToolkit.Core.Tests.Undo;

public class UndoManagerTests
{
    private static FormatSnapshot Snapshot(string label, string workbook = "Book1.xlsx", string sheet = "Sheet1") =>
        FormatSnapshot.Create(label, CycleKind.NumberFormat, workbook, sheet, new[]
        {
            new SnapshotBlock(new CellRect(1, 1, 1, 1), CycleValue.FromNumberFormat("General"), CycleValue.FromNumberFormat("0.0")),
        });

    [Theory]
    [InlineData(UndoKey.Undo, false, true, UndoDecision.PassToExcel)]
    [InlineData(UndoKey.Undo, false, false, UndoDecision.PassToExcel)]
    [InlineData(UndoKey.Undo, false, null, UndoDecision.PassToExcel)]
    [InlineData(UndoKey.Undo, true, true, UndoDecision.PassToExcel)]
    [InlineData(UndoKey.Undo, true, null, UndoDecision.PassToExcel)]
    [InlineData(UndoKey.Undo, true, false, UndoDecision.HandleOurs)]
    [InlineData(UndoKey.Redo, false, true, UndoDecision.PassToExcel)]
    [InlineData(UndoKey.Redo, false, false, UndoDecision.PassToExcel)]
    [InlineData(UndoKey.Redo, false, null, UndoDecision.PassToExcel)]
    [InlineData(UndoKey.Redo, true, true, UndoDecision.PassToExcel)]
    [InlineData(UndoKey.Redo, true, null, UndoDecision.PassToExcel)]
    [InlineData(UndoKey.Redo, true, false, UndoDecision.HandleOurs)]
    public void Decide_handles_ours_only_when_we_have_one_and_excel_has_none(
        UndoKey key, bool ourStackNonEmpty, bool? nativeAvailable, UndoDecision expected)
    {
        Assert.Equal(expected, UndoManager.Decide(key, ourStackNonEmpty, nativeAvailable));
    }

    [Fact]
    public void Decide_rejects_an_undefined_key()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => UndoManager.Decide((UndoKey)7, true, false));
    }

    [Fact]
    public void New_manager_is_empty()
    {
        var manager = new UndoManager();

        Assert.Equal(UndoManager.DefaultMaxDepth, manager.MaxDepth);
        Assert.Equal(100, manager.MaxDepth);
        Assert.Equal(0, manager.UndoCount);
        Assert.Equal(0, manager.RedoCount);
        Assert.Null(manager.Peek(UndoKey.Undo));
        Assert.Null(manager.Peek(UndoKey.Redo));
        Assert.Null(manager.Undo());
        Assert.Null(manager.Redo());
        Assert.Null(manager.Discard(UndoKey.Undo));
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
    public void Push_rejects_null_and_unavailable_snapshots()
    {
        var manager = new UndoManager();

        Assert.Throws<ArgumentNullException>(() => manager.Push(null!));
        var error = Assert.Throws<ArgumentException>(() => manager.Push(
            FormatSnapshot.Unavailable("Number", CycleKind.NumberFormat, "Book1.xlsx", "Sheet1", "too many formats")));
        Assert.Contains("too many formats", error.Message);
        Assert.Equal(0, manager.UndoCount);
    }

    [Fact]
    public void Constructor_rejects_a_depth_below_one()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new UndoManager(0));
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

    /// <summary>
    /// Excel's native undo and redo as counts, per spike K2c: every COM write (our cycle, and our restore) wipes
    /// both; typing adds an undo entry and clears redo. Keys go through <see cref="UndoManager.Decide"/>.
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

        public string Press(UndoKey key)
        {
            var native = Unknown ? (bool?)null : (key == UndoKey.Undo ? _nativeUndo : _nativeRedo) > 0;
            var decision = UndoManager.Decide(key, Manager.Count(key) > 0, native);
            if (decision == UndoDecision.HandleOurs)
            {
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
