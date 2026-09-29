using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelModelingToolkit.Core.Formatting;
using ExcelModelingToolkit.Core.Keys;
using ExcelModelingToolkit.Core.Settings;

namespace ExcelModelingToolkit.AddIn;

/// <summary>
/// The settings dialog (docs/PLAN.md section 4.6): a modal WinForms window owned by Excel's active window that
/// edits a <see cref="SettingsDraft"/>, where all the editing logic lives. Cycles tab: items added, removed, moved,
/// renamed and given a code or color, with a live preview; Shortcuts tab: keys typed or captured from the keyboard,
/// with inline problems; General tab: the undo cap and the diagnostics log. Import, Export, Reset to defaults, OK
/// (validate, save settings.json atomically) and Cancel (discard).
/// </summary>
/// <remarks>
/// WinForms, not WPF: WPF keyboard focus is unreliable on Excel's thread (spike K4). The dialog runs on Excel's
/// main thread inside a macro (<see cref="Commands.EmtSettings"/>), so the number format preview calls Excel
/// synchronously. While it is open, <see cref="Form.ShowDialog(IWin32Window)"/> disables every Excel window on the
/// thread. The Ctrl+Z hook passes keys through here, since the focus is not on a worksheet grid.
/// </remarks>
internal sealed class SettingsDialog : Form
{
    private const int MaxProblemsShown = 12;

    /// <summary>Values shown in the number format preview, as Excel's TEXT function renders them.</summary>
    private static readonly object[] PreviewSamples = { 1234.5, -1234.5, 0.0, "Text" };

    private readonly SettingsDraft _draft;
    private readonly Func<object, string, string> _renderSample;

    // Cycles tab.
    private readonly ListBox _cycleList = new ListBox { Dock = DockStyle.Fill, IntegralHeight = false };
    private readonly ListView _itemList = CreateList();
    private readonly ColumnHeader _itemNameColumn = new ColumnHeader { Text = "Name" };
    private readonly ColumnHeader _itemValueColumn = new ColumnHeader { Text = "Code" };
    private readonly ImageList _swatches = new ImageList { ColorDepth = ColorDepth.Depth32Bit };
    private readonly Button _addButton = CreateButton("&Add");
    private readonly Button _removeButton = CreateButton("&Remove");
    private readonly Button _upButton = CreateButton("Move &up");
    private readonly Button _downButton = CreateButton("Move do&wn");
    private readonly Label _provisionalNote = CreateNote();
    private readonly TextBox _nameBox = CreateTextBox();
    private readonly Label _codeLabel = CreateLabel("Format c&ode:");
    private readonly TextBox _codeBox = CreateTextBox();
    private readonly Label _colorLabel = CreateLabel("Co&lor:");
    private readonly TextBox _colorBox = CreateTextBox();
    private readonly Button _pickColorButton = CreateButton("&Pick color...");
    private readonly Label _swatch = new Label { AutoSize = false, BorderStyle = BorderStyle.FixedSingle, TextAlign = ContentAlignment.MiddleCenter, Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Bottom };
    private readonly CheckBox _noFillBox = new CheckBox { Text = "No &fill", AutoSize = true, Anchor = AnchorStyles.Left };
    private readonly Label _colorError = CreateError();
    private readonly GroupBox _previewGroup = new GroupBox { Text = "Preview (rendered by Excel)", Dock = DockStyle.Fill, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
    private readonly Font _previewFont;
    private readonly Label[] _previewResults = PreviewSamples.Select(_ => new Label { AutoSize = true, Anchor = AnchorStyles.Left, UseMnemonic = false }).ToArray();
    private readonly Control[] _itemEditors;

    // Shortcuts tab.
    private readonly ListView _actionList = CreateList();
    private readonly TextBox _keyBox = CreateTextBox();
    private readonly Button _captureButton = CreateButton("&Capture...");
    private readonly Button _clearKeyButton = CreateButton("C&lear");
    private readonly Label _keyError = CreateError();
    private bool _capturing;

    // General tab.
    private readonly NumericUpDown _undoCapBox = new NumericUpDown { Minimum = 1, Maximum = int.MaxValue, ThousandsSeparator = true, Anchor = AnchorStyles.Left };
    private readonly CheckBox _logBox = new CheckBox { Text = "Write a dia&gnostics log", AutoSize = true, Anchor = AnchorStyles.Left };

    // Bottom bar.
    private readonly Label _status = new Label { AutoSize = false, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, AutoEllipsis = true, UseMnemonic = false };

    // True while controls are being filled from the draft, so their change events do not edit it.
    private bool _loading;

    /// <summary>Creates the dialog; <paramref name="renderSample"/> formats a preview value with a number format code.</summary>
    private SettingsDialog(SettingsDraft draft, Func<object, string, string> renderSample)
    {
        _draft = draft;
        _renderSample = renderSample;
        _itemEditors = new Control[] { _nameBox, _codeBox, _colorBox, _pickColorButton, _noFillBox };
        _previewFont = new Font(FontFamily.GenericMonospace, SystemFonts.MessageBoxFont.SizeInPoints);

        Text = ProductInfo.Name + " Settings";
        Font = SystemFonts.MessageBoxFont;
        AutoScaleMode = AutoScaleMode.Dpi;
        AutoScaleDimensions = new SizeF(96F, 96F);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.Sizable;
        MinimizeBox = false;
        ShowInTaskbar = false;
        ShowIcon = false;
        ClientSize = new Size(820, 600);
        MinimumSize = new Size(680, 520);

        var tabs = new TabControl { Dock = DockStyle.Fill };
        tabs.TabPages.Add(CreateCyclesPage());
        tabs.TabPages.Add(CreateShortcutsPage());
        tabs.TabPages.Add(CreateGeneralPage());

        var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(8), ColumnCount = 1, RowCount = 2 };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.Controls.Add(tabs, 0, 0);
        root.Controls.Add(CreateButtonBar(), 0, 1);
        Controls.Add(root);

        LoadFromDraft();
    }

    /// <summary>The settings saved by OK, or null if the dialog was cancelled.</summary>
    private ToolkitSettings? SavedSettings { get; set; }

    /// <summary>
    /// Shows the dialog, modal and owned by Excel's active window, editing a copy of <paramref name="current"/>.
    /// Returns the settings saved to settings.json by OK, or null if the dialog was cancelled. Call in macro context
    /// on Excel's main thread.
    /// </summary>
    public static ToolkitSettings? Edit(ToolkitSettings current)
    {
        // Showing a form would otherwise install a WinForms SynchronizationContext on Excel's main thread.
        WindowsFormsSynchronizationContext.AutoInstall = false;
        Application.EnableVisualStyles();

        using var dialog = new SettingsDialog(new SettingsDraft(current), RenderSample);
        return dialog.ShowDialog(new ExcelWindow()) == DialogResult.OK ? dialog.SavedSettings : null;
    }

    /// <summary>Capture mode takes the next key press, before any control or mnemonic sees it.</summary>
    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (!_capturing)
        {
            return base.ProcessCmdKey(ref msg, keyData);
        }

        var keyCode = (int)(keyData & Keys.KeyCode);
        if ((keyData & Keys.KeyCode) == Keys.Escape && (keyData & Keys.Modifiers) == Keys.None)
        {
            StopCapture("Capture cancelled.");
            return true;
        }

        if (KeyCapture.IsModifierKey(keyCode))
        {
            return true; // Wait for the key that goes with the modifiers.
        }

        var modifiers = KeyModifiers.None;
        if ((keyData & Keys.Control) != 0)
        {
            modifiers |= KeyModifiers.Ctrl;
        }

        if ((keyData & Keys.Alt) != 0)
        {
            modifiers |= KeyModifiers.Alt;
        }

        if ((keyData & Keys.Shift) != 0)
        {
            modifiers |= KeyModifiers.Shift;
        }

        var text = KeyCapture.ToKeyString(keyCode, modifiers);
        if (text is null)
        {
            StopCapture(
                "That key can't be part of a shortcut. Use a letter, digit, punctuation key, F1-F12, an arrow, " +
                "PgUp, PgDn, Home, End, Ins or Del, with Ctrl, Alt and/or Shift.");
            return true;
        }

        StopCapture(null);
        _keyBox.Text = text; // Updates the draft and shows any problem.
        return true;
    }

    /// <inheritdoc />
    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _swatches.Dispose();
            _previewFont.Dispose();
        }

        base.Dispose(disposing);
    }

    private static string ColorText(OleColor color) =>
        color.IsNoFill
            ? "No fill"
            : string.Format(CultureInfo.InvariantCulture, "{0}  rgb({1}, {2}, {3})", color, color.R, color.G, color.B);

    private static string CycleLabel(CycleDraft cycle) => cycle.Provisional ? cycle.DisplayName + " (provisional)" : cycle.DisplayName;

    /// <summary>
    /// <paramref name="value"/> formatted with <paramref name="code"/> by Excel's own formatter,
    /// <c>WorksheetFunction.Text</c> (late-bound COM), or a note if Excel rejects the code. TEXT reads the code in
    /// Excel's display language, so outside en-US the preview can differ from the cell, which uses en-US syntax.
    /// </summary>
    private static string RenderSample(object value, string code)
    {
        if (code.Trim().Length == 0)
        {
            return "(no format code)";
        }

        try
        {
            dynamic application = ExcelDnaUtil.Application;
            object result = application.WorksheetFunction.Text(value, code);
            return Convert.ToString(result, CultureInfo.CurrentCulture) ?? string.Empty;
        }
        catch (Exception)
        {
            return "(invalid format)";
        }
    }

    private static string DescribeSample(object value) =>
        value is string text ? "\"" + text + "\"" : Convert.ToString(value, CultureInfo.InvariantCulture)!;

    private static ListView CreateList() => new ListView
    {
        Dock = DockStyle.Fill,
        View = View.Details,
        FullRowSelect = true,
        HideSelection = false,
        MultiSelect = false,
        HeaderStyle = ColumnHeaderStyle.Nonclickable,
    };

    private static Button CreateButton(string text) => new Button { Text = text, AutoSize = true, MinimumSize = new Size(88, 0), Anchor = AnchorStyles.Left | AnchorStyles.Right };

    private static Label CreateLabel(string text) => new Label { Text = text, AutoSize = true, Anchor = AnchorStyles.Left };

    private static TextBox CreateTextBox() => new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right };

    private static Label CreateNote() => new Label { AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Right, ForeColor = SystemColors.GrayText, UseMnemonic = false, Margin = new Padding(3, 6, 3, 3) };

    private static Label CreateError() => new Label { AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Right, ForeColor = Color.Firebrick, UseMnemonic = false };

    private static TableLayoutPanel CreateTable(int columns) => new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = columns, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };

    private TabPage CreateCyclesPage()
    {
        var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = Padding.Empty };
        buttons.Controls.AddRange(new Control[] { _addButton, _removeButton, _upButton, _downButton });
        foreach (Button button in buttons.Controls)
        {
            button.Anchor = AnchorStyles.Left | AnchorStyles.Right;
        }

        _itemList.Columns.AddRange(new[] { _itemNameColumn, _itemValueColumn });
        _swatches.ImageSize = new Size(LogicalToDeviceUnits(16), LogicalToDeviceUnits(16));

        var itemsRow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = Padding.Empty };
        itemsRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        itemsRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        itemsRow.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        itemsRow.Controls.Add(_itemList, 0, 0);
        itemsRow.Controls.Add(buttons, 1, 0);

        // The selected item's fields. Rows that do not apply to the cycle's kind are hidden (and take no space).
        _swatch.Size = new Size(72, 20); // Logical units: the form scales them; the anchors give it the row's height.
        var editor = CreateTable(4);
        editor.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        editor.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        editor.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        editor.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        editor.Controls.Add(CreateLabel("&Name:"), 0, 0);
        editor.Controls.Add(_nameBox, 1, 0);
        editor.SetColumnSpan(_nameBox, 3);
        editor.Controls.Add(_codeLabel, 0, 1);
        editor.Controls.Add(_codeBox, 1, 1);
        editor.SetColumnSpan(_codeBox, 3);
        editor.Controls.Add(_colorLabel, 0, 2);
        editor.Controls.Add(_colorBox, 1, 2);
        editor.Controls.Add(_swatch, 2, 2);
        editor.Controls.Add(_pickColorButton, 3, 2);
        editor.Controls.Add(_noFillBox, 1, 3);
        editor.Controls.Add(_colorError, 1, 4);
        editor.SetColumnSpan(_colorError, 3);

        var preview = CreateTable(2);
        preview.Dock = DockStyle.Top; // A docked-fill child would give the auto-sized group no height.
        for (var i = 0; i < PreviewSamples.Length; i++)
        {
            preview.Controls.Add(new Label { Text = DescribeSample(PreviewSamples[i]), AutoSize = true, Anchor = AnchorStyles.Left, UseMnemonic = false }, 0, i);
            _previewResults[i].Font = _previewFont;
            preview.Controls.Add(_previewResults[i], 1, i);
        }

        _previewGroup.Controls.Add(preview);

        var right = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5, Margin = Padding.Empty };
        right.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        right.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        right.Controls.Add(CreateLabel("I&tems (applied in this order):"), 0, 0);
        right.Controls.Add(itemsRow, 0, 1);
        right.Controls.Add(_provisionalNote, 0, 2);
        right.Controls.Add(editor, 0, 3);
        right.Controls.Add(_previewGroup, 0, 4);

        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2, Padding = new Padding(4) };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 28));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 72));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.Controls.Add(CreateLabel("&Cycles:"), 0, 0);
        layout.Controls.Add(_cycleList, 0, 1);
        layout.Controls.Add(right, 1, 0);
        layout.SetRowSpan(right, 2);

        _cycleList.SelectedIndexChanged += (s, e) => ShowCycle(0);
        _itemList.SelectedIndexChanged += (s, e) => ShowItem();
        _itemList.Resize += (s, e) => _itemValueColumn.Width = -2; // The last column fills the list.
        _addButton.Click += (s, e) => AddItem();
        _removeButton.Click += (s, e) => EditItems(c =>
        {
            var index = SelectedItem;
            c.RemoveItem(index);
            return Math.Min(index, c.Items.Count - 1);
        });
        _upButton.Click += (s, e) => EditItems(c => c.MoveItemUp(SelectedItem));
        _downButton.Click += (s, e) => EditItems(c => c.MoveItemDown(SelectedItem));
        _nameBox.TextChanged += (s, e) => EditSelectedItem(c => c.RenameItem(SelectedItem, _nameBox.Text));
        _codeBox.TextChanged += (s, e) => EditSelectedItem(c => c.SetCode(SelectedItem, _codeBox.Text));
        _colorBox.TextChanged += (s, e) => ColorTyped();
        _colorBox.Leave += (s, e) => ShowItem(); // Puts back the item's color if the text was left invalid.
        _noFillBox.CheckedChanged += (s, e) => NoFillToggled();
        _pickColorButton.Click += (s, e) => PickColor();

        var page = new TabPage("Cycles") { UseVisualStyleBackColor = true };
        page.Controls.Add(layout);
        return page;
    }

    private TabPage CreateShortcutsPage()
    {
        _actionList.Columns.Add("Action", LogicalToDeviceUnits(140));
        _actionList.Columns.Add("Shortcut", LogicalToDeviceUnits(150));
        _actionList.Columns.Add("Problem", LogicalToDeviceUnits(400));

        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 5, Padding = new Padding(4) };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var listLabel = CreateLabel("&Actions:");
        layout.Controls.Add(listLabel, 0, 0);
        layout.SetColumnSpan(listLabel, 4);
        layout.Controls.Add(_actionList, 0, 1);
        layout.SetColumnSpan(_actionList, 4);
        layout.Controls.Add(CreateLabel("&Shortcut:"), 0, 2);
        layout.Controls.Add(_keyBox, 1, 2);
        layout.Controls.Add(_captureButton, 2, 2);
        layout.Controls.Add(_clearKeyButton, 3, 2);
        layout.Controls.Add(_keyError, 1, 3);
        layout.SetColumnSpan(_keyError, 3);
        var note = CreateNote();
        note.Text =
            "Type a shortcut such as Ctrl+Shift+K, or click Capture and press it. Use Ctrl and/or Alt, with or " +
            "without Shift. Some keys override Excel's built-ins while the add-in is loaded (e.g. Ctrl+; inserts " +
            "the date). Changes take effect when you click OK.";
        layout.Controls.Add(note, 0, 4);
        layout.SetColumnSpan(note, 4);

        _actionList.SelectedIndexChanged += (s, e) => ShowAction();
        _keyBox.TextChanged += (s, e) => KeyTyped();
        _keyBox.Leave += (s, e) => ShowAction(); // Shows the key in its standard form.
        _captureButton.Click += (s, e) => StartCapture();
        _captureButton.Leave += (s, e) => StopCapture(null);
        _clearKeyButton.Click += (s, e) =>
        {
            _keyBox.Text = string.Empty;
            _keyBox.Focus();
        };

        var page = new TabPage("Shortcuts") { UseVisualStyleBackColor = true };
        page.Controls.Add(layout);
        return page;
    }

    private TabPage CreateGeneralPage()
    {
        var openLog = CreateButton("Open log f&older");
        var openSettings = CreateButton("Open settings &file");
        openLog.Anchor = AnchorStyles.Left;
        openSettings.Anchor = AnchorStyles.Left;

        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 7, Padding = new Padding(8) };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        for (var row = 0; row < 6; row++)
        {
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        }

        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // Spare space goes below the options.
        layout.Controls.Add(CreateLabel("&Undo capture limit (format reads per change):"), 0, 0);
        layout.Controls.Add(_undoCapBox, 1, 0);
        AddNote(layout, 1, "A change that needs more reads than this is applied, but Ctrl+Z cannot undo it. A block of cells with the same format costs one read.");
        layout.Controls.Add(_logBox, 0, 2);
        layout.SetColumnSpan(_logBox, 2);
        AddNote(layout, 3, "Records each command and its timing in " + Path.Combine(DiagnosticsLog.FolderPath, "log.txt") + ".");
        var buttons = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = new Padding(0, 12, 0, 0) };
        buttons.Controls.AddRange(new Control[] { openLog, openSettings });
        layout.Controls.Add(buttons, 0, 4);
        layout.SetColumnSpan(buttons, 2);
        AddNote(layout, 5, "Settings file: " + SettingsStore.FilePath);

        _undoCapBox.ValueChanged += (s, e) =>
        {
            if (!_loading)
            {
                _draft.UndoCellCap = (int)_undoCapBox.Value;
            }
        };
        _logBox.CheckedChanged += (s, e) =>
        {
            if (!_loading)
            {
                _draft.DiagnosticsLog = _logBox.Checked;
            }
        };
        openLog.Click += (s, e) => OpenLogFolder();
        openSettings.Click += (s, e) =>
        {
            var problem = Commands.OpenSettingsFile();
            if (problem is not null)
            {
                MessageBox.Show(this, problem, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        };

        var page = new TabPage("General") { UseVisualStyleBackColor = true };
        page.Controls.Add(layout);
        return page;
    }

    private static void AddNote(TableLayoutPanel layout, int row, string text)
    {
        var note = CreateNote();
        note.Text = text;
        note.Margin = new Padding(3, 0, 3, 8);
        layout.Controls.Add(note, 0, row);
        layout.SetColumnSpan(note, 2);
    }

    private Control CreateButtonBar()
    {
        var import = CreateButton("&Import...");
        var export = CreateButton("&Export...");
        var reset = CreateButton("Reset to &defaults");
        var ok = CreateButton("OK");
        var cancel = CreateButton("Cancel");
        cancel.DialogResult = DialogResult.Cancel;
        AcceptButton = ok;
        CancelButton = cancel;

        var bar = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 6, RowCount = 1, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = new Padding(0, 8, 0, 0) };
        bar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        bar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bar.Controls.Add(import, 0, 0);
        bar.Controls.Add(export, 1, 0);
        bar.Controls.Add(reset, 2, 0);
        bar.Controls.Add(_status, 3, 0);
        bar.Controls.Add(ok, 4, 0);
        bar.Controls.Add(cancel, 5, 0);

        import.Click += (s, e) => Import();
        export.Click += (s, e) => Export();
        reset.Click += (s, e) => ResetToDefaults();
        ok.Click += (s, e) => Save();
        return bar;
    }

    // ---- Filling the controls from the draft ----

    private CycleDraft? SelectedCycle =>
        _cycleList.SelectedIndex >= 0 && _cycleList.SelectedIndex < _draft.Cycles.Count ? _draft.Cycles[_cycleList.SelectedIndex] : null;

    private int SelectedItem => _itemList.SelectedIndices.Count > 0 ? _itemList.SelectedIndices[0] : -1;

    private string? SelectedAction =>
        _actionList.SelectedItems.Count > 0 ? (string)_actionList.SelectedItems[0].Tag : null;

    /// <summary>Fills every control from the draft (on open, and after an import or a reset).</summary>
    private void LoadFromDraft()
    {
        var cycleIndex = Math.Max(0, _cycleList.SelectedIndex);
        var actionIndex = _actionList.SelectedIndices.Count > 0 ? _actionList.SelectedIndices[0] : 0;

        _loading = true;
        try
        {
            _cycleList.BeginUpdate();
            _cycleList.Items.Clear();
            foreach (var cycle in _draft.Cycles)
            {
                _cycleList.Items.Add(CycleLabel(cycle));
            }

            _cycleList.EndUpdate();
            _cycleList.SelectedIndex = _draft.Cycles.Count == 0 ? -1 : Math.Min(cycleIndex, _draft.Cycles.Count - 1);

            _actionList.BeginUpdate();
            _actionList.Items.Clear();
            foreach (var actionId in _draft.ShortcutActions)
            {
                _actionList.Items.Add(new ListViewItem(new[] { _draft.ActionDisplayName(actionId), string.Empty, string.Empty }) { Tag = actionId });
            }

            _actionList.EndUpdate();
            RefreshActionRows();
            if (_actionList.Items.Count > 0)
            {
                var item = _actionList.Items[Math.Min(actionIndex, _actionList.Items.Count - 1)];
                item.Selected = true;
                item.Focused = true;
            }

            _undoCapBox.Value = Math.Max(1, _draft.UndoCellCap);
            _logBox.Checked = _draft.DiagnosticsLog;
        }
        finally
        {
            _loading = false;
        }

        ShowCycle(0);
        ShowAction();
    }

    /// <summary>Shows the selected cycle's items and selects the item at <paramref name="itemIndex"/>.</summary>
    private void ShowCycle(int itemIndex)
    {
        if (_loading)
        {
            return;
        }

        var cycle = SelectedCycle;
        var isFormat = cycle?.Kind == CycleKind.NumberFormat;
        _itemValueColumn.Text = isFormat ? "Code" : "Color";
        _itemList.SmallImageList = isFormat ? null : _swatches;
        _codeLabel.Visible = _codeBox.Visible = isFormat;
        _previewGroup.Text = isFormat ? "Preview (rendered by Excel's TEXT function)" : "Preview";
        _colorLabel.Visible = _colorBox.Visible = _swatch.Visible = _pickColorButton.Visible = cycle is not null && !isFormat;
        _noFillBox.Visible = cycle?.AllowsNoFill == true;
        _previewGroup.Visible = isFormat;
        _provisionalNote.Visible = cycle?.Provisional == true;
        _provisionalNote.Text = cycle is null
            ? string.Empty
            : $"The {cycle.DisplayName} list is a provisional placeholder. Editing its items makes it your own list.";
        FillItems(itemIndex);
    }

    private void FillItems(int selectIndex)
    {
        var cycle = SelectedCycle;
        _loading = true;
        try
        {
            _itemList.BeginUpdate();
            _itemList.Items.Clear();
            _swatches.Images.Clear();
            if (cycle is not null)
            {
                for (var i = 0; i < cycle.Items.Count; i++)
                {
                    _itemList.Items.Add(new ListViewItem(new[] { string.Empty, string.Empty }));
                    UpdateItemRow(i);
                }
            }

            _itemList.EndUpdate();
            _itemNameColumn.Width = LogicalToDeviceUnits(200);
            _itemValueColumn.Width = -2; // Fill the rest of the list.
            if (_itemList.Items.Count > 0)
            {
                var item = _itemList.Items[Math.Max(0, Math.Min(selectIndex, _itemList.Items.Count - 1))];
                item.Selected = true;
                item.Focused = true;
                item.EnsureVisible();
            }
        }
        finally
        {
            _loading = false;
        }

        ShowItem();
    }

    /// <summary>Updates the list row of item <paramref name="index"/> of the selected cycle (text and swatch).</summary>
    private void UpdateItemRow(int index)
    {
        var cycle = SelectedCycle;
        if (cycle is null || index < 0 || index >= _itemList.Items.Count)
        {
            return;
        }

        var row = _itemList.Items[index];
        var item = cycle.Items[index];
        row.Text = item.Name;
        if (item is ColorItem color)
        {
            row.SubItems[1].Text = ColorText(color.Color);
            var key = color.Color.ToString();
            if (!_swatches.Images.ContainsKey(key))
            {
                _swatches.Images.Add(key, CreateSwatch(color.Color));
            }

            row.ImageKey = key;
        }
        else
        {
            row.SubItems[1].Text = ((NumberFormatItem)item).Code;
        }
    }

    private Bitmap CreateSwatch(OleColor color)
    {
        var size = _swatches.ImageSize;
        var bitmap = new Bitmap(size.Width, size.Height);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.Transparent);
        var box = new Rectangle(1, 1, size.Width - 3, size.Height - 3);
        if (color.IsNoFill)
        {
            graphics.FillRectangle(SystemBrushes.Window, box);
            graphics.DrawLine(Pens.Firebrick, box.Left, box.Bottom, box.Right, box.Top);
        }
        else
        {
            using var brush = new SolidBrush(Color.FromArgb(color.R, color.G, color.B));
            graphics.FillRectangle(brush, box);
        }

        graphics.DrawRectangle(SystemPens.ControlDarkDark, box);
        return bitmap;
    }

    /// <summary>Fills the item fields and the preview from the selected item, and enables the buttons that apply.</summary>
    private void ShowItem()
    {
        if (_loading)
        {
            return;
        }

        var cycle = SelectedCycle;
        var index = SelectedItem;
        var item = cycle is not null && index >= 0 ? cycle.Items[index] : null;
        _addButton.Enabled = cycle is not null;
        _removeButton.Enabled = item is not null;
        _upButton.Enabled = item is not null && index > 0;
        _downButton.Enabled = item is not null && index < cycle!.Items.Count - 1;
        foreach (var editor in _itemEditors)
        {
            editor.Enabled = item is not null;
        }

        _loading = true;
        try
        {
            _nameBox.Text = item?.Name ?? string.Empty;
            _codeBox.Text = (item as NumberFormatItem)?.Code ?? string.Empty;
            var color = (item as ColorItem)?.Color;
            _noFillBox.Checked = color?.IsNoFill == true;
            _colorBox.Text = color is null || color.Value.IsNoFill ? string.Empty : color.Value.ToString();
            _colorBox.Enabled = _pickColorButton.Enabled = item is not null && !_noFillBox.Checked;
            _colorError.Text = string.Empty;
        }
        finally
        {
            _loading = false;
        }

        UpdatePreview();
    }

    private void UpdatePreview()
    {
        var item = SelectedCycle is { } cycle && SelectedItem >= 0 ? cycle.Items[SelectedItem] : null;
        if (item is NumberFormatItem format)
        {
            for (var i = 0; i < PreviewSamples.Length; i++)
            {
                _previewResults[i].Text = _renderSample(PreviewSamples[i], format.Code);
            }
        }
        else
        {
            foreach (var result in _previewResults)
            {
                result.Text = string.Empty;
            }
        }

        var color = (item as ColorItem)?.Color;
        _swatch.BackColor = color is null || color.Value.IsNoFill ? SystemColors.Window : Color.FromArgb(color.Value.R, color.Value.G, color.Value.B);
        _swatch.Text = color?.IsNoFill == true ? "No fill" : string.Empty;
    }

    private void RefreshCycleLabel()
    {
        var index = _cycleList.SelectedIndex;
        var cycle = SelectedCycle;
        if (cycle is null || Equals(_cycleList.Items[index], CycleLabel(cycle)))
        {
            return;
        }

        _loading = true;
        try
        {
            _cycleList.Items[index] = CycleLabel(cycle);
            _cycleList.SelectedIndex = index;
        }
        finally
        {
            _loading = false;
        }

        _provisionalNote.Visible = cycle.Provisional;
    }

    // ---- Editing cycles ----

    private void AddItem()
    {
        var cycle = SelectedCycle;
        if (cycle is null)
        {
            return;
        }

        var index = cycle.AddItem(SelectedItem >= 0 ? SelectedItem + 1 : cycle.Items.Count);
        RefreshCycleLabel();
        FillItems(index);
        _nameBox.Focus();
        _nameBox.SelectAll();
    }

    /// <summary>Applies a list edit (remove or move) and selects the item index it returns.</summary>
    private void EditItems(Func<CycleDraft, int> edit)
    {
        var cycle = SelectedCycle;
        if (cycle is null || SelectedItem < 0)
        {
            return;
        }

        var select = edit(cycle);
        RefreshCycleLabel();
        FillItems(select);
        if (_itemList.Items.Count == 0)
        {
            _addButton.Focus();
        }
    }

    /// <summary>Applies a field edit to the selected item and updates its row and the preview.</summary>
    private void EditSelectedItem(Action<CycleDraft> edit)
    {
        var cycle = SelectedCycle;
        if (_loading || cycle is null || SelectedItem < 0)
        {
            return;
        }

        edit(cycle);
        UpdateItemRow(SelectedItem);
        RefreshCycleLabel();
        UpdatePreview();
    }

    private void ColorTyped()
    {
        var cycle = SelectedCycle;
        if (_loading || cycle is null || SelectedItem < 0)
        {
            return;
        }

        OleColor color;
        try
        {
            color = OleColor.Parse(_colorBox.Text);
        }
        catch (FormatException)
        {
            _colorError.Text = "Enter a color as #RRGGBB or rgb(r, g, b), e.g. #1C4587 or rgb(28, 69, 135).";
            return;
        }

        if (color.IsNoFill)
        {
            _colorError.Text = cycle.AllowsNoFill ? "Tick \"No fill\" instead." : "\"No fill\" is only for fill colors.";
            return;
        }

        _colorError.Text = string.Empty;
        EditSelectedItem(c => c.SetColor(SelectedItem, color));
    }

    private void NoFillToggled()
    {
        var cycle = SelectedCycle;
        if (_loading || cycle is null || SelectedItem < 0)
        {
            return;
        }

        OleColor color;
        if (_noFillBox.Checked)
        {
            color = OleColor.NoFill;
        }
        else
        {
            try
            {
                color = OleColor.Parse(_colorBox.Text);
            }
            catch (FormatException)
            {
                color = OleColor.FromRgb(255, 255, 255);
            }
        }

        EditSelectedItem(c => c.SetColor(SelectedItem, color));
        ShowItem();
    }

    private void PickColor()
    {
        var cycle = SelectedCycle;
        if (cycle is null || SelectedItem < 0 || !(cycle.Items[SelectedItem] is ColorItem item))
        {
            return;
        }

        using var picker = new ColorDialog { FullOpen = true, AnyColor = true };
        if (!item.Color.IsNoFill)
        {
            picker.Color = Color.FromArgb(item.Color.R, item.Color.G, item.Color.B);
        }

        if (picker.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        var picked = OleColor.FromRgb(picker.Color.R, picker.Color.G, picker.Color.B);
        EditSelectedItem(c => c.SetColor(SelectedItem, picked));
        ShowItem();
    }

    // ---- Shortcuts ----

    private void ShowAction()
    {
        if (_loading)
        {
            return;
        }

        var actionId = SelectedAction;
        _keyBox.Enabled = _captureButton.Enabled = _clearKeyButton.Enabled = actionId is not null;
        _loading = true;
        try
        {
            _keyBox.Text = actionId is null ? string.Empty : _draft.GetKey(actionId);
        }
        finally
        {
            _loading = false;
        }

        _keyError.Text = actionId is null ? string.Empty : _draft.KeyProblem(actionId) ?? string.Empty;
    }

    private void KeyTyped()
    {
        var actionId = SelectedAction;
        if (_loading || actionId is null)
        {
            return;
        }

        _draft.SetKey(actionId, _keyBox.Text);
        RefreshActionRows();
        _keyError.Text = _draft.KeyProblem(actionId) ?? string.Empty;
    }

    /// <summary>Updates every row's key and problem: changing one key can create or clear a duplicate on another.</summary>
    private void RefreshActionRows()
    {
        foreach (ListViewItem row in _actionList.Items)
        {
            var actionId = (string)row.Tag;
            var key = _draft.GetKey(actionId);
            row.SubItems[1].Text = key.Length == 0 ? "(none)" : key;
            row.SubItems[2].Text = _draft.KeyProblem(actionId) ?? string.Empty;
        }
    }

    private void StartCapture()
    {
        if (SelectedAction is null)
        {
            return;
        }

        _capturing = true;
        _captureButton.Text = "Press keys...";
        _keyError.Text = "Press the shortcut now, or Esc to cancel.";
        _captureButton.Focus();
    }

    private void StopCapture(string? message)
    {
        if (!_capturing)
        {
            return;
        }

        _capturing = false;
        _captureButton.Text = "&Capture...";
        _keyError.Text = message ?? string.Empty;
    }

    // ---- General ----

    private void OpenLogFolder()
    {
        try
        {
            Directory.CreateDirectory(DiagnosticsLog.FolderPath);
            Process.Start(new ProcessStartInfo(DiagnosticsLog.FolderPath) { UseShellExecute = true })?.Dispose();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Could not open {DiagnosticsLog.FolderPath}: {ex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    // ---- Import, export, reset, OK ----

    private void Import()
    {
        using var open = new OpenFileDialog
        {
            Title = "Import settings",
            Filter = "Settings files (*.json)|*.json|All files (*.*)|*.*",
            CheckFileExists = true,
        };
        if (open.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        IReadOnlyList<string> problems;
        try
        {
            var tooLarge = ToolkitSettings.CheckFileSize(new FileInfo(open.FileName).Length);
            problems = tooLarge is not null ? new[] { tooLarge } : _draft.ImportFile(File.ReadAllBytes(open.FileName));
        }
        catch (Exception ex)
        {
            problems = new[] { ex.Message };
        }

        if (problems.Count > 0)
        {
            ShowProblems($"Could not import {Path.GetFileName(open.FileName)}. Your settings are unchanged.", problems);
            return;
        }

        LoadFromDraft();
        _status.Text = $"Imported {Path.GetFileName(open.FileName)}. Click OK to save and apply.";
    }

    private void Export()
    {
        CommitPendingEdits();
        var json = _draft.ExportJson(out var problems);
        if (json is null)
        {
            ShowProblems("Fix these problems before exporting:", problems);
            return;
        }

        using var save = new SaveFileDialog
        {
            Title = "Export settings",
            Filter = "Settings files (*.json)|*.json|All files (*.*)|*.*",
            DefaultExt = "json",
            FileName = ProductInfo.FolderName + "-settings.json",
            OverwritePrompt = true,
        };
        if (save.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        try
        {
            File.WriteAllText(save.FileName, json, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            _status.Text = $"Exported to {Path.GetFileName(save.FileName)}.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Could not export to {save.FileName}: {ex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private void ResetToDefaults()
    {
        var answer = MessageBox.Show(
            this,
            "Reset every cycle, shortcut and option to the factory defaults?\n\nNothing is saved until you click OK.",
            Text,
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question,
            MessageBoxDefaultButton.Button2);
        if (answer != DialogResult.Yes)
        {
            return;
        }

        _draft.ResetToDefaults();
        LoadFromDraft();
        _status.Text = "Defaults restored. Click OK to save and apply, or Cancel to keep your settings.";
    }

    private void Save()
    {
        CommitPendingEdits();
        var settings = _draft.ToSettings(out var problems);
        if (problems.Count > 0)
        {
            ShowProblems("The settings can't be saved yet. Fix these problems, or click Cancel:", problems);
            return;
        }

        var saveProblem = SettingsStore.Save(settings);
        if (saveProblem is not null)
        {
            MessageBox.Show(this, $"The settings were not saved: {saveProblem}", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        SavedSettings = settings;
        DialogResult = DialogResult.OK;
    }

    /// <summary>
    /// Takes a number typed into the undo cap box but not yet confirmed (Enter pressed there clicks OK without
    /// leaving the box). Reading <see cref="NumericUpDown.Value"/> validates the typed text.
    /// </summary>
    private void CommitPendingEdits() => _draft.UndoCellCap = (int)_undoCapBox.Value;

    private void ShowProblems(string heading, IReadOnlyList<string> problems)
    {
        var text = new StringBuilder(heading).Append("\n");
        foreach (var problem in problems.Take(MaxProblemsShown))
        {
            text.Append("\n- ").Append(problem);
        }

        if (problems.Count > MaxProblemsShown)
        {
            text.Append($"\n\n...and {problems.Count - MaxProblemsShown} more.");
        }

        MessageBox.Show(this, text.ToString(), Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }

    /// <summary>Excel's active window (else its main window), as the dialog's owner.</summary>
    private sealed class ExcelWindow : IWin32Window
    {
        public ExcelWindow()
        {
            var active = GetActiveWindow();
            Handle = active != IntPtr.Zero ? active : ExcelDnaUtil.WindowHandle;
        }

        public IntPtr Handle { get; }

        [DllImport("user32.dll")]
        private static extern IntPtr GetActiveWindow();
    }
}
