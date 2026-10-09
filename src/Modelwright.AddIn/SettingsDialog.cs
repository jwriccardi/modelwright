using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using ExcelDna.Integration;
using Modelwright.Core.Formatting;
using Modelwright.Core.Keys;
using Modelwright.Core.Settings;

namespace Modelwright.AddIn;

/// <summary>
/// The settings dialog (docs/PLAN.md section 4.6): a modal WinForms window owned by Excel's active window that
/// edits a <see cref="SettingsDraft"/>, where all the editing logic lives. Cycles tab: items added, removed, moved,
/// renamed and given a code or color, with a live preview; Shortcuts tab: the switch for all keyboard shortcuts, and
/// keys typed or captured from the keyboard, with inline problems; General tab: the undo cap and the diagnostics log. Import, Export, Reset to defaults, OK
/// (validate, save settings.json atomically) and Cancel (discard).
/// </summary>
/// <remarks>
/// WinForms, not WPF: WPF keyboard focus is unreliable on Excel's thread (spike K4). The dialog runs on Excel's
/// main thread inside a macro (<see cref="Commands.MwSettings"/>), so the number format preview calls Excel
/// synchronously. While it is open, <see cref="Form.ShowDialog(IWin32Window)"/> disables every Excel window on the
/// thread. The Ctrl+Z hook passes keys through here, since the focus is not on a worksheet grid.
/// <para>
/// Nothing here may throw into Excel: every event handler runs through <see cref="Safe"/>, and while the dialog is
/// open an unhandled WinForms exception is reported in a message box, never in the ThreadExceptionDialog (whose
/// Quit button would end Excel's process).
/// </para>
/// </remarks>
internal sealed class SettingsDialog : Form
{
    private const int MaxProblemsShown = 12;

    /// <summary>Values shown in the number format preview, as Excel's TEXT function renders them.</summary>
    private static readonly object[] PreviewSamples = { 1234.5, -1234.5, 0.0, "Text" };

    /// <summary>The dialog now open, or null: at most one is open at a time.</summary>
    private static SettingsDialog? _open;

    /// <summary><see cref="Environment.TickCount"/> when the last dialog closed, or null if none has been open.</summary>
    private static int? _closedAt;

    private readonly SettingsDraft _draft;

    /// <summary>The draft's contents when the dialog opened, to tell whether Cancel would discard changes.</summary>
    private readonly string _initialJson;

    private readonly Font _formFont;
    private readonly TabControl _tabs = new TabControl { Dock = DockStyle.Fill };
    private readonly TabPage _cyclesPage;

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
    private readonly Label[] _previewResults = PreviewSamples.Select(_ => new Label { AutoSize = false, AutoEllipsis = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, UseMnemonic = false }).ToArray();
    private readonly Label _localeNote = CreateNote();
    private readonly Control[] _itemEditors;

    /// <summary>The code the preview shows, so it is rendered again only when the code changes.</summary>
    private string? _previewCode;

    /// <summary>Excel's <c>Application.WorksheetFunction</c> (late-bound COM), fetched once per dialog.</summary>
    private object? _worksheetFunction;

    // Shortcuts tab.
    private readonly CheckBox _shortcutsBox = new CheckBox { Text = $"Use &{ProductInfo.Name}'s keyboard shortcuts", AutoSize = true, Anchor = AnchorStyles.Left };
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

    /// <summary>
    /// Creates the dialog. <paramref name="previewMayDiffer"/> shows a note that the number format preview may
    /// not match the cells (Excel's language or decimal separator is not en-US).
    /// </summary>
    private SettingsDialog(SettingsDraft draft, bool previewMayDiffer)
    {
        _draft = draft;
        _initialJson = draft.ContentJson();
        _itemEditors = new Control[] { _nameBox, _codeBox, _colorBox, _pickColorButton, _noFillBox };
        _formFont = SystemFonts.MessageBoxFont; // A new Font each call: this one is the form's, disposed with it.
        _previewFont = new Font(FontFamily.GenericMonospace, _formFont.SizeInPoints);
        _localeNote.Visible = previewMayDiffer;
        _localeNote.Text =
            "Excel's TEXT function reads format codes in your Excel language and regional settings, so this " +
            "preview may differ from the cells, which use the en-US code.";

        Text = ProductInfo.Name + " Settings";
        Font = _formFont;
        AutoScaleMode = AutoScaleMode.Dpi;
        AutoScaleDimensions = new SizeF(96F, 96F);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.Sizable;
        MinimizeBox = false;
        ShowInTaskbar = false;
        ShowIcon = false;
        ClientSize = new Size(820, 600);
        MinimumSize = new Size(680, 520);

        _cyclesPage = CreateCyclesPage();
        _tabs.TabPages.Add(_cyclesPage);
        _tabs.TabPages.Add(CreateShortcutsPage());
        _tabs.TabPages.Add(CreateGeneralPage());

        var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(8), ColumnCount = 1, RowCount = 2 };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.Controls.Add(_tabs, 0, 0);
        root.Controls.Add(CreateButtonBar(), 0, 1);
        Controls.Add(root);

        LoadFromDraft();
        if (Session.SourceState == SettingsLoadOutcome.Rejected)
        {
            SetStatus("settings.json has problems and is not in use; OK will replace it (a backup is kept).", warning: true);
        }
    }

    /// <summary>The settings saved by OK and how they were saved, or null if the dialog was cancelled.</summary>
    private (ToolkitSettings Settings, SettingsFileSaveResult Save)? Saved { get; set; }

    /// <summary>
    /// Shows the dialog, modal and owned by Excel's active window, editing a copy of <paramref name="current"/>.
    /// Returns the settings saved to settings.json by OK with the save's result, or null if the dialog was
    /// cancelled. If a dialog is already open, brings it to the front and returns null; a request made
    /// (<paramref name="requestedAt"/>, <see cref="Environment.TickCount"/>) before the last dialog closed is
    /// ignored (returns null). Call in macro context on Excel's main thread.
    /// </summary>
    public static (ToolkitSettings Settings, SettingsFileSaveResult Save)? Edit(ToolkitSettings current, int requestedAt)
    {
        if (_open is not null)
        {
            _open.Activate();
            return null;
        }

        if (_closedAt is int closedAt && unchecked(requestedAt - closedAt) < 0)
        {
            DiagnosticsLog.Write("SettingsDialogSkipped", "requested before the last dialog closed");
            return null;
        }

        // Showing a form would otherwise install a WinForms SynchronizationContext on Excel's main thread.
        WindowsFormsSynchronizationContext.AutoInstall = false;
        Application.EnableVisualStyles();
        try
        {
            // Unhandled exceptions go to ThreadException (below), not the ThreadExceptionDialog. Only allowed before
            // the thread's first WinForms window; after that, ThreadException is used anyway once it has a handler.
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        }
        catch (InvalidOperationException)
        {
            // A window already exists on this thread (an earlier dialog); the handler below still applies.
        }

        ThreadExceptionEventHandler onThreadException = (sender, e) => ReportError(_open, e.Exception);
        Application.ThreadException += onThreadException;
        var previousDpiContext = DpiContext.EnterSystemAware();
        try
        {
            using var dialog = new SettingsDialog(new SettingsDraft(current), PreviewMayDiffer());
            _open = dialog;
            return dialog.ShowDialog(new ExcelWindow()) == DialogResult.OK ? dialog.Saved : null;
        }
        finally
        {
            _open = null;
            _closedAt = Environment.TickCount;
            DpiContext.Restore(previousDpiContext);
            Application.ThreadException -= onThreadException;
        }
    }

    /// <summary>Capture mode takes the next key press, before any control or mnemonic sees it.</summary>
    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (!_capturing)
        {
            return base.ProcessCmdKey(ref msg, keyData);
        }

        var handled = true;
        var key = keyData;
        Safe(() => handled = CaptureKey(key));
        return handled;
    }

    /// <inheritdoc />
    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        Safe(() =>
        {
            if (DialogResult != DialogResult.OK &&
                (e.CloseReason == CloseReason.UserClosing || e.CloseReason == CloseReason.None) &&
                HasUnsavedChanges())
            {
                var answer = MessageBox.Show(
                    this,
                    "Discard your changes?",
                    Text,
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2);
                e.Cancel = answer != DialogResult.Yes;
            }
        });
        base.OnFormClosing(e);
    }

    /// <summary>Handles a key press in capture mode; returns true if it was used.</summary>
    private bool CaptureKey(Keys keyData)
    {
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

        // KeyCapture names punctuation keys by their US character; the layout's own character is used where it differs.
        var text = KeyboardLayout.ToLayoutKeyString(keyCode, modifiers, KeyCapture.ToKeyString(keyCode, modifiers), out var layoutProblem);
        if (text is null)
        {
            StopCapture(layoutProblem ??
                "That key can't be part of a shortcut. Use a letter, digit, punctuation key, F1-F12, an arrow, " +
                "PgUp, PgDn, Home, End, Ins or Del, with Ctrl, Alt and/or Shift.");
            return true;
        }

        var altGrWarning = KeyboardLayout.AltGrWarning(keyCode, modifiers);
        StopCapture(null);
        _keyBox.Text = text; // Updates the draft.
        var problem = SelectedAction is { } actionId ? _draft.KeyProblem(actionId) : null;
        _keyError.Text = string.Join(" ", new[] { problem, altGrWarning }.Where(m => m is not null));
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
        if (disposing)
        {
            _formFont.Dispose(); // After the controls, which use it.
        }
    }

    /// <summary>
    /// Logs <paramref name="ex"/> and shows its message over <paramref name="owner"/> (or Excel, if null). Never
    /// throws: an error in a settings dialog handler must not reach Excel.
    /// </summary>
    private static void ReportError(IWin32Window? owner, Exception ex)
    {
        DiagnosticsLog.Write("SettingsDialogError", ex.ToString());
        try
        {
            MessageBox.Show(
                owner,
                $"Something went wrong: {ex.Message}\n\nThe dialog is still open; the details are in the diagnostics log.",
                ProductInfo.Name + " Settings",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        catch (Exception)
        {
            // Nothing more can be done; the error is logged.
        }
    }

    /// <summary>
    /// True if Excel's number format preview may not match the cells: its decimal separator is not "." or its
    /// display language is not English (US). Call in macro context. Any error counts as false.
    /// </summary>
    private static bool PreviewMayDiffer()
    {
        const int xlDecimalSeparator = 3;
        const int msoLanguageIDUI = 2;
        const int englishUs = 1033;
        try
        {
            var separator = GetProperty(ExcelDnaUtil.Application, "International", xlDecimalSeparator);
            if (!string.Equals(Convert.ToString(separator, CultureInfo.InvariantCulture), ".", StringComparison.Ordinal))
            {
                return true;
            }
        }
        catch (Exception)
        {
            // Unknown: fall through to the language check.
        }

        try
        {
            var languageSettings = GetProperty(ExcelDnaUtil.Application, "LanguageSettings");
            return Convert.ToInt32(GetProperty(languageSettings, "LanguageID", msoLanguageIDUI), CultureInfo.InvariantCulture) != englishUs;
        }
        catch (Exception)
        {
            return false;
        }
    }

    /// <summary>A late-bound COM property get, with arguments for a parameterized property.</summary>
    private static object GetProperty(object target, string name, params object[] arguments) =>
        target.GetType().InvokeMember(name, BindingFlags.GetProperty, null, target, arguments, CultureInfo.InvariantCulture);

    /// <summary>Runs an event handler's work, reporting any exception instead of letting it reach Excel.</summary>
    private void Safe(Action action)
    {
        try
        {
            action();
        }
        catch (Exception ex)
        {
            ReportError(this, ex);
        }
    }

    /// <summary>Shows <paramref name="text"/> in the bottom bar, in red if it is a <paramref name="warning"/>.</summary>
    private void SetStatus(string text, bool warning = false)
    {
        _status.ForeColor = warning ? Color.Firebrick : SystemColors.ControlText;
        _status.Text = text;
    }

    private bool HasUnsavedChanges()
    {
        CommitPendingEdits();
        return !string.Equals(_draft.ContentJson(), _initialJson, StringComparison.Ordinal);
    }

    private static string ColorText(OleColor color) =>
        color.IsNoFill
            ? "No fill"
            : string.Format(CultureInfo.InvariantCulture, "{0}  rgb({1}, {2}, {3})", color, color.R, color.G, color.B);

    private static string CycleLabel(CycleDraft cycle) => cycle.Provisional ? cycle.DisplayName + " (provisional)" : cycle.DisplayName;

    /// <summary>
    /// <paramref name="value"/> formatted with <paramref name="code"/> by Excel's own formatter,
    /// <c>WorksheetFunction.Text</c> (late-bound COM), or a note if Excel rejects the code. TEXT reads the code in
    /// Excel's display language, so outside en-US the preview can differ from the cell, which uses en-US syntax
    /// (the dialog then says so under the preview).
    /// </summary>
    private string RenderSample(object value, string code)
    {
        if (code.Trim().Length == 0)
        {
            return "(no format code)";
        }

        try
        {
            if (_worksheetFunction is null)
            {
                dynamic application = ExcelDnaUtil.Application;
                _worksheetFunction = application.WorksheetFunction;
            }

            dynamic worksheetFunction = _worksheetFunction!;
            object result = worksheetFunction.Text(value, code);
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
        preview.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        preview.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); // Long results are cut with an ellipsis.
        for (var i = 0; i < PreviewSamples.Length; i++)
        {
            preview.Controls.Add(new Label { Text = DescribeSample(PreviewSamples[i]), AutoSize = true, Anchor = AnchorStyles.Left, UseMnemonic = false }, 0, i);
            _previewResults[i].Font = _previewFont;
            preview.Controls.Add(_previewResults[i], 1, i);
        }

        preview.Controls.Add(_localeNote, 0, PreviewSamples.Length);
        preview.SetColumnSpan(_localeNote, 2);

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

        _cycleList.SelectedIndexChanged += (s, e) => Safe(() => ShowCycle(0));
        _itemList.SelectedIndexChanged += (s, e) => Safe(ShowItem);
        _itemList.Resize += (s, e) => Safe(() => _itemValueColumn.Width = -2); // The last column fills the list.
        _addButton.Click += (s, e) => Safe(AddItem);
        _removeButton.Click += (s, e) => Safe(() => EditItems(c =>
        {
            var index = SelectedItem;
            c.RemoveItem(index);
            return Math.Min(index, c.Items.Count - 1);
        }));
        _upButton.Click += (s, e) => Safe(() => EditItems(c => c.MoveItemUp(SelectedItem)));
        _downButton.Click += (s, e) => Safe(() => EditItems(c => c.MoveItemDown(SelectedItem)));
        _nameBox.TextChanged += (s, e) => Safe(() => EditSelectedItem(c => c.RenameItem(SelectedItem, _nameBox.Text)));
        _codeBox.TextChanged += (s, e) => Safe(() => EditSelectedItem(c => c.SetCode(SelectedItem, _codeBox.Text)));
        _colorBox.TextChanged += (s, e) => Safe(ColorTyped); // Invalid text stays, with its error, until fixed; OK refuses it.
        _noFillBox.CheckedChanged += (s, e) => Safe(NoFillToggled);
        _pickColorButton.Click += (s, e) => Safe(PickColor);

        var page = new TabPage("Cycles") { UseVisualStyleBackColor = true };
        page.Controls.Add(layout);
        return page;
    }

    private TabPage CreateShortcutsPage()
    {
        _actionList.Columns.Add("Action", LogicalToDeviceUnits(140));
        _actionList.Columns.Add("Shortcut", LogicalToDeviceUnits(150));
        _actionList.Columns.Add("Problem", LogicalToDeviceUnits(400));

        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 7, Padding = new Padding(4) };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        layout.Controls.Add(_shortcutsBox, 0, 0);
        layout.SetColumnSpan(_shortcutsBox, 4);
        var switchNote = CreateNote();
        switchNote.Text =
            "Off: no key below is bound and Ctrl+Z / Ctrl+Y are left to Excel, so Macabacus, which uses the same " +
            "keys, answers them. Switched off while Excel runs, they do Excel's usual thing until you restart Excel; " +
            "then Macabacus has them. The ribbon works either way, and the keys are kept for when you switch back on " +
            "(that takes effect at once). The ribbon's Shortcuts button does the same.";
        switchNote.Margin = new Padding(3, 0, 3, 8);
        layout.Controls.Add(switchNote, 0, 1);
        layout.SetColumnSpan(switchNote, 4);

        var listLabel = CreateLabel("&Actions:");
        layout.Controls.Add(listLabel, 0, 2);
        layout.SetColumnSpan(listLabel, 4);
        layout.Controls.Add(_actionList, 0, 3);
        layout.SetColumnSpan(_actionList, 4);
        layout.Controls.Add(CreateLabel("&Shortcut:"), 0, 4);
        layout.Controls.Add(_keyBox, 1, 4);
        layout.Controls.Add(_captureButton, 2, 4);
        layout.Controls.Add(_clearKeyButton, 3, 4);
        layout.Controls.Add(_keyError, 1, 5);
        layout.SetColumnSpan(_keyError, 3);
        var note = CreateNote();
        note.Text =
            "Type a shortcut such as Ctrl+Shift+K, or click Capture and press it. Use Ctrl and/or Alt, with or " +
            "without Shift. Some keys override Excel's built-ins while the add-in is loaded (e.g. Ctrl+; inserts " +
            "the date). Changes take effect when you click OK.";
        layout.Controls.Add(note, 0, 6);
        layout.SetColumnSpan(note, 4);

        _shortcutsBox.CheckedChanged += (s, e) => Safe(() =>
        {
            if (!_loading)
            {
                _draft.UseKeyboardShortcuts = _shortcutsBox.Checked;
            }
        });
        _actionList.SelectedIndexChanged += (s, e) => Safe(ShowAction);
        _keyBox.TextChanged += (s, e) => Safe(KeyTyped);
        _keyBox.Leave += (s, e) => Safe(ShowAction); // Shows the key in its standard form.
        _captureButton.Click += (s, e) => Safe(StartCapture);
        _captureButton.Leave += (s, e) => Safe(() => StopCapture(null));
        _clearKeyButton.Click += (s, e) => Safe(() =>
        {
            _keyBox.Text = string.Empty;
            _keyBox.Focus();
        });

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

        _undoCapBox.ValueChanged += (s, e) => Safe(() =>
        {
            if (!_loading)
            {
                _draft.UndoCellCap = (int)_undoCapBox.Value;
            }
        });
        _logBox.CheckedChanged += (s, e) => Safe(() =>
        {
            if (!_loading)
            {
                _draft.DiagnosticsLog = _logBox.Checked;
            }
        });
        openLog.Click += (s, e) => Safe(OpenLogFolder);
        openSettings.Click += (s, e) => Safe(() =>
        {
            var problem = Commands.OpenSettingsFile();
            if (problem is not null)
            {
                MessageBox.Show(this, problem, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        });

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

        import.Click += (s, e) => Safe(Import);
        export.Click += (s, e) => Safe(Export);
        reset.Click += (s, e) => Safe(ResetToDefaults);
        ok.Click += (s, e) => Safe(Save);
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
            _shortcutsBox.Checked = _draft.UseKeyboardShortcuts;
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
                // With its handle created, the image list copies the bitmap, so it can be disposed at once.
                _ = _swatches.Handle;
                using var swatch = CreateSwatch(color.Color);
                _swatches.Images.Add(key, swatch);
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
        var code = (item as NumberFormatItem)?.Code;
        if (!string.Equals(code, _previewCode, StringComparison.Ordinal))
        {
            // Each render is a call into Excel, so only when the code changes.
            _previewCode = code;
            for (var i = 0; i < PreviewSamples.Length; i++)
            {
                _previewResults[i].Text = code is null ? string.Empty : RenderSample(PreviewSamples[i], code);
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
        var same = cycle.IndexOfSameValue(index);
        if (same >= 0)
        {
            SetStatus(
                $"Item {same + 1} ({cycle.Items[same].Name}) has the same {(cycle.Kind == CycleKind.NumberFormat ? "code" : "color")}: " +
                "the cycle resumes after the first copy. Give the new item its own value.",
                warning: true);
        }
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
            var fileName = Path.GetFileName(open.FileName);
            var tooLarge = ToolkitSettings.CheckFileSize(new FileInfo(open.FileName).Length, fileName);
            problems = tooLarge is not null ? new[] { tooLarge } : _draft.ImportFile(File.ReadAllBytes(open.FileName), fileName);
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
        SetStatus($"Imported {Path.GetFileName(open.FileName)}. Click OK to save and apply.");
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
            SetStatus($"Exported to {Path.GetFileName(save.FileName)}.");
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
        SetStatus("Defaults restored. Click OK to save and apply, or Cancel to keep your settings.");
    }

    private void Save()
    {
        if (_colorError.Text.Length > 0)
        {
            // A color typed but not valid: the item still has its old color, which OK would silently keep.
            SetStatus("Fix the color (or pick one) before clicking OK.", warning: true);
            _tabs.SelectedTab = _cyclesPage;
            _colorBox.Focus();
            _colorBox.SelectAll();
            return;
        }

        CommitPendingEdits();
        var settings = _draft.ToSettings(out var problems);
        if (problems.Count > 0)
        {
            ShowProblems("The settings can't be saved yet. Fix these problems, or click Cancel:", problems);
            return;
        }

        // Never overwrite silently a file edited by hand since we read it, or one that was rejected.
        if (SettingsStore.Exists() &&
            (Session.SourceState == SettingsLoadOutcome.Rejected || SettingsStore.HasChangedSince(Session.SourceHash)))
        {
            var answer = MessageBox.Show(
                this,
                "settings.json was changed outside this dialog (or has problems and isn't in use). Replace it with " +
                "these settings? The current file will be kept as a backup.\n\nChoose No to return; you can Cancel " +
                "and use Reload settings instead.",
                Text,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);
            if (answer != DialogResult.Yes)
            {
                return;
            }
        }

        var save = SettingsStore.Save(settings);
        if (!save.Succeeded)
        {
            MessageBox.Show(this, $"The settings were not saved: {save.Problem}", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        Saved = (settings, save);
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

    /// <summary>
    /// The thread's DPI awareness while the dialog exists: system-aware, so WinForms (.NET Framework) lays the
    /// dialog out once at the system DPI and Windows scales it on other monitors, rather than a per-monitor-aware
    /// Excel thread leaving it unscaled or clipped on a monitor at another DPI.
    /// </summary>
    internal static class DpiContext
    {
        private static readonly IntPtr SystemAware = new IntPtr(-2); // DPI_AWARENESS_CONTEXT_SYSTEM_AWARE

        /// <summary>Makes the thread system-aware. Returns the previous context, or zero if it was not changed.</summary>
        public static IntPtr EnterSystemAware()
        {
            try
            {
                return SetThreadDpiAwarenessContext(SystemAware);
            }
            catch (EntryPointNotFoundException)
            {
                return IntPtr.Zero; // Before Windows 10 1607 a thread cannot change its awareness.
            }
        }

        /// <summary>Puts back the context <see cref="EnterSystemAware"/> returned.</summary>
        public static void Restore(IntPtr previous)
        {
            if (previous != IntPtr.Zero)
            {
                SetThreadDpiAwarenessContext(previous);
            }
        }

        [DllImport("user32.dll")]
        private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr dpiContext);
    }

    /// <summary>Excel's active window (else its main window), as the owner of a dialog or message box.</summary>
    internal static IWin32Window ExcelOwner() => new ExcelWindow();

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
