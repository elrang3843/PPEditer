using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using PPEditer.Dialogs;
using PPEditer.Export;
using PPEditer.Models;
using PPEditer.Services;

namespace PPEditer;

public partial class MainWindow : Window
{
    private readonly PresentationModel _model = new();
    private int  _currentSlide;
    private bool _suppressZoomEvents;
    private bool _suppressFormatEvents;
    private bool _suppressNotesEvents;
    private bool _notesDirty;
    private RichTextBox? _activeEditor;

    private const int RecentMax = 10;

    public MainWindow()
    {
        InitializeComponent();

        SlidePanel.SlideSelected    += idx =>
        {
            SaveCurrentNotes();
            _currentSlide = idx;
            ShowCurrentSlide();
            LoadCurrentNotes();
            UpdateActions();
        };
        EditorCanvas.EditingStarted += OnEditorEditingStarted;
        EditorCanvas.TextCommitted  += OnEditorTextCommitted;
        EditorCanvas.ShapeMoved     += OnShapeMoved;
        EditorCanvas.ShapeDeleted            += OnShapeDeleted;
        EditorCanvas.ShapeResized            += OnShapeResized;
        EditorCanvas.ShapePropertiesRequested += OnShapePropertiesRequested;
        EditorCanvas.ShapeOrderChanged        += OnShapeOrderChanged;
        EditorCanvas.ShapeDrawn               += OnShapeDrawn;
        EditorCanvas.ShapesGroupRequested     += OnShapesGroupRequested;
        EditorCanvas.ShapeUngroupRequested    += OnShapeUngroupRequested;
        EditorCanvas.CharPropertiesRequested  += (_, _) => OnCharProperties();
        EditorCanvas.ParaPropertiesRequested  += (_, _) => OnParaProperties();
        EditorCanvas.TextBoxDrawn             += OnTextBoxDrawn;
        EditorCanvas.ShapeRotated             += OnShapeRotated;
        EditorCanvas.SelectionChanged         += UpdateActions;

        RegisterKeyBindings();
        InitSettings();
        RebuildRecentMenu();
        UpdateActions();
    }

    // ── Settings initialisation ───────────────────────────────────────

    private void InitSettings()
    {
        var s = AppSettings.Current;
        MenuFontFilter.IsChecked   = s.FontLicenseFilterEnabled;
        ChkOpenFontsOnly.IsChecked = s.FontLicenseFilterEnabled;
        SyncLangMenu(s.Language);
        SyncThemeMenu(s.Theme);
        PopulateFontCombo(s.FontLicenseFilterEnabled);
        PopulateSizeCombo();
    }

    private void PopulateFontCombo(bool openOnly)
    {
        if (FontFamilyCombo is null) return;   // guard: called before InitializeComponent finishes
        _suppressFormatEvents = true;
        FontFamilyCombo.ItemsSource = openOnly
            ? FontService.GetLicenseFilteredFonts()
            : FontService.GetAllFonts();
        _suppressFormatEvents = false;
    }

    private static void PopulateSizeCombo() { /* sizes are defined in XAML */ }

    // ── Keyboard bindings ─────────────────────────────────────────────

    private void RegisterKeyBindings()
    {
        var kb = InputBindings;
        kb.Add(new KeyBinding(new RelayCommand(OnNew),           Key.N, ModifierKeys.Control));
        kb.Add(new KeyBinding(new RelayCommand(OnOpen),          Key.O, ModifierKeys.Control));
        kb.Add(new KeyBinding(new RelayCommand(OnSave),          Key.S, ModifierKeys.Control));
        kb.Add(new KeyBinding(new RelayCommand(OnSaveAs),        Key.S, ModifierKeys.Control | ModifierKeys.Shift));
        kb.Add(new KeyBinding(new RelayCommand(OnExportPdf),     Key.E, ModifierKeys.Control | ModifierKeys.Shift));
        kb.Add(new KeyBinding(new RelayCommand(OnUndo),          Key.Z, ModifierKeys.Control));
        kb.Add(new KeyBinding(new RelayCommand(OnRedo),          Key.Y, ModifierKeys.Control));
        kb.Add(new KeyBinding(new RelayCommand(OnAddSlide),      Key.M, ModifierKeys.Control));
        kb.Add(new KeyBinding(new RelayCommand(OnDupSlide),      Key.D, ModifierKeys.Control));
        kb.Add(new KeyBinding(new RelayCommand(OnDelSlide),      Key.Delete, ModifierKeys.Control | ModifierKeys.Shift));
        kb.Add(new KeyBinding(new RelayCommand(OnSlideUp),       Key.Up,   ModifierKeys.Control | ModifierKeys.Shift));
        kb.Add(new KeyBinding(new RelayCommand(OnSlideDown),     Key.Down, ModifierKeys.Control | ModifierKeys.Shift));
        kb.Add(new KeyBinding(new RelayCommand(OnZoomIn),        Key.OemPlus,  ModifierKeys.Control));
        kb.Add(new KeyBinding(new RelayCommand(OnZoomOut),       Key.OemMinus, ModifierKeys.Control));
        kb.Add(new KeyBinding(new RelayCommand(OnZoomFit),       Key.D0, ModifierKeys.Control));
        kb.Add(new KeyBinding(new RelayCommand(OnInsertTextBox), Key.T, ModifierKeys.Control | ModifierKeys.Shift));
        kb.Add(new KeyBinding(new RelayCommand(OnDocInfo),       Key.I, ModifierKeys.Control | ModifierKeys.Shift));
        kb.Add(new KeyBinding(new RelayCommand(OnInsertMath),    Key.M, ModifierKeys.Control | ModifierKeys.Shift));
        kb.Add(new KeyBinding(new RelayCommand(OnInsertCharMap), Key.K, ModifierKeys.Control | ModifierKeys.Shift));
        kb.Add(new KeyBinding(new RelayCommand(OnInsertEmoji),   Key.J, ModifierKeys.Control | ModifierKeys.Shift));
        kb.Add(new KeyBinding(new RelayCommand(_ => OnGroup()),   Key.G, ModifierKeys.Control));
        kb.Add(new KeyBinding(new RelayCommand(_ => OnUngroup()), Key.G, ModifierKeys.Control | ModifierKeys.Shift));
        kb.Add(new KeyBinding(new RelayCommand(OnCharProperties), Key.F, ModifierKeys.Control | ModifierKeys.Shift));
        kb.Add(new KeyBinding(new RelayCommand(OnParaProperties), Key.P, ModifierKeys.Control | ModifierKeys.Shift));
        kb.Add(new KeyBinding(new RelayCommand(_ => OnPrint()),    Key.P, ModifierKeys.Control));
        kb.Add(new KeyBinding(new RelayCommand(_ => OnSlideShow()), Key.F5, ModifierKeys.None));
    }

    // ── File commands ─────────────────────────────────────────────────

    private void OnNew(object? _ = null)
    {
        if (!ConfirmDiscard()) return;
        _model.New();
        _currentSlide = 0;
        RefreshAll();
        SetStatus(S("Msg_NewCreated"));
    }

    private void OnOpen(object? _ = null)
    {
        if (!ConfirmDiscard()) return;
        var dlg = new OpenFileDialog
        {
            Filter = S("Dlg_OpenFilter"),
            Title  = S("Dlg_OpenTitle"),
        };
        if (dlg.ShowDialog(this) == true)
            OpenFile(dlg.FileName);
    }

    public void OpenFile(string path)
    {
        try
        {
            _model.Open(path);
            _currentSlide = 0;
            AddRecent(path);
            RebuildRecentMenu();
            RefreshAll();
            SetStatus(string.Format(S("Msg_Opened"), Path.GetFileName(path)));
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"{S("Err_OpenFailed")}\n{ex.Message}",
                            S("Lbl_OpenError"), MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnSave(object? _ = null)
    {
        if (!_model.IsOpen) return;
        if (string.IsNullOrEmpty(_model.FilePath)) { OnSaveAs(); return; }
        SaveCurrentNotes();
        try
        {
            _model.Save();
            UpdateTitle();
            UpdateActions();
            SetStatus(S("Msg_Saved"));
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"{S("Err_SaveFailed")}\n{ex.Message}",
                            S("Lbl_SaveError"), MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnSaveAs(object? _ = null)
    {
        if (!_model.IsOpen) return;
        SaveCurrentNotes();
        var dlg = new SaveFileDialog
        {
            Filter     = S("Dlg_SaveFilter"),
            Title      = S("Dlg_SaveTitle"),
            DefaultExt = "pptx",
            FileName   = _model.FileName,
        };
        if (dlg.ShowDialog(this) != true) return;
        var path = dlg.FileName;
        if (!path.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase)) path += ".pptx";
        try
        {
            _model.Save(path);
            AddRecent(path);
            RebuildRecentMenu();
            UpdateTitle();
            UpdateActions();
            SetStatus(string.Format(S("Msg_SavedAs"), Path.GetFileName(path)));
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"{S("Err_SaveFailed")}\n{ex.Message}",
                            S("Lbl_SaveError"), MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnExportPdf(object? _ = null)
    {
        if (!_model.IsOpen) return;
        var defaultName = Path.ChangeExtension(_model.FileName, ".pdf");
        var dlg = new SaveFileDialog
        {
            Filter     = S("Dlg_PdfFilter"),
            Title      = S("Dlg_PdfTitle"),
            DefaultExt = "pdf",
            FileName   = defaultName,
        };
        if (dlg.ShowDialog(this) != true) return;

        var progressDlg = new Window
        {
            Title  = S("Dlg_Exporting"),
            Width  = 300, Height = 100,
            Owner  = this,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ResizeMode = ResizeMode.NoResize,
        };
        var bar = new ProgressBar { Minimum = 0, Maximum = _model.SlideCount,
                                    Margin = new Thickness(16) };
        progressDlg.Content = bar;
        progressDlg.Show();

        try
        {
            var progress = new Progress<(int c, int t)>(p =>
            {
                bar.Value = p.c;
                progressDlg.Title = string.Format(S("Msg_Exporting"), p.c, p.t);
            });
            bool ok = PdfExporter.Export(_model, dlg.FileName, progress);
            progressDlg.Close();
            if (ok)
                SetStatus(string.Format(S("Msg_PdfDone"), Path.GetFileName(dlg.FileName)));
            else
                MessageBox.Show(this, S("Err_PdfFailed"), S("Lbl_Error"),
                                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        catch (Exception ex)
        {
            progressDlg.Close();
            MessageBox.Show(this, $"{S("Err_PdfFailed")}\n{ex.Message}", S("Lbl_Error"),
                            MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnExit(object? _ = null) => Close();

    // ── Document info ─────────────────────────────────────────────────

    private void OnDocInfo(object? _ = null)
    {
        if (!_model.IsOpen) return;
        var dlg = new PPEditer.Dialogs.DocInfoDialog(
            _model.GetDocProperties(), _model.HasWriteProtection)
        {
            Owner = this
        };
        if (dlg.ShowDialog() != true || dlg.Result is null) return;

        _model.UpdateDocInfo(dlg.Result, dlg.SetProtect, dlg.WritePassword, dlg.RemoveProtect);
        UpdateTitle();
        UpdateActions();
        SetStatus(S("St_DocInfoSaved"));
    }

    // ── Edit commands ─────────────────────────────────────────────────

    private void OnUndo(object? _ = null)
    {
        if (_model.Undo())
        {
            _currentSlide = Math.Min(_currentSlide, _model.SlideCount - 1);
            RefreshAll();
        }
    }

    private void OnRedo(object? _ = null)
    {
        if (_model.Redo())
        {
            _currentSlide = Math.Min(_currentSlide, _model.SlideCount - 1);
            RefreshAll();
        }
    }

    // ── Slide commands ────────────────────────────────────────────────

    private void OnAddSlide(object? _ = null)
    {
        if (!_model.IsOpen) return;
        _currentSlide = _model.AddSlide(_currentSlide);
        RefreshAll();
    }

    private void OnDupSlide(object? _ = null)
    {
        if (!_model.IsOpen) return;
        _currentSlide = _model.DuplicateSlide(_currentSlide);
        RefreshAll();
    }

    private void OnDelSlide(object? _ = null)
    {
        if (!_model.IsOpen) return;
        if (_model.SlideCount <= 1)
        {
            MessageBox.Show(this, S("Err_LastSlide"),
                            S("Lbl_DeleteFail"), MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        var r = MessageBox.Show(this,
            string.Format(S("Msg_ConfirmDelete"), _currentSlide + 1),
            S("Lbl_DeleteConfirm"), MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (r != MessageBoxResult.Yes) return;
        _model.DeleteSlide(_currentSlide);
        _currentSlide = Math.Min(_currentSlide, _model.SlideCount - 1);
        RefreshAll();
    }

    private void OnSlideUp(object? _ = null)
    {
        if (!_model.IsOpen || _currentSlide <= 0) return;
        _model.MoveSlide(_currentSlide, _currentSlide - 1);
        _currentSlide--;
        RefreshAll();
    }

    private void OnSlideDown(object? _ = null)
    {
        if (!_model.IsOpen || _currentSlide >= _model.SlideCount - 1) return;
        _model.MoveSlide(_currentSlide, _currentSlide + 1);
        _currentSlide++;
        RefreshAll();
    }

    // ── Insert commands ───────────────────────────────────────────────

    private void OnInsertTextBox(object? _ = null)
    {
        if (!_model.IsOpen) return;
        EditorCanvas.ActiveTool = DrawTool.TextBox;
        SetStatus(S("St_DrawHint_Drag"));
    }

    private void OnTextBoxDrawn(int slideIdx, long leftEmu, long topEmu, long widthEmu, long heightEmu)
    {
        int treeIdx = _model.AddTextBox(slideIdx, leftEmu, topEmu, widthEmu, heightEmu);
        RefreshAll();
        SetStatus(S("Msg_TextBoxAdded"));
        if (treeIdx >= 0)
            EditorCanvas.StartEditByTreeIndex(treeIdx, selectAll: true);
        UpdateActions();
    }

    // ── View commands ─────────────────────────────────────────────────

    private void OnZoomIn(object?  _ = null) { EditorCanvas.ZoomIn();      UpdateZoomDisplay(); }
    private void OnZoomOut(object? _ = null) { EditorCanvas.ZoomOut();     UpdateZoomDisplay(); }
    private void OnZoomFit(object? _ = null) { EditorCanvas.FitToWindow(); UpdateZoomDisplay(); }

    private void OnTogglePanel(object? _ = null)
    {
        bool vis = MenuTogglePanel.IsChecked;
        SlidePanel.Visibility = vis ? Visibility.Visible : Visibility.Collapsed;
        PanelCol.Width        = vis ? new GridLength(200) : new GridLength(0);
    }

    private void OnToggleStatus(object? _ = null)
        => AppStatusBar.Visibility = MenuToggleStatus.IsChecked
            ? Visibility.Visible : Visibility.Collapsed;

    private void OnToggleNotes(object sender, RoutedEventArgs e)
    {
        bool show = MenuToggleNotes.IsChecked;
        NotesRow.Height          = show ? new GridLength(150, GridUnitType.Pixel) : new GridLength(0);
        NotesSplitterRow.Height  = show ? new GridLength(4)                       : new GridLength(0);
    }

    // ── Notes panel ───────────────────────────────────────────────────

    private void LoadCurrentNotes()
    {
        if (!_model.IsOpen) return;
        _suppressNotesEvents = true;
        NotesTextBox.Text    = _model.GetSlideNotes(_currentSlide);
        _suppressNotesEvents = false;
        _notesDirty          = false;
    }

    private void SaveCurrentNotes()
    {
        if (!_model.IsOpen || !_notesDirty) return;
        _model.SetSlideNotes(_currentSlide, NotesTextBox.Text);
        _notesDirty = false;
    }

    private void OnNotesTextChanged(object sender, TextChangedEventArgs e)
    {
        if (!_suppressNotesEvents)
            _notesDirty = true;
    }

    // ── Print ─────────────────────────────────────────────────────────

    private void OnPrint(object sender, RoutedEventArgs e) => OnPrint();
    private void OnPrint()
    {
        if (!_model.IsOpen) return;
        SaveCurrentNotes();
        var dlg = new PrintLayoutDialog(this);
        if (dlg.ShowDialog() != true) return;
        try
        {
            PrintExporter.Print(_model, dlg.SelectedLayout, dlg.BlackWhite);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"{S("Err_PrintFailed")}\n{ex.Message}",
                            S("Lbl_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnUserManual(object s, RoutedEventArgs e) => new UserManualDialog { Owner = this }.ShowDialog();
    private void OnLicense   (object s, RoutedEventArgs e) => new LicenseDialog   { Owner = this }.ShowDialog();
    private void OnAbout(object? _ = null) => new AboutDialog { Owner = this }.ShowDialog();

    // ── Zoom combo ────────────────────────────────────────────────────

    private void ZoomCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppressZoomEvents) return;
        if (ZoomCombo.SelectedItem is ComboBoxItem item)
            ApplyZoomText(item.Content.ToString()!);
    }

    private void ZoomCombo_LostFocus(object sender, RoutedEventArgs e)
    {
        if (_suppressZoomEvents) return;
        ApplyZoomText(ZoomCombo.Text);
    }

    private void ApplyZoomText(string text)
    {
        if (EditorCanvas is null) return;
        if (double.TryParse(text.TrimEnd('%'), out double pct))
            EditorCanvas.SetZoom(pct / 100.0);
        UpdateZoomDisplay();
    }

    // ── Editing events from EditorCanvas ──────────────────────────────

    private void OnEditorEditingStarted(RichTextBox rtb)
    {
        _activeEditor = rtb;
        FormatToolBar.IsEnabled = true;
        TbInsertTxBox.IsEnabled = false;
        rtb.SelectionChanged += Editor_SelectionChanged;
        SyncFormatToolbar();
    }

    private void OnShapeMoved(int slideIdx, int treeIdx, long dx, long dy)
    {
        _model.MoveShape(slideIdx, treeIdx, dx, dy);
        EditorCanvas.Invalidate();
        SlidePanel.RefreshSingle(slideIdx);
        UpdateTitle();
        UpdateActions();
    }

    private void OnShapeDeleted(int slideIdx, int treeIdx)
    {
        _model.DeleteShape(slideIdx, treeIdx);
        EditorCanvas.Invalidate(preserveSelection: false);
        SlidePanel.RefreshSingle(slideIdx);
        UpdateTitle();
        UpdateActions();
        SetStatus(S("Msg_ShapeDeleted"));
    }

    private void OnShapeResized(int slideIdx, int treeIdx,
        long leftEmu, long topEmu, long widthEmu, long heightEmu)
    {
        _model.ResizeShape(slideIdx, treeIdx, leftEmu, topEmu, widthEmu, heightEmu);
        EditorCanvas.Invalidate();
        SlidePanel.RefreshSingle(slideIdx);
        UpdateTitle();
        UpdateActions();
    }

    private void OnShapeRotated(int slideIdx, int treeIdx, double angleDelta)
    {
        _model.RotateShape(slideIdx, treeIdx, angleDelta);
        EditorCanvas.Invalidate(treeIdx);
        SlidePanel.RefreshSingle(slideIdx);
        UpdateTitle();
        UpdateActions();
    }

    private void OnShapePropertiesRequested(int slideIdx, int treeIdx)
    {
        var style = _model.GetShapeStyle(slideIdx, treeIdx);
        if (style is null) return;
        if (string.IsNullOrEmpty(style.Name))
            style.Name = $"개체 {treeIdx}";

        var dlg = new PPEditer.Dialogs.ShapePropertiesDialog(style) { Owner = this };
        EditorCanvas.SuppressLostFocusCommit = true;
        try
        {
            if (dlg.ShowDialog() != true) return;
        }
        finally
        {
            EditorCanvas.SuppressLostFocusCommit = false;
        }
        _model.UpdateShapeStyle(slideIdx, treeIdx, dlg.Result);
        EditorCanvas.Invalidate();
        SlidePanel.RefreshSingle(slideIdx);
        UpdateTitle();
        UpdateActions();
        SetStatus(S("Msg_ShapePropsSaved"));
    }

    private void OnShapeOrderChanged(int slideIdx, int treeIdx, int delta)
    {
        int newIdx = delta switch
        {
             2 => _model.BringShapeToFront(slideIdx, treeIdx),
             1 => _model.BringShapeForward(slideIdx, treeIdx),
            -1 => _model.SendShapeBackward(slideIdx, treeIdx),
            -2 => _model.SendShapeToBack(slideIdx, treeIdx),
            _  => treeIdx,
        };
        EditorCanvas.Invalidate(newIdx);
        SlidePanel.RefreshSingle(slideIdx);
        UpdateTitle();
        UpdateActions();
    }

    // ── Insert shape drawing ─────────────────────────────────────────

    private void OnDrawShape(object sender, RoutedEventArgs e)
    {
        if (!_model.IsOpen) return;
        if (sender is not MenuItem mi || mi.Tag is not string tagStr) return;
        if (!Enum.TryParse<DrawTool>(tagStr, out var tool)) return;
        EditorCanvas.CommitEdit(save: true);
        EditorCanvas.ActiveTool = tool;
        string hint = tool switch
        {
            DrawTool.ScaleneTriangle                                    => S("St_DrawHint_3Click"),
            DrawTool.Trapezoid                                          => S("St_DrawHint_4Click"),
            DrawTool.PolyLine or DrawTool.SplineLine or
            DrawTool.Polygon  or DrawTool.SplinePolygon                 => S("St_DrawHint_Click"),
            _                                                           => S("St_DrawHint_Drag"),
        };
        SetStatus(hint);
    }

    private void OnShapeDrawn(int slideIdx, DrawTool tool, System.Windows.Point[] points)
    {
        double epp = EditorCanvas.EmuPerPixel;
        long[] xs = points.Select(p => (long)(p.X * epp)).ToArray();
        long[] ys = points.Select(p => (long)(p.Y * epp)).ToArray();
        int newIdx = _model.AddDrawnShape(slideIdx, tool, xs, ys);
        EditorCanvas.Invalidate(newIdx >= 0 ? newIdx : 0);
        SlidePanel.RefreshSingle(slideIdx);
        UpdateTitle();
        UpdateActions();
        SetStatus(S("Msg_ShapeDrawn"));
    }

    // ── Group / Ungroup ───────────────────────────────────────────────

    private void OnGroup(object? _ = null, System.Windows.RoutedEventArgs? __ = null)
    {
        if (!_model.IsOpen) return;
        var indices = EditorCanvas.SelectedTreeIndices;
        if (indices.Length < 2) return;
        OnShapesGroupRequested(_currentSlide, indices);
    }

    private void OnUngroup(object? _ = null, System.Windows.RoutedEventArgs? __ = null)
    {
        if (!_model.IsOpen) return;
        int treeIdx = EditorCanvas.SelectedTreeIndex;
        if (treeIdx < 0 || !_model.IsGroupShape(_currentSlide, treeIdx)) return;
        OnShapeUngroupRequested(_currentSlide, treeIdx);
    }

    private void OnShapesGroupRequested(int slideIdx, int[] treeIndices)
    {
        if (treeIndices.Length < 2) return;
        int newIdx = _model.GroupShapes(slideIdx, treeIndices);
        EditorCanvas.Invalidate(newIdx >= 0 ? newIdx : 0);
        SlidePanel.RefreshSingle(slideIdx);
        UpdateTitle();
        UpdateActions();
        SetStatus(S("Msg_Grouped"));
    }

    private void OnShapeUngroupRequested(int slideIdx, int treeIdx)
    {
        int[] newIndices = _model.UngroupShape(slideIdx, treeIdx);
        int sel = newIndices.Length > 0 ? newIndices[0] : 0;
        EditorCanvas.Invalidate(sel);
        SlidePanel.RefreshSingle(slideIdx);
        UpdateTitle();
        UpdateActions();
        SetStatus(S("Msg_Ungrouped"));
    }

    // ── Character properties ──────────────────────────────────────────

    private void OnCharProperties(object? _ = null)
    {
        if (!_model.IsOpen || !EditorCanvas.IsEditing) return;
        var style = EditorCanvas.GetEditorCharStyle();
        var dlg   = new PPEditer.Dialogs.CharPropertiesDialog(style) { Owner = this };
        EditorCanvas.SuppressLostFocusCommit = true;
        try { if (dlg.ShowDialog() != true) return; }
        finally { EditorCanvas.SuppressLostFocusCommit = false; }
        EditorCanvas.ApplyEditorCharStyle(dlg.Result);
    }

    // ── Paragraph properties ──────────────────────────────────────────

    private void OnParaProperties(object? _ = null)
    {
        if (!_model.IsOpen || !EditorCanvas.IsEditing) return;
        var style = EditorCanvas.GetEditorParaStyle();
        var dlg   = new PPEditer.Dialogs.ParaPropertiesDialog(style) { Owner = this };
        EditorCanvas.SuppressLostFocusCommit = true;
        try { if (dlg.ShowDialog() != true) return; }
        finally { EditorCanvas.SuppressLostFocusCommit = false; }
        EditorCanvas.ApplyEditorParaStyle(dlg.Result);
        // Apply VertAnchor to PPTX model immediately
        int treeIdx = EditorCanvas.SelectedTreeIndex;
        if (treeIdx >= 0)
            _model.UpdateBodyVertAnchor(_currentSlide, treeIdx, dlg.Result.VertAnchor);
        SetStatus(S("Msg_ParaPropsSaved"));
    }

    // ── Insert media ──────────────────────────────────────────────────

    private void OnInsertImage(object? _ = null)
    {
        if (!_model.IsOpen) return;
        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Filter = S("Dlg_ImageFilter"),
            Title  = S("Dlg_ImageTitle"),
        };
        if (dlg.ShowDialog(this) != true) return;
        try
        {
            _model.AddImage(_currentSlide, dlg.FileName);
            RefreshAll();
            SetStatus(S("Msg_ImageAdded"));
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"{S("Err_InsertFailed")}\n{ex.Message}",
                S("Lbl_InsertFailed"), MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnInsertVideo(object? _ = null)
    {
        if (!_model.IsOpen) return;
        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Filter = S("Dlg_VideoFilter"),
            Title  = S("Dlg_VideoTitle"),
        };
        if (dlg.ShowDialog(this) != true) return;
        try
        {
            _model.AddVideo(_currentSlide, dlg.FileName);
            RefreshAll();
            SetStatus(S("Msg_VideoAdded"));
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"{S("Err_InsertFailed")}\n{ex.Message}",
                S("Lbl_InsertFailed"), MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnInsertAudio(object? _ = null)
    {
        if (!_model.IsOpen) return;
        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Filter = S("Dlg_AudioFilter"),
            Title  = S("Dlg_AudioTitle"),
        };
        if (dlg.ShowDialog(this) != true) return;
        try
        {
            _model.AddAudio(_currentSlide, dlg.FileName);
            RefreshAll();
            SetStatus(S("Msg_AudioAdded"));
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"{S("Err_InsertFailed")}\n{ex.Message}",
                S("Lbl_InsertFailed"), MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // ── Insert text (CharMap / Emoji) ─────────────────────────────────

    private void OnInsertCharMap(object? _ = null)
    {
        if (!EditorCanvas.IsEditing) return;
        EditorCanvas.SuppressLostFocusCommit = true;
        var dlg = new PPEditer.Dialogs.CharMapDialog { Owner = this };
        dlg.CharInserted += ch => EditorCanvas.InsertText(ch);
        dlg.Closed += (_, _) =>
        {
            EditorCanvas.SuppressLostFocusCommit = false;
            _activeEditor?.Focus();
        };
        dlg.Show();
    }

    private void OnInsertEmoji(object? _ = null)
    {
        if (!EditorCanvas.IsEditing) return;
        EditorCanvas.SuppressLostFocusCommit = true;
        var dlg = new PPEditer.Dialogs.EmojiPickerDialog { Owner = this };
        dlg.EmojiInserted += em => EditorCanvas.InsertText(em);
        dlg.Closed += (_, _) =>
        {
            EditorCanvas.SuppressLostFocusCommit = false;
            _activeEditor?.Focus();
        };
        dlg.Show();
    }

    private void OnInsertMath(object? _ = null)
    {
        if (!EditorCanvas.IsEditing) return;
        EditorCanvas.SuppressLostFocusCommit = true;
        var dlg = new PPEditer.Dialogs.MathSymbolDialog { Owner = this };
        dlg.SymbolInserted += sym => EditorCanvas.InsertText(sym);
        dlg.Closed += (_, _) =>
        {
            EditorCanvas.SuppressLostFocusCommit = false;
            _activeEditor?.Focus();
        };
        dlg.Show();
    }

    private void OnEditorTextCommitted(int slideIdx, int treeIdx,
        IReadOnlyList<PptxConverter.PptxParagraph> paragraphs)
    {
        _model.UpdateShapeContent(slideIdx, treeIdx, paragraphs);
        // Rebuild canvas so the updated text is rendered (editor state is already null here).
        EditorCanvas.Invalidate(treeIdx);
        SlidePanel.RefreshSingle(slideIdx >= 0 ? slideIdx : _currentSlide);
        UpdateTitle();
        UpdateActions();
        FormatToolBar.IsEnabled = false;
        TbInsertTxBox.IsEnabled = _model.IsOpen;
        _activeEditor = null;
        SetStatus(S("Msg_TextSaved"));
    }

    private void Editor_SelectionChanged(object sender, RoutedEventArgs e)
        => SyncFormatToolbar();

    private void SyncFormatToolbar()
    {
        if (_activeEditor is null || _suppressFormatEvents) return;
        _suppressFormatEvents = true;
        try
        {
            var fw = EditorCanvas.GetSelectionProperty(TextElement.FontWeightProperty);
            TbBold.IsChecked = fw is FontWeight w && w == FontWeights.Bold;

            var fs = EditorCanvas.GetSelectionProperty(TextElement.FontStyleProperty);
            TbItalic.IsChecked = fs is FontStyle style && style == FontStyles.Italic;

            var td = EditorCanvas.GetSelectionProperty(Inline.TextDecorationsProperty);
            TbUnderline.IsChecked = td is TextDecorationCollection tdc && tdc.Count > 0;

            var ff = EditorCanvas.GetSelectionProperty(TextElement.FontFamilyProperty);
            if (ff is FontFamily family)
                FontFamilyCombo.Text = family.Source;

            var fsize = EditorCanvas.GetSelectionProperty(TextElement.FontSizeProperty);
            if (fsize is double sizePx && sizePx > 0)
                FontSizeCombo.Text = $"{sizePx * 72.0 / 96.0:0.##}";
        }
        finally
        {
            _suppressFormatEvents = false;
        }
    }

    // ── Formatting toolbar handlers ───────────────────────────────────

    private void FontFilterChanged(object sender, RoutedEventArgs e)
    {
        bool enabled = ChkOpenFontsOnly.IsChecked == true;
        MenuFontFilter.IsChecked = enabled;
        AppSettings.Current.FontLicenseFilterEnabled = enabled;
        AppSettings.Current.Save();
        PopulateFontCombo(enabled);
    }

    private void FontFamilyCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppressFormatEvents || !EditorCanvas.IsEditing) return;
        if (FontFamilyCombo.SelectedItem is FontInfo fi)
            EditorCanvas.ApplySelectionProperty(TextElement.FontFamilyProperty,
                new FontFamily(fi.Name));
    }

    private void FontFamilyCombo_LostFocus(object sender, RoutedEventArgs e)
    {
        if (_suppressFormatEvents || !EditorCanvas.IsEditing) return;
        var name = FontFamilyCombo.Text.Trim();
        if (!string.IsNullOrEmpty(name))
            EditorCanvas.ApplySelectionProperty(TextElement.FontFamilyProperty,
                new FontFamily(name));
    }

    private void FontSizeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppressFormatEvents || !EditorCanvas.IsEditing) return;
        if (FontSizeCombo.SelectedItem is ComboBoxItem item &&
            double.TryParse(item.Content?.ToString(), out double pt))
            EditorCanvas.ApplySelectionProperty(TextElement.FontSizeProperty,
                pt * 96.0 / 72.0);
    }

    private void FontSizeCombo_LostFocus(object sender, RoutedEventArgs e)
    {
        if (_suppressFormatEvents || !EditorCanvas.IsEditing) return;
        if (double.TryParse(FontSizeCombo.Text.TrimEnd("pt ".ToCharArray()), out double pt) && pt > 0)
            EditorCanvas.ApplySelectionProperty(TextElement.FontSizeProperty,
                pt * 96.0 / 72.0);
    }

    private void OnBoldClick(object sender, RoutedEventArgs e)
    {
        if (_activeEditor is null) return;
        EditingCommands.ToggleBold.Execute(null, _activeEditor);
        SyncFormatToolbar();
    }

    private void OnItalicClick(object sender, RoutedEventArgs e)
    {
        if (_activeEditor is null) return;
        EditingCommands.ToggleItalic.Execute(null, _activeEditor);
        SyncFormatToolbar();
    }

    private void OnUnderlineClick(object sender, RoutedEventArgs e)
    {
        if (_activeEditor is null) return;
        EditingCommands.ToggleUnderline.Execute(null, _activeEditor);
        SyncFormatToolbar();
    }

    private void OnTextColorClick(object sender, RoutedEventArgs e)
    {
        if (!EditorCanvas.IsEditing) return;

        var picker = new Window
        {
            Title  = S("Dlg_ColorPicker"),
            Width  = 228, Height = 180,
            Owner  = this,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ResizeMode = ResizeMode.NoResize,
        };
        var panel = new WrapPanel { Margin = new Thickness(8) };
        Color[] palette =
        [
            Colors.Black,    Colors.White,      Colors.Red,        Colors.Green,
            Colors.Blue,     Colors.Yellow,     Colors.Orange,     Colors.Purple,
            Colors.Cyan,     Colors.Magenta,    Colors.Gray,       Colors.DarkGray,
            Colors.LightGray,Colors.Brown,      Colors.Navy,       Colors.Teal,
            Colors.Olive,    Colors.Maroon,     Colors.DarkGreen,  Colors.Indigo,
        ];
        foreach (var c in palette)
        {
            var col = c;
            var btn = new Button
            {
                Width  = 28, Height = 28,
                Margin = new Thickness(2),
                Background = new SolidColorBrush(col),
                BorderBrush = new SolidColorBrush(Colors.Gray),
                ToolTip = $"#{col.R:X2}{col.G:X2}{col.B:X2}",
            };
            btn.Click += (_, _) =>
            {
                EditorCanvas.ApplySelectionProperty(TextElement.ForegroundProperty,
                    new SolidColorBrush(col));
                ColorBar.Fill = new SolidColorBrush(col);
                picker.Close();
            };
            panel.Children.Add(btn);
        }
        picker.Content = panel;
        picker.ShowDialog();
    }

    private void OnCommitEdit(object sender, RoutedEventArgs e)
    {
        EditorCanvas.CommitEdit(save: true);
        FormatToolBar.IsEnabled = false;
        TbInsertTxBox.IsEnabled = _model.IsOpen;
        _activeEditor = null;
    }

    // ── Settings handlers ─────────────────────────────────────────────

    private void OnFontFilterToggle(object sender, RoutedEventArgs e)
    {
        bool enabled = MenuFontFilter.IsChecked;
        ChkOpenFontsOnly.IsChecked = enabled;
        AppSettings.Current.FontLicenseFilterEnabled = enabled;
        AppSettings.Current.Save();
        PopulateFontCombo(enabled);
    }

    private void OnLangKo(object sender, RoutedEventArgs e)
    {
        AppLanguage.Apply("ko");
        AppSettings.Current.Language = "ko";
        AppSettings.Current.Save();
        SyncLangMenu("ko");
    }

    private void OnLangEn(object sender, RoutedEventArgs e)
    {
        AppLanguage.Apply("en");
        AppSettings.Current.Language = "en";
        AppSettings.Current.Save();
        SyncLangMenu("en");
    }

    private void OnThemeLight(object sender, RoutedEventArgs e)
    {
        AppTheme.Apply("Light");
        AppSettings.Current.Theme = "Light";
        AppSettings.Current.Save();
        SyncThemeMenu("Light");
    }

    private void OnThemeDark(object sender, RoutedEventArgs e)
    {
        AppTheme.Apply("Dark");
        AppSettings.Current.Theme = "Dark";
        AppSettings.Current.Save();
        SyncThemeMenu("Dark");
    }

    private void SyncLangMenu(string lang)
    {
        MenuLangKo.IsChecked = lang == "ko";
        MenuLangEn.IsChecked = lang == "en";
    }

    private void SyncThemeMenu(string theme)
    {
        MenuThemeLight.IsChecked = theme == "Light";
        MenuThemeDark.IsChecked  = theme == "Dark";
    }

    // ── Window events ─────────────────────────────────────────────────

    private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        if (!ConfirmDiscard()) e.Cancel = true;
    }

    protected override void OnClosed(EventArgs e)
    {
        base.OnClosed(e);
        _model.Dispose();
    }

    // ── Refresh helpers ────────────────────────────────────────────────

    private void RefreshAll()
    {
        SlidePanel.Refresh(_model, _currentSlide);
        ShowCurrentSlide();
        LoadCurrentNotes();
        UpdateTitle();
        UpdateActions();
        UpdateSlideInfo();
    }

    private void ShowCurrentSlide()
    {
        if (!_model.IsOpen) return;
        EditorCanvas.ShowSlide(_model, _currentSlide);
        SlidePanel.SetCurrent(_currentSlide);
        UpdateSlideInfo();
    }

    private void UpdateTitle() => Title = _model.IsOpen ? _model.WindowTitle : "PPEditer";

    private void UpdateSlideInfo()
    {
        SlideInfoText.Text = _model.IsOpen
            ? $"{_currentSlide + 1} / {_model.SlideCount}"
            : string.Empty;
    }

    private void UpdateZoomDisplay()
    {
        if (EditorCanvas is null) return;
        _suppressZoomEvents = true;
        int pct = (int)(EditorCanvas.ZoomFactor * 100);
        ZoomText.Text  = $"{pct}%";
        ZoomCombo.Text = $"{pct}%";
        _suppressZoomEvents = false;
    }

    private void UpdateActions()
    {
        bool has     = _model.IsOpen;
        bool editing = EditorCanvas.IsEditing;
        int  selIdx  = EditorCanvas.SelectedTreeIndex;
        bool hasMulti = has && EditorCanvas.SelectedTreeIndices.Length >= 2;
        bool isGroup  = has && selIdx >= 0 && _model.IsGroupShape(_currentSlide, selIdx);

        MenuSave.IsEnabled      = has && _model.Modified;
        MenuSaveAs.IsEnabled    = has;
        MenuPrint.IsEnabled     = has;
        MenuExportPdf.IsEnabled = has;
        MenuAddSlide.IsEnabled  = has;
        MenuDupSlide.IsEnabled  = has;
        MenuDelSlide.IsEnabled  = has && _model.SlideCount > 1;
        MenuSlideUp.IsEnabled   = has && _currentSlide > 0;
        MenuSlideDown.IsEnabled = has && _currentSlide < _model.SlideCount - 1;
        MenuSlideShow.IsEnabled      = has;
        MenuSlideTransition.IsEnabled = has;
        MenuShapeAnimation.IsEnabled  = has && EditorCanvas.SelectedTreeIndex >= 0;
        MenuUndo.IsEnabled      = _model.CanUndo;
        MenuRedo.IsEnabled      = _model.CanRedo;
        MenuInsertTextBox.IsEnabled  = has;
        MenuInsertImage.IsEnabled   = has;
        MenuInsertVideo.IsEnabled   = has;
        MenuInsertAudio.IsEnabled   = has;
        MenuInsertShape.IsEnabled   = has;
        MenuInsertMath.IsEnabled    = has && editing;
        MenuInsertCharMap.IsEnabled = has && editing;
        MenuInsertEmoji.IsEnabled   = has && editing;
        MenuDocInfo.IsEnabled       = has;

        TbSave.IsEnabled       = MenuSave.IsEnabled;
        TbUndo.IsEnabled       = MenuUndo.IsEnabled;
        TbRedo.IsEnabled       = MenuRedo.IsEnabled;
        TbAddSlide.IsEnabled   = has;
        TbDelSlide.IsEnabled   = MenuDelSlide.IsEnabled;
        TbExportPdf.IsEnabled  = has;
        TbInsertTxBox.IsEnabled = has && !editing;

        MenuGroup.IsEnabled     = hasMulti;
        MenuUngroup.IsEnabled   = isGroup;
        MenuCharProps.IsEnabled = has && editing;
        MenuParaProps.IsEnabled = has && editing;
    }

    private void SetStatus(string msg) => StatusMsg.Text = msg;

    private bool ConfirmDiscard()
    {
        if (!_model.IsOpen || !_model.Modified) return true;
        var r = MessageBox.Show(this,
            string.Format(S("Msg_UnsavedChanges"), _model.FileName),
            S("Lbl_UnsavedChanges"),
            MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
        if (r == MessageBoxResult.Yes) { OnSave(); return !_model.Modified; }
        return r == MessageBoxResult.No;
    }

    // ── Localisation helper ───────────────────────────────────────────

    private static string S(string key)
        => Application.Current.TryFindResource(key) is string s ? s : key;

    // ── Recent files ──────────────────────────────────────────────────

    private List<string> LoadRecent()
    {
        var val = Microsoft.Win32.Registry.CurrentUser
                      .OpenSubKey("Software")?.OpenSubKey("PPEditer")
                      ?.GetValue("RecentFiles") as string;
        if (string.IsNullOrEmpty(val)) return [];
        return [.. val.Split('|', StringSplitOptions.RemoveEmptyEntries)];
    }

    private void SaveRecent(List<string> list)
    {
        var key = Microsoft.Win32.Registry.CurrentUser
                     .CreateSubKey("Software\\PPEditer");
        key.SetValue("RecentFiles", string.Join("|", list));
    }

    private void AddRecent(string path)
    {
        var list = LoadRecent();
        list.Remove(path);
        list.Insert(0, path);
        SaveRecent([.. list.Take(RecentMax)]);
    }

    private void RebuildRecentMenu()
    {
        RecentMenu.Items.Clear();
        var list = LoadRecent();
        if (list.Count == 0)
        {
            RecentMenu.Items.Add(new MenuItem { Header = "(없음)", IsEnabled = false });
            return;
        }
        for (int i = 0; i < list.Count; i++)
        {
            var path = list[i];
            var item = new MenuItem
            {
                Header  = $"_{i + 1}. {Path.GetFileName(path)}",
                ToolTip = path,
            };
            item.Click += (_, _) =>
            {
                if (!File.Exists(path))
                {
                    MessageBox.Show(this, $"{S("Err_FileNotFound")}\n{path}",
                                    S("Lbl_FileNotFound"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                if (ConfirmDiscard()) OpenFile(path);
            };
            RecentMenu.Items.Add(item);
        }
    }

    // ── XAML event shims ──────────────────────────────────────────────

    private void OnNew(object s, RoutedEventArgs e)           => OnNew();
    private void OnOpen(object s, RoutedEventArgs e)          => OnOpen();
    private void OnSave(object s, RoutedEventArgs e)          => OnSave();
    private void OnSaveAs(object s, RoutedEventArgs e)        => OnSaveAs();
    private void OnExportPdf(object s, RoutedEventArgs e)     => OnExportPdf();
    private void OnExit(object s, RoutedEventArgs e)          => OnExit();
    private void OnUndo(object s, RoutedEventArgs e)          => OnUndo();
    private void OnRedo(object s, RoutedEventArgs e)          => OnRedo();
    // ── Effects ───────────────────────────────────────────────────────

    private void OnSlideTransition(object? _ = null)
    {
        if (!_model.IsOpen) return;
        var current = _model.GetSlideTransition(_currentSlide);
        var dlg     = new Dialogs.TransitionDialog(current) { Owner = this };
        if (dlg.ShowDialog() != true) return;
        _model.SetSlideTransition(_currentSlide,
            new Models.SlideTransition { Kind = dlg.SelectedKind, DurationMs = dlg.DurationSeconds * 1000 },
            dlg.ApplyToAll);
        UpdateActions();
        SetStatus(S("Msg_TransitionSet"));
    }

    private void OnShapeAnimation(object? _ = null)
    {
        if (!_model.IsOpen) return;
        int treeIdx = EditorCanvas.SelectedTreeIndex;
        if (treeIdx < 0) return;
        var current = _model.GetShapeAnimation(_currentSlide, treeIdx);
        var dlg     = new Dialogs.AnimationDialog(current) { Owner = this };
        if (dlg.ShowDialog() != true) return;
        _model.SetShapeAnimation(_currentSlide, treeIdx, dlg.SelectedKind,
            dlg.DurationSeconds * 1000, dlg.AutoPlay, dlg.RepeatCount);
        UpdateActions();
        SetStatus(S("Msg_AnimationSet"));
    }

    // ── Slide show ────────────────────────────────────────────────────

    private void OnSlideShow(object? _ = null)
    {
        if (!_model.IsOpen) return;
        SaveCurrentNotes();
        EditorCanvas.CommitEdit(save: true);

        var showWnd = new Dialogs.SlideShowWindow(_model, _currentSlide);
        var presWnd = new Dialogs.PresenterViewWindow(_model, _currentSlide);
        presWnd.AttachShowWindow(showWnd);

        // Close both when either is closed
        showWnd.Closed += (_, _) => { try { presWnd.Close(); } catch { } };
        presWnd.Closed += (_, _) => { try { if (showWnd.IsLoaded) showWnd.Close(); } catch { } };

        PositionPresenterView(showWnd, presWnd);

        presWnd.Show();
        showWnd.ShowDialog();
    }

    private void PositionPresenterView(Dialogs.SlideShowWindow showWnd,
                                       Dialogs.PresenterViewWindow presWnd)
    {
        var monitors = Services.ScreenHelper.GetMonitors();
        var s        = Services.AppSettings.Current;

        if (monitors.Count >= 2)
        {
            // Resolve show monitor: saved preference, or primary (index 0 is primary by EnumDisplayMonitors order).
            int primaryIdx = monitors.Select((m, i) => (m, i)).FirstOrDefault(x => x.m.IsPrimary).i;
            int showIdx = (s.ShowMonitorIndex >= 0 && s.ShowMonitorIndex < monitors.Count)
                          ? s.ShowMonitorIndex : primaryIdx;

            // Resolve presenter monitor: saved preference if different, otherwise the other monitor.
            int presIdx = (s.PresenterMonitorIndex >= 0 && s.PresenterMonitorIndex < monitors.Count
                           && s.PresenterMonitorIndex != showIdx)
                          ? s.PresenterMonitorIndex
                          : showIdx == 0 ? 1 : 0;

            Services.ScreenHelper.MaximizeOnMonitor(showWnd, monitors, showIdx);
            Services.ScreenHelper.MaximizeOnMonitor(presWnd, monitors, presIdx);
            return;
        }

        // Single monitor — maximize show, presenter as floating window.
        showWnd.WindowState = WindowState.Maximized;
        presWnd.WindowStartupLocation = WindowStartupLocation.Manual;
        presWnd.Width  = 780;
        presWnd.Height = 540;
        presWnd.Left   = 30;
        presWnd.Top    = 30;
        presWnd.WindowState = WindowState.Normal;
    }

    private void OnDisplaySettings(object sender, RoutedEventArgs e)
    {
        var dlg = new Dialogs.DisplaySettingsDialog { Owner = this };
        dlg.ShowDialog();
    }

    // ── XAML event bridges ────────────────────────────────────────────

    private void OnSlideTransition(object s, RoutedEventArgs e) => OnSlideTransition();
    private void OnShapeAnimation(object s, RoutedEventArgs e)  => OnShapeAnimation();
    private void OnSlideShow(object s, RoutedEventArgs e)       => OnSlideShow();
    private void OnAddSlide(object s, RoutedEventArgs e)      => OnAddSlide();
    private void OnDupSlide(object s, RoutedEventArgs e)      => OnDupSlide();
    private void OnDelSlide(object s, RoutedEventArgs e)      => OnDelSlide();
    private void OnSlideUp(object s, RoutedEventArgs e)       => OnSlideUp();
    private void OnSlideDown(object s, RoutedEventArgs e)     => OnSlideDown();
    private void OnZoomIn(object s, RoutedEventArgs e)        => OnZoomIn();
    private void OnZoomOut(object s, RoutedEventArgs e)       => OnZoomOut();
    private void OnZoomFit(object s, RoutedEventArgs e)       => OnZoomFit();
    private void OnTogglePanel(object s, RoutedEventArgs e)   => OnTogglePanel();
    private void OnToggleStatus(object s, RoutedEventArgs e)  => OnToggleStatus();
    private void OnAbout(object s, RoutedEventArgs e)         => OnAbout();
    private void OnInsertTextBox(object s, RoutedEventArgs e) => OnInsertTextBox();
    private void OnDocInfo(object s, RoutedEventArgs e)       => OnDocInfo();
    private void OnInsertImage(object s, RoutedEventArgs e)   => OnInsertImage();
    private void OnInsertVideo(object s, RoutedEventArgs e)   => OnInsertVideo();
    private void OnInsertAudio(object s, RoutedEventArgs e)   => OnInsertAudio();
    private void OnInsertMath(object s, RoutedEventArgs e)     => OnInsertMath();
    private void OnInsertCharMap(object s, RoutedEventArgs e) => OnInsertCharMap();
    private void OnInsertEmoji(object s, RoutedEventArgs e)   => OnInsertEmoji();
    private void OnCharProperties(object s, RoutedEventArgs e) => OnCharProperties();
    private void OnParaProperties(object s, RoutedEventArgs e) => OnParaProperties();
    private void OnParaPropsClick(object s, RoutedEventArgs e) => OnParaProperties();
}

// ── Tiny relay command ─────────────────────────────────────────────────────────

internal sealed class RelayCommand(Action<object?> execute) : ICommand
{
    public event EventHandler? CanExecuteChanged { add { } remove { } }
    public bool CanExecute(object? p) => true;
    public void Execute(object? p)    => execute(p);
}
