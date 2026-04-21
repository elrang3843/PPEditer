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
    private RichTextBox? _activeEditor;

    private const int RecentMax = 10;

    public MainWindow()
    {
        InitializeComponent();

        SlidePanel.SlideSelected   += idx => { _currentSlide = idx; ShowCurrentSlide(); UpdateActions(); };
        EditorCanvas.EditingStarted += OnEditorEditingStarted;
        EditorCanvas.TextCommitted  += OnEditorTextCommitted;

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
        MenuDocPassword.IsChecked  = s.DocumentPasswordEnabled;
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
            ? FontService.GetRecommendedFonts()
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
        long cx = _model.SlideWidth;
        long cy = _model.SlideHeight;
        long w  = 4572000L;            // 5 in
        long h  = 914400L;             // 1 in
        long l  = (cx - w) / 2;
        long t  = (cy - h) / 2;
        _model.AddTextBox(_currentSlide, l, t, w, h);
        RefreshAll();
        SetStatus(S("Msg_TextBoxAdded"));
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

    private void OnEditorTextCommitted(int slideIdx, int treeIdx,
        IReadOnlyList<PptxConverter.PptxParagraph> paragraphs)
    {
        _model.UpdateShapeContent(slideIdx, treeIdx, paragraphs);
        SlidePanel.RefreshSingle(treeIdx >= 0 ? slideIdx : _currentSlide);
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

    private void OnDocPasswordToggle(object sender, RoutedEventArgs e)
    {
        if (MenuDocPassword.IsChecked)
        {
            if (!PromptNewPassword())
                MenuDocPassword.IsChecked = false;
        }
        else
        {
            AppSettings.Current.DocumentPasswordEnabled = false;
            AppSettings.Current.DocumentPassword = string.Empty;
            AppSettings.Current.Save();
        }
    }

    private bool PromptNewPassword()
    {
        var dlg = new Window
        {
            Title  = S("Dlg_Password"),
            Width  = 320, Height = 200,
            Owner  = this,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ResizeMode = ResizeMode.NoResize,
        };
        var grid = new Grid { Margin = new Thickness(16) };
        for (int i = 0; i < 5; i++)
            grid.RowDefinitions.Add(new RowDefinition
                { Height = i == 3 ? new GridLength(1, GridUnitType.Star) : GridLength.Auto });

        var lbl1   = new TextBlock { Text = S("Dlg_PasswordPrompt"),  Margin = new Thickness(0, 0, 0, 4) };
        var pw1    = new PasswordBox { Margin = new Thickness(0, 0, 0, 8) };
        var lbl2   = new TextBlock { Text = S("Dlg_PasswordConfirm"), Margin = new Thickness(0, 0, 0, 4) };
        var pw2    = new PasswordBox { Margin = new Thickness(0, 0, 0, 8) };

        var btnRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
        };
        var btnOk     = new Button { Content = S("Btn_Ok"),     Width = 72, IsDefault = true, Margin = new Thickness(0, 0, 8, 0) };
        var btnCancel = new Button { Content = S("Btn_Cancel"), Width = 72, IsCancel  = true };
        btnRow.Children.Add(btnOk);
        btnRow.Children.Add(btnCancel);

        Grid.SetRow(lbl1, 0); Grid.SetRow(pw1, 1);
        Grid.SetRow(lbl2, 2); Grid.SetRow(pw2, 3);
        Grid.SetRow(btnRow, 4);
        foreach (UIElement el in new UIElement[] { lbl1, pw1, lbl2, pw2, btnRow })
            grid.Children.Add(el);
        dlg.Content = grid;

        bool ok = false;
        btnOk.Click += (_, _) =>
        {
            if (pw1.Password != pw2.Password)
            {
                MessageBox.Show(dlg, S("Dlg_PasswordMismatch"), S("Dlg_Password"),
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            AppSettings.Current.DocumentPasswordEnabled = true;
            AppSettings.Current.DocumentPassword = pw1.Password;
            AppSettings.Current.Save();
            ok = true;
            dlg.Close();
        };
        btnCancel.Click += (_, _) => dlg.Close();
        dlg.ShowDialog();
        return ok;
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
        bool has = _model.IsOpen;
        bool editing = EditorCanvas.IsEditing;

        MenuSave.IsEnabled      = has && _model.Modified;
        MenuSaveAs.IsEnabled    = has;
        MenuExportPdf.IsEnabled = has;
        MenuAddSlide.IsEnabled  = has;
        MenuDupSlide.IsEnabled  = has;
        MenuDelSlide.IsEnabled  = has && _model.SlideCount > 1;
        MenuSlideUp.IsEnabled   = has && _currentSlide > 0;
        MenuSlideDown.IsEnabled = has && _currentSlide < _model.SlideCount - 1;
        MenuUndo.IsEnabled      = _model.CanUndo;
        MenuRedo.IsEnabled      = _model.CanRedo;
        MenuInsertTextBox.IsEnabled = has;
        MenuDocInfo.IsEnabled       = has;

        TbSave.IsEnabled       = MenuSave.IsEnabled;
        TbUndo.IsEnabled       = MenuUndo.IsEnabled;
        TbRedo.IsEnabled       = MenuRedo.IsEnabled;
        TbAddSlide.IsEnabled   = has;
        TbDelSlide.IsEnabled   = MenuDelSlide.IsEnabled;
        TbExportPdf.IsEnabled  = has;
        TbInsertTxBox.IsEnabled = has && !editing;
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
}

// ── Tiny relay command ─────────────────────────────────────────────────────────

internal sealed class RelayCommand(Action<object?> execute) : ICommand
{
    public event EventHandler? CanExecuteChanged { add { } remove { } }
    public bool CanExecute(object? p) => true;
    public void Execute(object? p)    => execute(p);
}
