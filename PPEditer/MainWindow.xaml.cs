using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using PPEditer.Dialogs;
using PPEditer.Export;
using PPEditer.Models;

namespace PPEditer;

public partial class MainWindow : Window
{
    private readonly PresentationModel _model = new();
    private int  _currentSlide;
    private bool _suppressZoomEvents;

    // ── Recent files (stored in registry via IsolatedStorage-free approach) ──
    private const int RecentMax = 10;
    private static readonly string RecentKey = "PPEditer_RecentFiles";

    public MainWindow()
    {
        InitializeComponent();

        SlidePanel.SlideSelected    += idx => { _currentSlide = idx; ShowCurrentSlide(); UpdateActions(); };
        EditorCanvas.ShapeTextEdited += OnShapeTextEdited;

        RegisterKeyBindings();
        RebuildRecentMenu();
        UpdateActions();
    }

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
    }

    // ── File commands ─────────────────────────────────────────────────

    private void OnNew(object? _ = null)
    {
        if (!ConfirmDiscard()) return;
        _model.New();
        _currentSlide = 0;
        RefreshAll();
        SetStatus("새 프레젠테이션을 만들었습니다.");
    }

    private void OnOpen(object? _ = null)
    {
        if (!ConfirmDiscard()) return;
        var dlg = new OpenFileDialog
        {
            Filter = "프레젠테이션 파일|*.pptx;*.ppt|모든 파일|*.*",
            Title  = "파일 열기",
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
            SetStatus($"열었습니다: {Path.GetFileName(path)}");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"파일을 열 수 없습니다:\n{ex.Message}",
                            "열기 오류", MessageBoxButton.OK, MessageBoxImage.Error);
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
            SetStatus("저장되었습니다.");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"저장 실패:\n{ex.Message}", "저장 오류",
                            MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnSaveAs(object? _ = null)
    {
        if (!_model.IsOpen) return;
        var dlg = new SaveFileDialog
        {
            Filter           = "PowerPoint 파일|*.pptx|모든 파일|*.*",
            Title            = "다른 이름으로 저장",
            DefaultExt       = "pptx",
            FileName         = _model.FileName,
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
            SetStatus($"저장되었습니다: {Path.GetFileName(path)}");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"저장 실패:\n{ex.Message}", "저장 오류",
                            MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnExportPdf(object? _ = null)
    {
        if (!_model.IsOpen) return;
        var defaultName = Path.ChangeExtension(_model.FileName, ".pdf");
        var dlg = new SaveFileDialog
        {
            Filter     = "PDF 파일|*.pdf|모든 파일|*.*",
            Title      = "PDF로 내보내기",
            DefaultExt = "pdf",
            FileName   = defaultName,
        };
        if (dlg.ShowDialog(this) != true) return;

        var progressDlg = new Window
        {
            Title  = "내보내기",
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
                progressDlg.Title = $"내보내기... {p.c}/{p.t}";
            });
            bool ok = PdfExporter.Export(_model, dlg.FileName, progress);
            progressDlg.Close();
            if (ok)
                SetStatus($"PDF 저장 완료: {Path.GetFileName(dlg.FileName)}");
            else
                MessageBox.Show(this, "PDF 내보내기에 실패했습니다.", "오류",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        catch (Exception ex)
        {
            progressDlg.Close();
            MessageBox.Show(this, $"PDF 내보내기 실패:\n{ex.Message}", "오류",
                            MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnExit(object? _ = null) => Close();

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
            MessageBox.Show(this, "마지막 슬라이드는 삭제할 수 없습니다.",
                            "삭제 불가", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        var r = MessageBox.Show(this,
            $"슬라이드 {_currentSlide + 1}을(를) 삭제하시겠습니까?",
            "슬라이드 삭제", MessageBoxButton.YesNo, MessageBoxImage.Question);
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

    // ── View commands ─────────────────────────────────────────────────

    private void OnZoomIn(object?  _ = null) { EditorCanvas.ZoomIn();       UpdateZoomDisplay(); }
    private void OnZoomOut(object? _ = null) { EditorCanvas.ZoomOut();      UpdateZoomDisplay(); }
    private void OnZoomFit(object? _ = null) { EditorCanvas.FitToWindow();  UpdateZoomDisplay(); }

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
        if (double.TryParse(text.TrimEnd('%'), out double pct))
            EditorCanvas.SetZoom(pct / 100.0);   // We'll add SetZoom below
        UpdateZoomDisplay();
    }

    // ── Shape text edited ─────────────────────────────────────────────

    private void OnShapeTextEdited(int shapeIdx, string newText)
    {
        // Note: text editing of PPTX shapes via OpenXml would require re-opening
        // the SlidePart XML — this is a future enhancement.
        // For now, mark as modified so the user can save the visual state.
        UpdateTitle();
        SetStatus("텍스트가 편집되었습니다. (저장하려면 Ctrl+S)");
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
            ? $"슬라이드 {_currentSlide + 1} / {_model.SlideCount}"
            : string.Empty;
    }

    private void UpdateZoomDisplay()
    {
        _suppressZoomEvents = true;
        int pct = (int)(EditorCanvas.ZoomFactor * 100);
        ZoomText.Text   = $"{pct}%";
        ZoomCombo.Text  = $"{pct}%";
        _suppressZoomEvents = false;
    }

    private void UpdateActions()
    {
        bool has = _model.IsOpen;
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

        TbSave.IsEnabled       = MenuSave.IsEnabled;
        TbUndo.IsEnabled       = MenuUndo.IsEnabled;
        TbRedo.IsEnabled       = MenuRedo.IsEnabled;
        TbAddSlide.IsEnabled   = has;
        TbDelSlide.IsEnabled   = MenuDelSlide.IsEnabled;
        TbExportPdf.IsEnabled  = has;
    }

    private void SetStatus(string msg) => StatusMsg.Text = msg;

    private bool ConfirmDiscard()
    {
        if (!_model.IsOpen || !_model.Modified) return true;
        var r = MessageBox.Show(this,
            $"'{_model.FileName}'의 변경사항을 저장하시겠습니까?",
            "저장되지 않은 변경사항",
            MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
        if (r == MessageBoxResult.Yes) { OnSave(); return !_model.Modified; }
        return r == MessageBoxResult.No;
    }

    // ── Recent files ──────────────────────────────────────────────────

    private List<string> LoadRecent()
    {
        var val = Microsoft.Win32.Registry.CurrentUser
                      .OpenSubKey("Software")?.OpenSubKey("PPEditer")
                      ?.GetValue("RecentFiles") as string;
        if (string.IsNullOrEmpty(val)) return new();
        return val.Split('|', StringSplitOptions.RemoveEmptyEntries).ToList();
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
        SaveRecent(list.Take(RecentMax).ToList());
    }

    private void RebuildRecentMenu()
    {
        RecentMenu.Items.Clear();
        var list = LoadRecent();
        if (list.Count == 0)
        {
            var empty = new MenuItem { Header = "(없음)", IsEnabled = false };
            RecentMenu.Items.Add(empty);
            return;
        }
        for (int i = 0; i < list.Count; i++)
        {
            var path = list[i];
            var item = new MenuItem
            {
                Header      = $"_{i + 1}. {Path.GetFileName(path)}",
                ToolTip     = path,
            };
            item.Click += (_, _) =>
            {
                if (!File.Exists(path))
                {
                    MessageBox.Show(this, $"파일을 찾을 수 없습니다:\n{path}",
                                    "파일 없음", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                if (ConfirmDiscard()) OpenFile(path);
            };
            RecentMenu.Items.Add(item);
        }
    }

    // ── Event handler wrappers (XAML Click → object? handler) ─────────

    private void OnNew(object s, RoutedEventArgs e)       => OnNew();
    private void OnOpen(object s, RoutedEventArgs e)      => OnOpen();
    private void OnSave(object s, RoutedEventArgs e)      => OnSave();
    private void OnSaveAs(object s, RoutedEventArgs e)    => OnSaveAs();
    private void OnExportPdf(object s, RoutedEventArgs e) => OnExportPdf();
    private void OnExit(object s, RoutedEventArgs e)      => OnExit();
    private void OnUndo(object s, RoutedEventArgs e)      => OnUndo();
    private void OnRedo(object s, RoutedEventArgs e)      => OnRedo();
    private void OnAddSlide(object s, RoutedEventArgs e)  => OnAddSlide();
    private void OnDupSlide(object s, RoutedEventArgs e)  => OnDupSlide();
    private void OnDelSlide(object s, RoutedEventArgs e)  => OnDelSlide();
    private void OnSlideUp(object s, RoutedEventArgs e)   => OnSlideUp();
    private void OnSlideDown(object s, RoutedEventArgs e) => OnSlideDown();
    private void OnZoomIn(object s, RoutedEventArgs e)    => OnZoomIn();
    private void OnZoomOut(object s, RoutedEventArgs e)   => OnZoomOut();
    private void OnZoomFit(object s, RoutedEventArgs e)   => OnZoomFit();
    private void OnTogglePanel(object s, RoutedEventArgs e)  => OnTogglePanel();
    private void OnToggleStatus(object s, RoutedEventArgs e) => OnToggleStatus();
    private void OnAbout(object s, RoutedEventArgs e)     => OnAbout();
}

// ── Tiny relay command ─────────────────────────────────────────────────────────

internal sealed class RelayCommand(Action<object?> execute) : ICommand
{
    public event EventHandler? CanExecuteChanged;
    public bool CanExecute(object? p) => true;
    public void Execute(object? p)    => execute(p);
}
