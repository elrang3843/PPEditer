using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PPEditer.Models;
using PPEditer.Rendering;
using PPEditer.Services;

namespace PPEditer.Controls;

public partial class SlideEditorCanvas : UserControl
{
    // ── Events ─────────────────────────────────────────────────────────

    /// <summary>Text editing session started. Payload = the active RichTextBox.</summary>
    public event Action<RichTextBox>? EditingStarted;
    /// <summary>User committed a text edit. (slideIdx, shapeTreeIdx, paragraphs)</summary>
    public event Action<int, int, IReadOnlyList<PptxConverter.PptxParagraph>>? TextCommitted;
    /// <summary>User moved a shape with arrow keys or mouse drag. (slideIdx, treeIdx, dxEmu, dyEmu)</summary>
    public event Action<int, int, long, long>? ShapeMoved;
    /// <summary>User deleted the selected shape. (slideIdx, treeIdx)</summary>
    public event Action<int, int>? ShapeDeleted;
    /// <summary>User resized the selected shape. (slideIdx, treeIdx, leftEmu, topEmu, widthEmu, heightEmu)</summary>
    public event Action<int, int, long, long, long, long>? ShapeResized;
    /// <summary>User requested object properties via context menu. (slideIdx, treeIdx)</summary>
    public event Action<int, int>? ShapePropertiesRequested;
    /// <summary>User changed z-order via context menu. (slideIdx, treeIdx, delta: +1=forward/-1=backward)</summary>
    public event Action<int, int, int>? ShapeOrderChanged;

    // ── Public property ────────────────────────────────────────────────

    /// <summary>Set true before opening a dialog so LostFocus does not commit and close the editor.</summary>
    public bool SuppressLostFocusCommit { get; set; }

    // ── State ──────────────────────────────────────────────────────────

    private PresentationModel? _model;
    private SlidePart?         _slidePart;
    private int                _slideIndex;
    private Canvas?            _nativeCanvas;
    private double             _zoom = 1.0;

    private int      _selectedIdx     = -1;   // canvas child index
    private int      _selectedTreeIdx = -1;   // shape-tree index
    private Canvas?  _selOverlay;             // selection overlay (border + 8 handles)
    private RichTextBox? _editor;
    private int      _editingTreeIdx  = -1;

    // ── Drag state ─────────────────────────────────────────────────────

    private enum DragMode { None, Move, Resize }
    private DragMode _dragMode     = DragMode.None;
    private int      _resizeHandle = -1;       // 0-7 (TL TC TR MR BR BC BL ML)
    private Point    _dragStart;               // position in native-canvas coords
    private double   _dOrigL, _dOrigT, _dOrigW, _dOrigH; // original px bounds

    private const double HandleSize = 9.0;
    private const double MinShapeSize = 20.0;

    // ── Native canvas dimensions ───────────────────────────────────────

    private double NativeW => _model is not null ? _model.SlideWidth  / 914400.0 * 96.0 : 960;
    private double NativeH => _model is not null ? _model.SlideHeight / 914400.0 * 96.0 : 540;
    private double EmuPerPx => _model is not null ? (double)_model.SlideWidth / NativeW : 9525.0;

    // ── Constructor ────────────────────────────────────────────────────

    public SlideEditorCanvas()
    {
        InitializeComponent();
        Loaded      += (_, _) => UpdateViewboxSize();
        SizeChanged += (_, _) => { if (_zoom == 1.0) UpdateViewboxSize(); };
    }

    // ── Public API ─────────────────────────────────────────────────────

    public void ShowSlide(PresentationModel model, int slideIndex)
    {
        CommitEdit(save: false);
        _model      = model;
        _slideIndex = slideIndex;
        _slidePart  = model.GetSlidePart(slideIndex);
        Rebuild();
    }

    public void Invalidate(bool preserveSelection = true)
    {
        int savedTreeIdx = preserveSelection ? _selectedTreeIdx : -1;
        CommitEdit(save: false);
        Rebuild();
        if (savedTreeIdx >= 0 && _nativeCanvas is not null)
            SelectShapeByTreeIndex(savedTreeIdx);
    }

    public void Invalidate(int treeIdxToSelect)
    {
        CommitEdit(save: false);
        Rebuild();
        if (treeIdxToSelect >= 0 && _nativeCanvas is not null)
            SelectShapeByTreeIndex(treeIdxToSelect);
    }

    public void ZoomIn()      { _zoom = Math.Min(3.0,  Math.Round(_zoom + 0.1, 1)); ApplyZoom(); }
    public void ZoomOut()     { _zoom = Math.Max(0.25, Math.Round(_zoom - 0.1, 1)); ApplyZoom(); }
    public void FitToWindow() { _zoom = 1.0; UpdateViewboxSize(); }
    public void SetZoom(double z) { _zoom = Math.Clamp(z, 0.25, 3.0); ApplyZoom(); }
    public double ZoomFactor  => _zoom;
    public bool   IsEditing   => _editor is not null;

    /// <summary>Insert text at the active editor's caret (used by CharMap / Emoji dialogs).</summary>
    public void InsertText(string text)
    {
        if (_editor is null) return;
        try
        {
            if (!_editor.Selection.IsEmpty)
                _editor.Selection.Text = "";
            _editor.CaretPosition.InsertTextInRun(text);
            var newPos = _editor.CaretPosition.GetPositionAtOffset(text.Length);
            if (newPos != null) _editor.CaretPosition = newPos;
        }
        catch
        {
            _editor.Selection.Text = text;
        }
        _editor.Focus();
    }

    // ── Rebuild ────────────────────────────────────────────────────────

    private void Rebuild()
    {
        _selectedIdx     = -1;
        _selectedTreeIdx = -1;
        _selOverlay      = null;
        _editingTreeIdx  = -1;
        _dragMode        = DragMode.None;
        _resizeHandle    = -1;
        NoDocMsg.Visibility = Visibility.Collapsed;

        if (_model is null || !_model.IsOpen || _slidePart is null)
        {
            SlideViewbox.Child  = new Canvas { Width = 960, Height = 540 };
            _nativeCanvas       = null;
            NoDocMsg.Visibility = Visibility.Visible;
            return;
        }

        var canvas = SlideRenderer.BuildCanvas(_slidePart, _model.SlideWidth, _model.SlideHeight);
        canvas.MouseLeftButtonDown  += Canvas_MouseDown;
        canvas.MouseMove            += Canvas_MouseMove;
        canvas.MouseLeftButtonUp    += Canvas_MouseUp;
        canvas.MouseRightButtonUp   += Canvas_RightMouseUp;
        canvas.KeyDown              += Canvas_KeyDown;
        canvas.Focusable             = true;

        _nativeCanvas      = canvas;
        SlideViewbox.Child = canvas;
        UpdateViewboxSize();
    }

    // ── Zoom & sizing ──────────────────────────────────────────────────

    private void ApplyZoom() => UpdateViewboxSize();

    private void UpdateViewboxSize()
    {
        double aspect = NativeW / Math.Max(1, NativeH);

        if (_zoom == 1.0)
        {
            double availW = Math.Max(200, Scroller.ActualWidth  - 80);
            double availH = Math.Max(120, Scroller.ActualHeight - 80);
            double w = availW / aspect <= availH ? availW : availH * aspect;
            double h = w / aspect;
            SlideViewbox.Width  = w;
            SlideViewbox.Height = h;
        }
        else
        {
            SlideViewbox.Width  = NativeW * _zoom;
            SlideViewbox.Height = NativeH * _zoom;
        }
    }

    // ── Mouse handling ─────────────────────────────────────────────────

    private void Canvas_MouseDown(object sender, MouseButtonEventArgs e)
    {
        var canvas   = (Canvas)sender;
        var clickPos = e.GetPosition(canvas);

        if (e.ClickCount == 1)
        {
            CommitEdit(save: true);
            int idx = HitTest(canvas, clickPos);
            SelectShape(canvas, idx);
        }
        else if (e.ClickCount == 2)
        {
            int idx = HitTest(canvas, clickPos);
            if (idx >= 0) StartEdit(canvas, idx);
        }
        canvas.Focus();
        e.Handled = true;
    }

    private void Canvas_MouseMove(object sender, MouseEventArgs e)
    {
        if (_dragMode == DragMode.None || _selOverlay is null || _nativeCanvas is null) return;

        var pos = e.GetPosition(_nativeCanvas);
        double dx = pos.X - _dragStart.X;
        double dy = pos.Y - _dragStart.Y;

        if (_dragMode == DragMode.Move)
        {
            Canvas.SetLeft(_selOverlay, _dOrigL + dx);
            Canvas.SetTop(_selOverlay,  _dOrigT + dy);
        }
        else if (_dragMode == DragMode.Resize)
        {
            var (l, t, w, h) = ComputeResizeBounds(dx, dy);
            UpdateOverlayBounds(l, t, w, h);
        }
    }

    private void Canvas_MouseUp(object sender, MouseButtonEventArgs e)
    {
        if (_dragMode == DragMode.None || _nativeCanvas is null) return;

        _nativeCanvas.ReleaseMouseCapture();

        var pos = e.GetPosition(_nativeCanvas);
        double dx = pos.X - _dragStart.X;
        double dy = pos.Y - _dragStart.Y;

        if (_dragMode == DragMode.Move && _selectedTreeIdx >= 0)
        {
            long dxEmu = (long)(dx * EmuPerPx);
            long dyEmu = (long)(dy * EmuPerPx);
            if (dxEmu != 0 || dyEmu != 0)
                ShapeMoved?.Invoke(_slideIndex, _selectedTreeIdx, dxEmu, dyEmu);
        }
        else if (_dragMode == DragMode.Resize && _selectedTreeIdx >= 0)
        {
            var (l, t, w, h) = ComputeResizeBounds(dx, dy);
            if (w >= MinShapeSize && h >= MinShapeSize)
                ShapeResized?.Invoke(_slideIndex, _selectedTreeIdx,
                    PxToEmu(l), PxToEmu(t), PxToEmu(w), PxToEmu(h));
        }

        _dragMode     = DragMode.None;
        _resizeHandle = -1;
        e.Handled     = true;
    }

    private void Canvas_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
        {
            CommitEdit(save: true);
            SelectShape((Canvas)sender, -1);
            e.Handled = true;
        }
        else if (e.Key == Key.Delete && _selectedIdx >= 0 && _editor is null)
        {
            if (_selectedTreeIdx >= 0)
                ShapeDeleted?.Invoke(_slideIndex, _selectedTreeIdx);
            e.Handled = true;
        }
        else if (_selectedIdx >= 0 && _editor is null &&
                 (e.Key == Key.Left || e.Key == Key.Right ||
                  e.Key == Key.Up   || e.Key == Key.Down))
        {
            bool fine = (Keyboard.Modifiers & ModifierKeys.Shift) != 0;
            long step = fine ? 7200L : 72000L;
            long kDx = e.Key == Key.Left  ? -step : e.Key == Key.Right ? step : 0L;
            long kDy = e.Key == Key.Up    ? -step : e.Key == Key.Down  ? step : 0L;
            if (_selectedTreeIdx >= 0)
                ShapeMoved?.Invoke(_slideIndex, _selectedTreeIdx, kDx, kDy);
            e.Handled = true;
        }
    }

    private void Canvas_RightMouseUp(object sender, MouseButtonEventArgs e)
    {
        var canvas   = (Canvas)sender;
        var clickPos = e.GetPosition(canvas);
        int idx = HitTest(canvas, clickPos);
        if (idx >= 0)
        {
            CommitEdit(save: true);
            SelectShape(canvas, idx);
            canvas.Focus();
            ShowShapeContextMenu();
        }
        e.Handled = true;
    }

    private void OverlayBody_RightMouseUp(object sender, MouseButtonEventArgs e)
    {
        if (_selectedIdx >= 0) ShowShapeContextMenu();
        e.Handled = true;
    }

    private void ShowShapeContextMenu()
    {
        if (_selectedIdx < 0 || _selectedTreeIdx < 0) return;

        string Res(string key, string fallback) =>
            Application.Current.TryFindResource(key) is string s ? s : fallback;

        var menu = new ContextMenu();

        var propItem = new MenuItem { Header = Res("Ctx_ShapeProperties", "개체 속성 편집...") };
        propItem.Click += (_, _) => ShapePropertiesRequested?.Invoke(_slideIndex, _selectedTreeIdx);
        menu.Items.Add(propItem);

        menu.Items.Add(new Separator());

        // ── Z-order ──────────────────────────────────────────────────
        var frontItem = new MenuItem { Header = Res("Ctx_BringToFront", "맨 앞으로 가져오기") };
        frontItem.Click += (_, _) => ShapeOrderChanged?.Invoke(_slideIndex, _selectedTreeIdx, 2);
        menu.Items.Add(frontItem);

        var fwdItem = new MenuItem { Header = Res("Ctx_BringForward", "앞으로 가져오기") };
        fwdItem.Click += (_, _) => ShapeOrderChanged?.Invoke(_slideIndex, _selectedTreeIdx, 1);
        menu.Items.Add(fwdItem);

        var bkdItem = new MenuItem { Header = Res("Ctx_SendBackward", "뒤로 보내기") };
        bkdItem.Click += (_, _) => ShapeOrderChanged?.Invoke(_slideIndex, _selectedTreeIdx, -1);
        menu.Items.Add(bkdItem);

        var backItem = new MenuItem { Header = Res("Ctx_SendToBack", "맨 뒤로 보내기") };
        backItem.Click += (_, _) => ShapeOrderChanged?.Invoke(_slideIndex, _selectedTreeIdx, -2);
        menu.Items.Add(backItem);

        menu.Items.Add(new Separator());

        var delItem = new MenuItem { Header = Res("Ctx_Delete", "삭제") };
        delItem.Click += (_, _) =>
        {
            if (_selectedTreeIdx >= 0)
                ShapeDeleted?.Invoke(_slideIndex, _selectedTreeIdx);
        };
        menu.Items.Add(delItem);

        menu.Placement = System.Windows.Controls.Primitives.PlacementMode.Mouse;
        menu.IsOpen    = true;
    }

    // ── Hit testing ───────────────────────────────────────────────────

    private static int HitTest(Canvas canvas, Point pos)
    {
        for (int i = canvas.Children.Count - 1; i >= 0; i--)
        {
            if (canvas.Children[i] is FrameworkElement fe && fe.Tag is int)
            {
                double l = Canvas.GetLeft(fe), t = Canvas.GetTop(fe);
                if (new Rect(l, t, fe.Width, fe.Height).Contains(pos))
                    return i;
            }
        }
        return -1;
    }

    // ── Selection & overlay ────────────────────────────────────────────

    private void SelectShape(Canvas canvas, int idx)
    {
        RemoveOverlay(canvas);

        _selectedIdx     = idx;
        _selectedTreeIdx = -1;
        if (idx < 0 || canvas.Children[idx] is not FrameworkElement target) return;
        if (target.Tag is int ti) _selectedTreeIdx = ti;

        double l = Canvas.GetLeft(target);
        double t = Canvas.GetTop(target);
        _selOverlay = BuildOverlay(l, t, target.Width, target.Height);
        Panel.SetZIndex(_selOverlay, 9999);
        canvas.Children.Add(_selOverlay);
    }

    private void RemoveOverlay(Canvas canvas)
    {
        if (_selOverlay is not null)
        {
            canvas.Children.Remove(_selOverlay);
            _selOverlay = null;
        }
    }

    private Canvas BuildOverlay(double l, double t, double w, double h)
    {
        var overlay = new Canvas { Width = w, Height = h };
        Canvas.SetLeft(overlay, l);
        Canvas.SetTop(overlay,  t);

        // Border rect — drag body to move
        var border = new Rectangle
        {
            Width           = w,
            Height          = h,
            Stroke          = new SolidColorBrush(Color.FromRgb(0, 0x78, 0xD4)),
            StrokeThickness = 2,
            StrokeDashArray = new DoubleCollection { 5, 3 },
            Fill            = new SolidColorBrush(Color.FromArgb(12, 0, 0x78, 0xD4)),
            Cursor          = Cursors.SizeAll,
            IsHitTestVisible = true,
        };
        border.MouseLeftButtonDown  += OverlayBody_MouseDown;
        border.MouseRightButtonUp  += OverlayBody_RightMouseUp;
        overlay.Children.Add(border);

        // 8 resize handles
        var pts = HandlePoints(w, h);
        for (int i = 0; i < 8; i++)
        {
            int hi = i;
            var handle = new Rectangle
            {
                Width           = HandleSize,
                Height          = HandleSize,
                Fill            = new SolidColorBrush(Color.FromRgb(0, 0x78, 0xD4)),
                Stroke          = Brushes.White,
                StrokeThickness = 1,
                Cursor          = HandleCursor(i),
                IsHitTestVisible = true,
            };
            Canvas.SetLeft(handle, pts[i].X - HandleSize / 2);
            Canvas.SetTop(handle,  pts[i].Y - HandleSize / 2);
            Panel.SetZIndex(handle, 1);
            handle.MouseLeftButtonDown += (_, ev) => HandleRect_MouseDown(hi, ev);
            overlay.Children.Add(handle);
        }
        return overlay;
    }

    private void OverlayBody_MouseDown(object sender, MouseButtonEventArgs e)
    {
        if (_nativeCanvas is null || _selectedIdx < 0) return;
        var target = _nativeCanvas.Children[_selectedIdx] as FrameworkElement;
        if (target is null) return;

        _dragMode  = DragMode.Move;
        _dragStart = e.GetPosition(_nativeCanvas);
        _dOrigL    = Canvas.GetLeft(target);
        _dOrigT    = Canvas.GetTop(target);
        _dOrigW    = target.Width;
        _dOrigH    = target.Height;
        _nativeCanvas.CaptureMouse();
        e.Handled = true;
    }

    private void HandleRect_MouseDown(int handleIdx, MouseButtonEventArgs e)
    {
        if (_nativeCanvas is null || _selectedIdx < 0) return;
        var target = _nativeCanvas.Children[_selectedIdx] as FrameworkElement;
        if (target is null) return;

        _dragMode     = DragMode.Resize;
        _resizeHandle = handleIdx;
        _dragStart    = e.GetPosition(_nativeCanvas);
        _dOrigL       = Canvas.GetLeft(target);
        _dOrigT       = Canvas.GetTop(target);
        _dOrigW       = target.Width;
        _dOrigH       = target.Height;
        _nativeCanvas.CaptureMouse();
        e.Handled = true;
    }

    private void SelectShapeByTreeIndex(int treeIdx)
    {
        if (_nativeCanvas is null) return;
        for (int i = _nativeCanvas.Children.Count - 1; i >= 0; i--)
        {
            if (_nativeCanvas.Children[i] is FrameworkElement fe &&
                fe.Tag is int t && t == treeIdx)
            {
                SelectShape(_nativeCanvas, i);
                _nativeCanvas.Focus();
                return;
            }
        }
    }

    // ── Resize calculation ─────────────────────────────────────────────

    // Handle indices:  0=TL  1=TC  2=TR  3=MR  4=BR  5=BC  6=BL  7=ML
    private (double l, double t, double w, double h) ComputeResizeBounds(double dx, double dy)
    {
        double l = _dOrigL, t = _dOrigT, w = _dOrigW, h = _dOrigH;

        switch (_resizeHandle)
        {
            case 0: l += dx; t += dy; w -= dx; h -= dy; break; // TL
            case 1:          t += dy;           h -= dy; break; // TC
            case 2:          t += dy; w += dx;  h -= dy; break; // TR
            case 3:                   w += dx;            break; // MR
            case 4:                   w += dx;  h += dy; break; // BR
            case 5:                             h += dy; break; // BC
            case 6: l += dx;          w -= dx;  h += dy; break; // BL
            case 7: l += dx;          w -= dx;            break; // ML
        }

        if (w < MinShapeSize)
        {
            if (_resizeHandle is 0 or 6 or 7) l = _dOrigL + _dOrigW - MinShapeSize;
            w = MinShapeSize;
        }
        if (h < MinShapeSize)
        {
            if (_resizeHandle is 0 or 1 or 2) t = _dOrigT + _dOrigH - MinShapeSize;
            h = MinShapeSize;
        }
        return (l, t, w, h);
    }

    private void UpdateOverlayBounds(double l, double t, double w, double h)
    {
        if (_selOverlay is null) return;
        Canvas.SetLeft(_selOverlay, l);
        Canvas.SetTop(_selOverlay,  t);
        _selOverlay.Width  = w;
        _selOverlay.Height = h;

        if (_selOverlay.Children[0] is Rectangle border)
        {
            border.Width  = w;
            border.Height = h;
        }
        var pts = HandlePoints(w, h);
        for (int i = 1; i <= 8 && i < _selOverlay.Children.Count; i++)
        {
            if (_selOverlay.Children[i] is Rectangle hr)
            {
                Canvas.SetLeft(hr, pts[i - 1].X - HandleSize / 2);
                Canvas.SetTop(hr,  pts[i - 1].Y - HandleSize / 2);
            }
        }
    }

    // ── Handle geometry helpers ────────────────────────────────────────

    private static Point[] HandlePoints(double w, double h) =>
    [
        new(0,   0),     // 0 TL
        new(w/2, 0),     // 1 TC
        new(w,   0),     // 2 TR
        new(w,   h/2),   // 3 MR
        new(w,   h),     // 4 BR
        new(w/2, h),     // 5 BC
        new(0,   h),     // 6 BL
        new(0,   h/2),   // 7 ML
    ];

    private static Cursor HandleCursor(int idx) => idx switch
    {
        0 or 4 => Cursors.SizeNWSE,
        1 or 5 => Cursors.SizeNS,
        2 or 6 => Cursors.SizeNESW,
        3 or 7 => Cursors.SizeWE,
        _      => Cursors.Arrow,
    };

    private long PxToEmu(double px) => (long)(px * EmuPerPx);

    // ── Inline text editing ────────────────────────────────────────────

    private void StartEdit(Canvas canvas, int canvasIdx)
    {
        if (canvas.Children[canvasIdx] is not FrameworkElement target) return;
        if (target.Tag is not int treeIdx) return;

        var elements = _slidePart?.Slide.CommonSlideData?.ShapeTree?
            .Elements<OpenXmlCompositeElement>().ToList();
        if (elements is null || treeIdx >= elements.Count) return;
        if (elements[treeIdx] is not DocumentFormat.OpenXml.Presentation.Shape shape || shape.TextBody is null) return;

        CommitEdit(save: false);
        SelectShape(canvas, canvasIdx);

        var doc = PptxConverter.ToFlowDocument(shape.TextBody);

        _editor = new RichTextBox(doc)
        {
            Width            = target.Width,
            Height           = target.Height,
            Background       = new SolidColorBrush(Color.FromArgb(230, 255, 255, 255)),
            BorderBrush      = new SolidColorBrush(Color.FromRgb(0, 0x78, 0xD4)),
            BorderThickness  = new Thickness(2),
            AcceptsReturn    = true,
            VerticalScrollBarVisibility   = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            FontFamily       = new FontFamily("맑은 고딕"),
        };
        Panel.SetZIndex(_editor, 10000);
        Canvas.SetLeft(_editor, Canvas.GetLeft(target));
        Canvas.SetTop(_editor,  Canvas.GetTop(target));

        _editingTreeIdx = treeIdx;
        canvas.Children.Add(_editor);
        _editor.Focus();
        _editor.SelectAll();

        EditingStarted?.Invoke(_editor);

        _editor.LostFocus += (_, _) => { if (!SuppressLostFocusCommit) CommitEdit(save: true); };
        _editor.KeyDown   += Editor_KeyDown;
    }

    private void Editor_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
        {
            CommitEdit(save: false);
            e.Handled = true;
        }
    }

    public void CommitEdit(bool save)
    {
        if (_editor is null || _nativeCanvas is null) return;

        if (save && _editingTreeIdx >= 0)
        {
            var paragraphs = PptxConverter.FromFlowDocument(_editor.Document);
            TextCommitted?.Invoke(_slideIndex, _editingTreeIdx, paragraphs);
        }

        _nativeCanvas.Children.Remove(_editor);
        _editor         = null;
        _editingTreeIdx = -1;
    }

    // ── Public helpers for formatting toolbar ──────────────────────────

    public void ApplySelectionProperty(DependencyProperty prop, object value)
        => _editor?.Selection.ApplyPropertyValue(prop, value);

    public object? GetSelectionProperty(DependencyProperty prop)
        => _editor?.Selection.GetPropertyValue(prop);
}
