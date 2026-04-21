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
using WpfTE    = System.Windows.Documents.TextElement;
using RgbColor = PPEditer.Models.RgbColor;

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
    /// <summary>User finished drawing a shape. (slideIdx, tool, pointsInCanvasPx)</summary>
    public event Action<int, DrawTool, System.Windows.Point[]>? ShapeDrawn;
    /// <summary>User requested grouping of multiple shapes. (slideIdx, treeIndices[])</summary>
    public event Action<int, int[]>? ShapesGroupRequested;
    /// <summary>User requested ungrouping of a GroupShape. (slideIdx, treeIdx)</summary>
    public event Action<int, int>? ShapeUngroupRequested;
    /// <summary>User requested character properties while editing. (slideIdx, treeIdx)</summary>
    public event Action<int, int>? CharPropertiesRequested;
    /// <summary>User requested paragraph properties while editing. (slideIdx, treeIdx)</summary>
    public event Action<int, int>? ParaPropertiesRequested;
    /// <summary>User finished drawing a text box. (slideIdx, leftEmu, topEmu, widthEmu, heightEmu)</summary>
    public event Action<int, long, long, long, long>? TextBoxDrawn;

    // ── Public properties ──────────────────────────────────────────────

    /// <summary>Set true before opening a dialog so LostFocus does not commit and close the editor.</summary>
    public bool SuppressLostFocusCommit { get; set; }

    /// <summary>Currently active drawing tool. Setting to non-Select changes cursor and cancels any in-progress draw.</summary>
    public DrawTool ActiveTool
    {
        get => _activeTool;
        set
        {
            if (_activeTool == value) return;
            CancelDrawing();
            _activeTool = value;
            if (_nativeCanvas is not null)
                _nativeCanvas.Cursor = value == DrawTool.Select ? Cursors.Arrow : Cursors.Cross;
        }
    }

    /// <summary>Exposes the canvas EMU-per-pixel ratio so callers can convert point coordinates to EMU.</summary>
    public double EmuPerPixel => EmuPerPx;

    /// <summary>Primary selected shape-tree index (-1 if none).</summary>
    public int SelectedTreeIndex => _selectedTreeIdx;

    /// <summary>All selected shape-tree indices (primary + multi-select).</summary>
    public int[] SelectedTreeIndices
    {
        get
        {
            if (_selectedTreeIdx < 0) return [];
            if (_multiTreeIdxs.Count == 0) return [_selectedTreeIdx];
            var all = new List<int> { _selectedTreeIdx };
            all.AddRange(_multiTreeIdxs);
            return [.. all.Distinct()];
        }
    }

    // ── State ──────────────────────────────────────────────────────────

    private PresentationModel? _model;
    private SlidePart?         _slidePart;
    private int                _slideIndex;
    private Canvas?            _nativeCanvas;
    private double             _zoom = 1.0;

    // ── Drawing state ──────────────────────────────────────────────────
    private DrawTool    _activeTool     = DrawTool.Select;
    private bool        _isDrawing      = false;
    private List<Point> _drawPoints     = [];
    private Point       _drawCurrentPos;
    private UIElement?  _drawPreview;

    private int      _selectedIdx     = -1;   // canvas child index
    private int      _selectedTreeIdx = -1;   // shape-tree index
    private Canvas?  _selOverlay;             // selection overlay (border + 8 handles)
    private RichTextBox? _editor;
    private int      _editingTreeIdx  = -1;

    // ── Multi-select state ─────────────────────────────────────────────
    private readonly HashSet<int>            _multiTreeIdxs  = [];
    private readonly Dictionary<int, Canvas> _multiOverlays  = [];

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

    /// <summary>Start editing the shape with the specified tree index (if it has a TextBody).</summary>
    public void StartEditByTreeIndex(int treeIdx)
    {
        if (_nativeCanvas is null) return;
        for (int i = 0; i < _nativeCanvas.Children.Count; i++)
        {
            if (_nativeCanvas.Children[i] is FrameworkElement fe && fe.Tag is int t && t == treeIdx)
            {
                StartEdit(_nativeCanvas, i);
                return;
            }
        }
    }

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
        _multiTreeIdxs.Clear();
        _multiOverlays.Clear();
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
        if (_activeTool != DrawTool.Select)
            canvas.Cursor = Cursors.Cross;

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

        if (_activeTool != DrawTool.Select)
        {
            HandleDrawMouseDown(canvas, clickPos, e.ClickCount);
            canvas.Focus();
            e.Handled = true;
            return;
        }

        if (e.ClickCount == 1)
        {
            CommitEdit(save: true);
            int idx = HitTest(canvas, clickPos);
            if ((Keyboard.Modifiers & ModifierKeys.Control) != 0 && idx >= 0 && _selectedIdx >= 0)
            {
                int treeIdx = canvas.Children[idx] is FrameworkElement fe2 &&
                              fe2.Tag is int ti ? ti : -1;
                if (treeIdx >= 0 && treeIdx != _selectedTreeIdx)
                {
                    if (_multiTreeIdxs.Contains(treeIdx))
                        RemoveMultiOverlay(canvas, treeIdx);
                    else
                        AddMultiSelect(canvas, idx);
                }
            }
            else
            {
                RemoveAllMultiOverlays(canvas);
                SelectShape(canvas, idx);
            }
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
        if (_nativeCanvas is null) return;
        var pos = e.GetPosition(_nativeCanvas);

        if (_activeTool != DrawTool.Select && _isDrawing)
        {
            _drawCurrentPos = pos;
            if (IsDragTool(_activeTool)) UpdateDragPreview(_nativeCanvas);
            else                         UpdateClickPreview(_nativeCanvas);
            return;
        }

        if (_dragMode == DragMode.None || _selOverlay is null) return;

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
        if (_activeTool != DrawTool.Select && _isDrawing && IsDragTool(_activeTool))
        {
            _nativeCanvas?.ReleaseMouseCapture();
            var endPos = e.GetPosition((Canvas)sender);
            if (_drawPoints.Count > 0) _drawPoints.Add(endPos);
            FinalizeDrawing((Canvas)sender);
            e.Handled = true;
            return;
        }

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
            if (_isDrawing)
            {
                CancelDrawing();
                e.Handled = true;
                return;
            }
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
        if (_isDrawing)
        {
            CancelDrawing();
            e.Handled = true;
            return;
        }
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

        // ── Group / Ungroup ──────────────────────────────────────────
        bool hasMulti = _multiTreeIdxs.Count > 0;
        bool isGroup  = _model?.IsGroupShape(_slideIndex, _selectedTreeIdx) ?? false;
        if (hasMulti)
        {
            var groupItem = new MenuItem { Header = Res("Ctx_Group", "그룹으로 묶기") };
            groupItem.Click += (_, _) => ShapesGroupRequested?.Invoke(_slideIndex, SelectedTreeIndices);
            menu.Items.Add(groupItem);
            menu.Items.Add(new Separator());
        }
        else if (isGroup)
        {
            var ungroupItem = new MenuItem { Header = Res("Ctx_Ungroup", "그룹 해제") };
            ungroupItem.Click += (_, _) => ShapeUngroupRequested?.Invoke(_slideIndex, _selectedTreeIdx);
            menu.Items.Add(ungroupItem);
            menu.Items.Add(new Separator());
        }

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

    // ── Drawing mode ────────────────────────────────────────────────────

    private static bool IsDragTool(DrawTool t) => t is
        DrawTool.Line or DrawTool.Square or DrawTool.Rect or
        DrawTool.Ellipse or DrawTool.Circle or DrawTool.EqTriangle or
        DrawTool.IsoTriangle or DrawTool.RightTriangle or
        DrawTool.Trapezoid or DrawTool.Parallelogram or
        DrawTool.Arc or DrawTool.Arrow or DrawTool.TextBox;

    private static bool IsClickTool(DrawTool t) => t is
        DrawTool.PolyLine or DrawTool.SplineLine or
        DrawTool.Polygon or DrawTool.SplinePolygon or
        DrawTool.ScaleneTriangle;

    private void HandleDrawMouseDown(Canvas canvas, Point pos, int clickCount)
    {
        if (IsDragTool(_activeTool))
        {
            if (clickCount != 1) return;
            _isDrawing = true;
            _drawPoints = [pos];
            _drawCurrentPos = pos;
            canvas.CaptureMouse();
            RemoveDrawPreview(canvas);
            CreateDragPreview(canvas);
            return;
        }

        if (!IsClickTool(_activeTool)) return;

        if (clickCount == 2)
        {
            if (_isDrawing && _drawPoints.Count >= 2)
            {
                // Remove duplicate point added by the first click of this double-click
                var last = _drawPoints[^1];
                var prev = _drawPoints[^2];
                if (Math.Abs(last.X - prev.X) < 5 && Math.Abs(last.Y - prev.Y) < 5)
                    _drawPoints.RemoveAt(_drawPoints.Count - 1);
                FinalizeDrawing(canvas);
            }
            return;
        }

        // Single click
        if (!_isDrawing)
        {
            _isDrawing = true;
            _drawPoints = [pos];
            _drawCurrentPos = pos;
            RemoveDrawPreview(canvas);
            CreateClickPreview(canvas);
        }
        else
        {
            _drawPoints.Add(pos);
            UpdateClickPreview(canvas);
            if (_activeTool == DrawTool.ScaleneTriangle && _drawPoints.Count == 3)
                FinalizeDrawing(canvas);
        }
    }

    private void FinalizeDrawing(Canvas canvas)
    {
        var pts  = _drawPoints.ToArray();
        var tool = _activeTool;
        RemoveDrawPreview(canvas);
        canvas.ReleaseMouseCapture();
        _isDrawing = false;
        _drawPoints.Clear();
        _activeTool = DrawTool.Select;
        if (_nativeCanvas is not null)
            _nativeCanvas.Cursor = Cursors.Arrow;
        if (pts.Length < 2) return;

        if (tool == DrawTool.TextBox)
        {
            // Fire TextBoxDrawn with EMU bounds
            double l = Math.Min(pts[0].X, pts[1].X);
            double t = Math.Min(pts[0].Y, pts[1].Y);
            double w = Math.Abs(pts[1].X - pts[0].X);
            double h = Math.Abs(pts[1].Y - pts[0].Y);
            if (w < 10) w = 200;
            if (h < 10) h = 40;
            TextBoxDrawn?.Invoke(_slideIndex,
                PxToEmu(l), PxToEmu(t), PxToEmu(w), PxToEmu(h));
        }
        else
        {
            ShapeDrawn?.Invoke(_slideIndex, tool, pts);
        }
    }

    private void CancelDrawing()
    {
        if (_nativeCanvas is not null)
            RemoveDrawPreview(_nativeCanvas);
        _nativeCanvas?.ReleaseMouseCapture();
        _isDrawing = false;
        _drawPoints.Clear();
        _activeTool = DrawTool.Select;
        if (_nativeCanvas is not null)
            _nativeCanvas.Cursor = Cursors.Arrow;
    }

    private void RemoveDrawPreview(Canvas canvas)
    {
        if (_drawPreview is null) return;
        canvas.Children.Remove(_drawPreview);
        _drawPreview = null;
    }

    private void CreateDragPreview(Canvas canvas)
    {
        if (_drawPoints.Count == 0) return;
        var start = _drawPoints[0];
        UIElement preview;
        if (_activeTool == DrawTool.Line)
        {
            preview = new System.Windows.Shapes.Line
            {
                X1 = start.X, Y1 = start.Y,
                X2 = start.X, Y2 = start.Y,
                Stroke           = new SolidColorBrush(Color.FromRgb(0x2E, 0x74, 0xB5)),
                StrokeThickness  = 1.5,
                StrokeDashArray  = new DoubleCollection { 4, 2 },
                IsHitTestVisible = false,
            };
        }
        else
        {
            var rect = new System.Windows.Shapes.Rectangle
            {
                Stroke           = new SolidColorBrush(Color.FromRgb(0x2E, 0x74, 0xB5)),
                StrokeThickness  = 1.5,
                StrokeDashArray  = new DoubleCollection { 4, 2 },
                Fill             = new SolidColorBrush(Color.FromArgb(30, 0xBD, 0xD7, 0xEE)),
                Width            = 0,
                Height           = 0,
                IsHitTestVisible = false,
            };
            Canvas.SetLeft(rect, start.X);
            Canvas.SetTop(rect,  start.Y);
            preview = rect;
        }
        Panel.SetZIndex(preview, 9998);
        canvas.Children.Add(preview);
        _drawPreview = preview;
    }

    private void UpdateDragPreview(Canvas canvas)
    {
        if (_drawPreview is null || _drawPoints.Count == 0) return;
        var start = _drawPoints[0];
        var end   = _drawCurrentPos;
        if (_drawPreview is System.Windows.Shapes.Line line)
        {
            line.X2 = end.X;
            line.Y2 = end.Y;
        }
        else if (_drawPreview is System.Windows.Shapes.Rectangle rect)
        {
            double x = Math.Min(start.X, end.X);
            double y = Math.Min(start.Y, end.Y);
            Canvas.SetLeft(rect, x);
            Canvas.SetTop(rect,  y);
            rect.Width  = Math.Max(1, Math.Abs(end.X - start.X));
            rect.Height = Math.Max(1, Math.Abs(end.Y - start.Y));
        }
    }

    private void CreateClickPreview(Canvas canvas)
    {
        bool closed = _activeTool is DrawTool.Polygon or DrawTool.SplinePolygon or DrawTool.ScaleneTriangle;
        UIElement preview;
        if (closed)
        {
            preview = new System.Windows.Shapes.Polygon
            {
                Stroke           = new SolidColorBrush(Color.FromRgb(0x2E, 0x74, 0xB5)),
                StrokeThickness  = 1.5,
                StrokeDashArray  = new DoubleCollection { 4, 2 },
                Fill             = new SolidColorBrush(Color.FromArgb(30, 0xBD, 0xD7, 0xEE)),
                IsHitTestVisible = false,
            };
        }
        else
        {
            preview = new System.Windows.Shapes.Polyline
            {
                Stroke           = new SolidColorBrush(Color.FromRgb(0x2E, 0x74, 0xB5)),
                StrokeThickness  = 1.5,
                StrokeDashArray  = new DoubleCollection { 4, 2 },
                IsHitTestVisible = false,
            };
        }
        Panel.SetZIndex(preview, 9998);
        canvas.Children.Add(preview);
        _drawPreview = preview;
        UpdateClickPreview(canvas);
    }

    private void UpdateClickPreview(Canvas canvas)
    {
        if (_drawPreview is null) return;
        var pts = new PointCollection(_drawPoints) { _drawCurrentPos };
        if (_drawPreview is System.Windows.Shapes.Polygon poly)
            poly.Points = pts;
        else if (_drawPreview is System.Windows.Shapes.Polyline pline)
            pline.Points = pts;
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

        if (e.ClickCount == 2)
        {
            // Double-click on selected shape → enter edit mode
            StartEdit(_nativeCanvas, _selectedIdx);
            e.Handled = true;
            return;
        }

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

    // ── Multi-select helpers ───────────────────────────────────────────

    private void AddMultiSelect(Canvas canvas, int canvasIdx)
    {
        if (canvas.Children[canvasIdx] is not FrameworkElement fe) return;
        if (fe.Tag is not int treeIdx) return;
        if (_multiTreeIdxs.Contains(treeIdx)) return;

        _multiTreeIdxs.Add(treeIdx);
        double l = Canvas.GetLeft(fe);
        double t = Canvas.GetTop(fe);
        var overlay = BuildSimpleOverlay(l, t, fe.Width, fe.Height);
        Panel.SetZIndex(overlay, 9999);
        canvas.Children.Add(overlay);
        _multiOverlays[treeIdx] = overlay;
    }

    private void RemoveMultiOverlay(Canvas canvas, int treeIdx)
    {
        if (_multiOverlays.TryGetValue(treeIdx, out var overlay))
        {
            canvas.Children.Remove(overlay);
            _multiOverlays.Remove(treeIdx);
        }
        _multiTreeIdxs.Remove(treeIdx);
    }

    private void RemoveAllMultiOverlays(Canvas canvas)
    {
        foreach (var overlay in _multiOverlays.Values)
            canvas.Children.Remove(overlay);
        _multiOverlays.Clear();
        _multiTreeIdxs.Clear();
    }

    private static Canvas BuildSimpleOverlay(double l, double t, double w, double h)
    {
        var overlay = new Canvas { Width = w, Height = h };
        Canvas.SetLeft(overlay, l);
        Canvas.SetTop(overlay,  t);
        var border = new Rectangle
        {
            Width           = w,
            Height          = h,
            Stroke          = new SolidColorBrush(Color.FromRgb(0, 0x78, 0xD4)),
            StrokeThickness = 1.5,
            StrokeDashArray = new DoubleCollection { 4, 2 },
            Fill            = new SolidColorBrush(Color.FromArgb(8, 0, 0x78, 0xD4)),
            IsHitTestVisible = false,
        };
        overlay.Children.Add(border);
        return overlay;
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

        // Add context menu for character/paragraph properties
        var charCtx  = new ContextMenu();
        int capturedSlideIdx = _slideIndex;
        int capturedTreeIdx  = treeIdx;
        var cpItem = new MenuItem
        {
            Header = Application.Current.TryFindResource("Ctx_CharProperties") as string ?? "글자 속성...",
        };
        cpItem.Click += (_, _) => CharPropertiesRequested?.Invoke(capturedSlideIdx, capturedTreeIdx);
        charCtx.Items.Add(cpItem);
        var ppItem = new MenuItem
        {
            Header = Application.Current.TryFindResource("Ctx_ParaProperties") as string ?? "문단 속성...",
        };
        ppItem.Click += (_, _) => ParaPropertiesRequested?.Invoke(capturedSlideIdx, capturedTreeIdx);
        charCtx.Items.Add(ppItem);
        _editor.ContextMenu = charCtx;

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

    // ── Character style (GetEditorCharStyle / ApplyEditorCharStyle) ────

    public CharStyle GetEditorCharStyle()
    {
        if (_editor is null) return new CharStyle();
        var sel = _editor.Selection;

        string? ff = (sel.GetPropertyValue(WpfTE.FontFamilyProperty) as FontFamily)?.Source;
        double? sizePt = sel.GetPropertyValue(WpfTE.FontSizeProperty) is double d && d > 0
            ? d * 72.0 / 96.0 : (double?)null;
        bool? bold   = sel.GetPropertyValue(WpfTE.FontWeightProperty) is FontWeight w
            ? w == FontWeights.Bold : (bool?)null;
        bool? italic = sel.GetPropertyValue(WpfTE.FontStyleProperty) is FontStyle fstyle
            ? fstyle == FontStyles.Italic : (bool?)null;

        var decos = sel.GetPropertyValue(Inline.TextDecorationsProperty) as TextDecorationCollection;
        bool? ul    = decos?.Any(d => d.Location == TextDecorationLocation.Underline);
        bool? strike = decos?.Any(d => d.Location == TextDecorationLocation.Strikethrough);
        bool? overl  = decos?.Any(d => d.Location == TextDecorationLocation.OverLine);

        var fv     = sel.GetPropertyValue(Typography.VariantsProperty) is FontVariants variants
            ? variants : FontVariants.Normal;
        var script = fv == FontVariants.Superscript ? ScriptKind.Superscript :
                     fv == FontVariants.Subscript   ? ScriptKind.Subscript   :
                                                      ScriptKind.None;

        RgbColor? fore = sel.GetPropertyValue(WpfTE.ForegroundProperty) is SolidColorBrush fg
            ? new RgbColor(fg.Color.R, fg.Color.G, fg.Color.B) : (RgbColor?)null;
        RgbColor? back = sel.GetPropertyValue(WpfTE.BackgroundProperty) is SolidColorBrush bg
                         && bg.Color.A > 0
            ? new RgbColor(bg.Color.R, bg.Color.G, bg.Color.B) : (RgbColor?)null;

        // Extra props from first Run in selection
        int spacing = 0;
        RgbColor? ulColor   = null;
        bool hasOutline     = false;
        var tp = sel.Start.GetAdjacentElement(LogicalDirection.Forward) as Run;
        if (tp?.Tag is PPEditer.Services.PptxConverter.CharExtraProps ex)
        {
            spacing    = ex.SpacingPt100;
            ulColor    = ex.UnderlineColor;
            hasOutline = ex.HasOutline;
        }

        return new CharStyle
        {
            FontFamily     = ff,
            FontSizePt     = sizePt,
            Bold           = bold,
            Italic         = italic,
            HasUnderline   = ul,
            HasStrike      = strike,
            HasOverline    = overl,
            HasOutline     = hasOutline,
            Script         = script,
            SpacingPt100   = spacing,
            ForeColor      = fore,
            HighlightColor = back,
            UnderlineColor = ulColor,
        };
    }

    // ── Paragraph style (GetEditorParaStyle / ApplyEditorParaStyle) ────

    public ParagraphStyle GetEditorParaStyle()
    {
        if (_editor is null) return new ParagraphStyle();

        // Read VertAnchor from FlowDocument.Tag (set during ToFlowDocument)
        var vertAnchor = _editor.Document.Tag is VertAnchor va ? va : VertAnchor.Top;

        // Find the paragraph at the caret
        var caretPara = _editor.CaretPosition.Paragraph
                        ?? _editor.Document.Blocks.OfType<Paragraph>().FirstOrDefault();
        if (caretPara is null) return new ParagraphStyle { VertAnchor = vertAnchor };

        var horzAlign = caretPara.TextAlignment switch
        {
            TextAlignment.Center  => HorzAlign.Center,
            TextAlignment.Right   => HorzAlign.Right,
            TextAlignment.Justify => HorzAlign.Justify,
            _                     => HorzAlign.Left,
        };

        double marginLeftCm  = caretPara.Margin.Left  / 96.0 * 2.54;
        double textIndentCm  = caretPara.TextIndent    / 96.0 * 2.54;
        double spaceBeforePt = caretPara.Margin.Top    * 72.0 / 96.0;
        double spaceAfterPt  = caretPara.Margin.Bottom * 72.0 / 96.0;

        double lineSpacePct = 100.0;
        if (caretPara.Tag is double pct && pct > 0)
            lineSpacePct = pct;

        return new ParagraphStyle
        {
            HorzAlign     = horzAlign,
            VertAnchor    = vertAnchor,
            MarginLeftCm  = Math.Max(0, marginLeftCm),
            TextIndentCm  = textIndentCm,  // negative = hanging
            LineSpacePct  = lineSpacePct,
            SpaceBeforePt = Math.Max(0, spaceBeforePt),
            SpaceAfterPt  = Math.Max(0, spaceAfterPt),
        };
    }

    public void ApplyEditorParaStyle(ParagraphStyle style)
    {
        if (_editor is null) return;

        // Update VertAnchor on document tag (written to PPTX on commit via UpdateBodyVertAnchor)
        _editor.Document.Tag = style.VertAnchor;

        // Apply to selected paragraphs (or all if no selection spans paragraphs)
        var sel = _editor.Selection;
        var startPara = sel.Start.Paragraph ?? _editor.Document.Blocks.OfType<Paragraph>().FirstOrDefault();
        var endPara   = sel.End.Paragraph   ?? startPara;

        foreach (var block in _editor.Document.Blocks.OfType<Paragraph>())
        {
            if (block.ContentEnd.CompareTo(sel.Start) < 0) continue;
            if (block.ContentStart.CompareTo(sel.End) > 0) break;

            block.TextAlignment = style.HorzAlign switch
            {
                HorzAlign.Center  => TextAlignment.Center,
                HorzAlign.Right   => TextAlignment.Right,
                HorzAlign.Justify => TextAlignment.Justify,
                _                 => TextAlignment.Left,
            };

            double leftPx  = style.MarginLeftCm  * 96.0 / 2.54;
            double indentPx = style.TextIndentCm  * 96.0 / 2.54;  // negative = hanging
            double topPx   = style.SpaceBeforePt * 96.0 / 72.0;
            double botPx   = style.SpaceAfterPt  * 96.0 / 72.0;

            block.Margin      = new Thickness(leftPx, topPx, 0, botPx);
            block.TextIndent  = indentPx;
            block.Tag         = style.LineSpacePct;

            if (style.LineSpacePct > 0 && !double.IsNaN(style.LineSpacePct))
            {
                double fontPx = block.Inlines.OfType<Run>()
                    .Select(r => r.FontSize)
                    .Where(s => s > 0)
                    .DefaultIfEmpty(14.0 * 96.0 / 72.0)
                    .Max();
                block.LineHeight = fontPx * style.LineSpacePct / 100.0;
            }
            else
                block.LineHeight = double.NaN;
        }

        _editor.Focus();
    }

    public void ApplyEditorCharStyle(CharStyle style)
    {
        if (_editor is null) return;
        var sel = _editor.Selection;
        if (sel.IsEmpty) _editor.SelectAll();

        if (style.FontFamily is not null)
            sel.ApplyPropertyValue(WpfTE.FontFamilyProperty, new FontFamily(style.FontFamily));
        if (style.FontSizePt.HasValue)
            sel.ApplyPropertyValue(WpfTE.FontSizeProperty, style.FontSizePt.Value * 96.0 / 72.0);
        if (style.Bold.HasValue)
            sel.ApplyPropertyValue(WpfTE.FontWeightProperty,
                style.Bold.Value ? FontWeights.Bold : FontWeights.Normal);
        if (style.Italic.HasValue)
            sel.ApplyPropertyValue(WpfTE.FontStyleProperty,
                style.Italic.Value ? FontStyles.Italic : FontStyles.Normal);

        var decos = new TextDecorationCollection();
        if (style.HasUnderline == true) decos.Add(TextDecorations.Underline[0]);
        if (style.HasStrike    == true) decos.Add(TextDecorations.Strikethrough[0]);
        if (style.HasOverline  == true) decos.Add(TextDecorations.OverLine[0]);
        sel.ApplyPropertyValue(Inline.TextDecorationsProperty, decos);

        var fv = style.Script switch
        {
            ScriptKind.Superscript => FontVariants.Superscript,
            ScriptKind.Subscript   => FontVariants.Subscript,
            _                      => FontVariants.Normal,
        };
        sel.ApplyPropertyValue(Typography.VariantsProperty, fv);

        if (style.ForeColor.HasValue)
            sel.ApplyPropertyValue(WpfTE.ForegroundProperty,
                new SolidColorBrush(Color.FromRgb(style.ForeColor.Value.R,
                                                  style.ForeColor.Value.G,
                                                  style.ForeColor.Value.B)));
        if (style.HighlightColor.HasValue)
            sel.ApplyPropertyValue(WpfTE.BackgroundProperty,
                new SolidColorBrush(Color.FromRgb(style.HighlightColor.Value.R,
                                                  style.HighlightColor.Value.G,
                                                  style.HighlightColor.Value.B)));

        // Update extra props (spacing, ulColor, outline) on each Run in selection
        foreach (var block in _editor.Document.Blocks.OfType<Paragraph>())
            foreach (var inline in block.Inlines.OfType<Run>())
            {
                if (inline.ContentEnd.CompareTo(sel.Start) < 0 ||
                    inline.ContentStart.CompareTo(sel.End) > 0) continue;
                var ex = inline.Tag as PPEditer.Services.PptxConverter.CharExtraProps
                         ?? new PPEditer.Services.PptxConverter.CharExtraProps();
                ex.SpacingPt100 = style.SpacingPt100;
                if (style.UnderlineColor.HasValue) ex.UnderlineColor = style.UnderlineColor;
                if (style.HasOutline.HasValue)     ex.HasOutline     = style.HasOutline.Value;
                if (style.HasOverline.HasValue)    ex.HasOverline    = style.HasOverline.Value;
                inline.Tag = ex;
            }
        _editor.Focus();
    }
}
