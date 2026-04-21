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

    // ── State ──────────────────────────────────────────────────────────

    private PresentationModel? _model;
    private SlidePart?         _slidePart;
    private int                _slideIndex;
    private Canvas?            _nativeCanvas;      // the live SlideRenderer canvas
    private double             _zoom = 1.0;

    private int       _selectedIdx  = -1;          // canvas child index
    private Rectangle? _selRect;                   // selection highlight
    private RichTextBox? _editor;                  // active inline editor
    private int       _editingTreeIdx = -1;        // shape-tree index being edited

    // ── Native canvas dimensions (EMU→WPF native pixels) ──────────────
    private double NativeW => _model is not null ? _model.SlideWidth  / 914400.0 * 96.0 : 960;
    private double NativeH => _model is not null ? _model.SlideHeight / 914400.0 * 96.0 : 540;

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

    public void Invalidate()
    {
        CommitEdit(save: false);
        Rebuild();
    }

    public void ZoomIn()      { _zoom = Math.Min(3.0,  Math.Round(_zoom + 0.1, 1)); ApplyZoom(); }
    public void ZoomOut()     { _zoom = Math.Max(0.25, Math.Round(_zoom - 0.1, 1)); ApplyZoom(); }
    public void FitToWindow() { _zoom = 1.0; UpdateViewboxSize(); }
    public void SetZoom(double z) { _zoom = Math.Clamp(z, 0.25, 3.0); ApplyZoom(); }
    public double ZoomFactor  => _zoom;

    // ── Rebuild (re-render slide) ──────────────────────────────────────

    private void Rebuild()
    {
        _selectedIdx    = -1;
        _selRect        = null;
        _editingTreeIdx = -1;
        NoDocMsg.Visibility = Visibility.Collapsed;

        if (_model is null || !_model.IsOpen || _slidePart is null)
        {
            SlideViewbox.Child  = new Canvas { Width = 960, Height = 540 };
            _nativeCanvas       = null;
            NoDocMsg.Visibility = Visibility.Visible;
            return;
        }

        var canvas = SlideRenderer.BuildCanvas(_slidePart, _model.SlideWidth, _model.SlideHeight);
        canvas.MouseLeftButtonDown += Canvas_MouseDown;
        canvas.KeyDown             += Canvas_KeyDown;
        canvas.Focusable            = true;

        _nativeCanvas      = canvas;
        SlideViewbox.Child = canvas;
        UpdateViewboxSize();
    }

    // ── Zoom & sizing ──────────────────────────────────────────────────

    private void ApplyZoom()    => UpdateViewboxSize();

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
            // Future: delete selected shape
        }
    }

    // ── Hit testing ───────────────────────────────────────────────────

    /// <summary>Returns canvas child index at <paramref name="pos"/> (top of Z-order wins).</summary>
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

    // ── Selection ─────────────────────────────────────────────────────

    private void SelectShape(Canvas canvas, int idx)
    {
        // Remove old selection overlay
        if (_selRect is not null)
        {
            canvas.Children.Remove(_selRect);
            _selRect = null;
        }

        _selectedIdx = idx;
        if (idx < 0 || canvas.Children[idx] is not FrameworkElement target) return;

        _selRect = new Rectangle
        {
            Width           = target.Width,
            Height          = target.Height,
            Stroke          = new SolidColorBrush(Color.FromRgb(0, 0x78, 0xD4)),
            StrokeThickness = 2,
            StrokeDashArray = new DoubleCollection { 5, 3 },
            Fill            = new SolidColorBrush(Color.FromArgb(20, 0, 0x78, 0xD4)),
            IsHitTestVisible = false,
        };
        Canvas.SetLeft(_selRect, Canvas.GetLeft(target));
        Canvas.SetTop(_selRect,  Canvas.GetTop(target));
        Panel.SetZIndex(_selRect, 9999);
        canvas.Children.Add(_selRect);
    }

    // ── Inline text editing ────────────────────────────────────────────

    private void StartEdit(Canvas canvas, int canvasIdx)
    {
        if (canvas.Children[canvasIdx] is not FrameworkElement target) return;
        if (target.Tag is not int treeIdx) return;

        // Only shapes with text bodies are editable
        var elements = _slidePart?.Slide.CommonSlideData?.ShapeTree?
            .Elements<OpenXmlCompositeElement>().ToList();
        if (elements is null || treeIdx >= elements.Count) return;
        if (elements[treeIdx] is not DocumentFormat.OpenXml.Presentation.Shape shape || shape.TextBody is null) return;

        CommitEdit(save: false);
        SelectShape(canvas, canvasIdx);

        // Build FlowDocument from PPTX text body
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

        // Notify MainWindow to enable the formatting toolbar
        EditingStarted?.Invoke(_editor);

        _editor.LostFocus += (_, _) => CommitEdit(save: true);
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

    /// <summary>Apply a property value to the current RichTextBox selection.</summary>
    public void ApplySelectionProperty(DependencyProperty prop, object value)
        => _editor?.Selection.ApplyPropertyValue(prop, value);

    public object? GetSelectionProperty(DependencyProperty prop)
        => _editor?.Selection.GetPropertyValue(prop);

    public bool IsEditing => _editor is not null;
}
