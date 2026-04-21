using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using DocumentFormat.OpenXml.Packaging;
using PPEditer.Models;
using PPEditer.Rendering;

namespace PPEditer.Controls;

public partial class SlideEditorCanvas : UserControl
{
    public event Action<int, string>? ShapeTextEdited;  // (shapeIndex, newText)

    private PresentationModel? _model;
    private int                _slideIndex;
    private double             _zoom = 1.0;

    // Desired display size at zoom=1 (fit-to-window baseline)
    private const double ThumbFitW = 960.0;
    private const double ThumbFitH = 540.0;

    public SlideEditorCanvas()
    {
        InitializeComponent();
        Loaded  += (_, _) => UpdateViewboxSize();
        SizeChanged += (_, _) => { if (_zoom == 1.0) FitToWindow(); };
    }

    // ── Public API ─────────────────────────────────────────────────────

    public void ShowSlide(PresentationModel model, int slideIndex)
    {
        _model      = model;
        _slideIndex = slideIndex;
        Rebuild();
    }

    public void Invalidate() => Rebuild();

    public void ZoomIn()  { _zoom = Math.Min(3.0, Math.Round(_zoom + 0.1, 1)); ApplyZoom(); }
    public void ZoomOut() { _zoom = Math.Max(0.25, Math.Round(_zoom - 0.1, 1)); ApplyZoom(); }
    public void FitToWindow() { _zoom = 1.0; ApplyZoom(); }
    public void SetZoom(double zoom) { _zoom = Math.Max(0.25, Math.Min(3.0, zoom)); ApplyZoom(); }
    public double ZoomFactor => _zoom;

    // ── Rebuild ────────────────────────────────────────────────────────

    private void Rebuild()
    {
        SlideCanvas.Children.Clear();
        NoDocMsg.Visibility = Visibility.Collapsed;

        if (_model is null || !_model.IsOpen)
        {
            NoDocMsg.Visibility = Visibility.Visible;
            return;
        }

        var slidePart = _model.GetSlidePart(_slideIndex);
        if (slidePart is null) return;

        // Build canvas at native pixel size
        var canvas = SlideRenderer.BuildCanvas(slidePart, _model.SlideWidth, _model.SlideHeight);

        // Swap into our Viewbox child
        SlideViewbox.Child = canvas;
        SlideCanvas        = canvas;   // keep reference for future reuse

        // Enable double-click to edit text
        canvas.MouseLeftButtonDown += Canvas_MouseLeftButtonDown;

        UpdateViewboxSize();
    }

    // ── Zoom & sizing ──────────────────────────────────────────────────

    private void ApplyZoom()
    {
        UpdateViewboxSize();
    }

    private void UpdateViewboxSize()
    {
        if (_model is null || !_model.IsOpen)
        {
            SlideViewbox.Width  = ThumbFitW;
            SlideViewbox.Height = ThumbFitH;
            return;
        }

        double aspect = (double)_model.SlideWidth / _model.SlideHeight;

        if (_zoom == 1.0)
        {
            // Fit inside available area
            double availW = Math.Max(200, Scroller.ActualWidth  - 80);
            double availH = Math.Max(120, Scroller.ActualHeight - 80);

            double w = availW;
            double h = w / aspect;
            if (h > availH)
            {
                h = availH;
                w = h * aspect;
            }
            SlideViewbox.Width  = w;
            SlideViewbox.Height = h;
        }
        else
        {
            // Manual zoom: base = native 96-dpi pixel size
            double baseW = _model.SlideWidth  / 914400.0 * 96.0;
            double baseH = _model.SlideHeight / 914400.0 * 96.0;
            SlideViewbox.Width  = baseW * _zoom;
            SlideViewbox.Height = baseH * _zoom;
        }
    }

    // ── Mouse / editing ────────────────────────────────────────────────

    private void Canvas_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ClickCount < 2) return;

        var canvas    = (Canvas)sender;
        var clickPos  = e.GetPosition(canvas);

        // Find topmost UIElement whose bounds contain the click
        int idx = -1;
        for (int i = canvas.Children.Count - 1; i >= 0; i--)
        {
            if (canvas.Children[i] is UIElement el)
            {
                var topLeft = new Point(
                    Canvas.GetLeft(el),
                    Canvas.GetTop(el));
                var bounds = new Rect(topLeft,
                    new Size(el is FrameworkElement fe ? fe.Width : 0,
                             el is FrameworkElement fe2 ? fe2.Height : 0));
                if (bounds.Contains(clickPos))
                {
                    idx = i;
                    break;
                }
            }
        }
        if (idx < 0) return;

        // Find text in that element and open inline editor
        OpenTextEditor(canvas, idx);
    }

    private void OpenTextEditor(Canvas canvas, int shapeIndex)
    {
        if (canvas.Children[shapeIndex] is not Grid container) return;

        // Collect current text from TextBlocks
        string currentText = string.Join("\n",
            FindTextBlocks(container).Select(tb => tb.Inlines
                .OfType<Run>().Select(r => r.Text).Aggregate("", (a, b) => a + b)));

        double left   = Canvas.GetLeft(container);
        double top    = Canvas.GetTop(container);
        double width  = container.Width;
        double height = container.Height;

        var editor = new TextBox
        {
            AcceptsReturn   = true,
            TextWrapping    = TextWrapping.Wrap,
            Text            = currentText,
            Background      = new SolidColorBrush(Color.FromArgb(220, 255, 255, 255)),
            BorderBrush     = new SolidColorBrush(Color.FromRgb(0, 120, 212)),
            BorderThickness = new Thickness(2),
            Width           = width,
            Height          = height,
            FontSize        = 14,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
        };

        Canvas.SetLeft(editor, left);
        Canvas.SetTop(editor, top);
        canvas.Children.Add(editor);
        editor.Focus();
        editor.SelectAll();

        editor.LostFocus += (_, _) =>
        {
            ShapeTextEdited?.Invoke(shapeIndex, editor.Text);
            canvas.Children.Remove(editor);
        };
        editor.KeyDown += (_, e) =>
        {
            if (e.Key == Key.Escape)
            {
                canvas.Children.Remove(editor);
                e.Handled = true;
            }
        };
    }

    private static IEnumerable<TextBlock> FindTextBlocks(DependencyObject parent)
    {
        for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
            if (child is TextBlock tb) yield return tb;
            foreach (var nested in FindTextBlocks(child)) yield return nested;
        }
    }
}
