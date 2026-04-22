using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using PPEditer.Models;

namespace PPEditer.Rendering;

public static class WatermarkRenderer
{
    private static readonly Brush Brush;

    static WatermarkRenderer()
    {
        Brush = new SolidColorBrush(Color.FromArgb(90, 180, 20, 20));
        Brush.Freeze();
    }

    /// <summary>
    /// Builds a Viewbox overlay in slide coordinate space (slideW × slideH) so it scales
    /// identically to the slide's own Viewbox.  Returns null when no watermark is configured.
    /// </summary>
    public static UIElement? BuildOverlay(DocProperties props, double slideW, double slideH)
    {
        if (props.WatermarkKind == WatermarkKind.None || string.IsNullOrWhiteSpace(props.WatermarkText))
            return null;
        return BuildOverlay(props.WatermarkText, props.WatermarkKind, slideW, slideH);
    }

    public static UIElement BuildOverlay(string text, WatermarkKind kind, double slideW, double slideH)
    {
        double fontSize = Math.Min(slideW, slideH) * 0.18;

        var tb = new TextBlock
        {
            Text             = text,
            FontSize         = fontSize,
            FontWeight       = FontWeights.Bold,
            Foreground       = Brush,
            IsHitTestVisible = false,
        };

        // Measure to centre manually inside the canvas.
        tb.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
        double tw = tb.DesiredSize.Width;
        double th = tb.DesiredSize.Height;

        Canvas.SetLeft(tb, (slideW - tw) / 2);
        Canvas.SetTop(tb, (slideH - th) / 2);

        if (kind == WatermarkKind.Diagonal)
        {
            tb.RenderTransformOrigin = new Point(0.5, 0.5);
            tb.RenderTransform       = new RotateTransform(-40);
        }

        var canvas = new Canvas
        {
            Width            = slideW,
            Height           = slideH,
            IsHitTestVisible = false,
        };
        canvas.Children.Add(tb);

        // Wrap in a Viewbox so it scales exactly like SlideViewbox (Stretch=Uniform).
        return new Viewbox
        {
            Stretch          = Stretch.Uniform,
            Child            = canvas,
            IsHitTestVisible = false,
        };
    }

    /// <summary>Draws a watermark directly onto a DrawingContext (for print output).</summary>
    public static void DrawOnContext(DrawingContext dc, string text, WatermarkKind kind, Rect bounds)
    {
        if (kind == WatermarkKind.None || string.IsNullOrWhiteSpace(text)) return;

        double fontSize = Math.Min(bounds.Width, bounds.Height) * 0.18;
        var ft = new FormattedText(
            text,
            CultureInfo.CurrentCulture,
            FlowDirection.LeftToRight,
            new Typeface(new FontFamily("맑은 고딕, Segoe UI"),
                         FontStyles.Normal, FontWeights.Bold, FontStretches.Normal),
            fontSize, Brush, 1.0);

        double cx = bounds.X + (bounds.Width  - ft.Width)  / 2;
        double cy = bounds.Y + (bounds.Height - ft.Height) / 2;

        if (kind == WatermarkKind.Diagonal)
            dc.PushTransform(new RotateTransform(-40,
                bounds.X + bounds.Width / 2, bounds.Y + bounds.Height / 2));

        dc.DrawText(ft, new Point(cx, cy));

        if (kind == WatermarkKind.Diagonal) dc.Pop();
    }
}
