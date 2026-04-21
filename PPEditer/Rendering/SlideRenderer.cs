using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PPEditer.Rendering;

/// <summary>
/// Converts a SlidePart into WPF visuals.
///
/// Strategy: build a Canvas at the slide's native WPF-pixel size (1 EMU = 1/914400 in = 96/914400 px),
/// then wrap it in a Viewbox for display scaling — no manual font-size arithmetic needed.
/// </summary>
public static class SlideRenderer
{
    private const double EmuPerInch = 914400.0;
    private const double WpfDpi     = 96.0;

    // ── Public API ────────────────────────────────────────────────────

    /// <summary>
    /// Build a Canvas representing <paramref name="slidePart"/> at native pixel size.
    /// Wrap the returned Canvas in a Viewbox to resize it for display.
    /// </summary>
    public static Canvas BuildCanvas(SlidePart slidePart, long slideW, long slideH)
    {
        double nativeW = EmuToPx(slideW);
        double nativeH = EmuToPx(slideH);

        var canvas = new Canvas
        {
            Width      = nativeW,
            Height     = nativeH,
            Background = GetSlideBrush(slidePart, nativeW, nativeH),
            ClipToBounds = true,
        };

        var tree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (tree is null) return canvas;

        foreach (var element in tree.Elements<OpenXmlCompositeElement>())
        {
            var child = BuildElement(element, slidePart);
            if (child is not null)
                canvas.Children.Add(child);
        }

        return canvas;
    }

    /// <summary>Render a slide to a <see cref="BitmapSource"/> of size <paramref name="px"/>×<paramref name="py"/>.</summary>
    public static BitmapSource RenderToBitmap(SlidePart slidePart, long slideW, long slideH,
                                              int px, int py)
    {
        var canvas = BuildCanvas(slidePart, slideW, slideH);
        var vb     = new Viewbox { Stretch = Stretch.Uniform, Child = canvas };
        vb.Measure(new Size(px, py));
        vb.Arrange(new Rect(0, 0, px, py));
        vb.UpdateLayout();

        var rtb = new RenderTargetBitmap(px, py, WpfDpi, WpfDpi, PixelFormats.Pbgra32);
        rtb.Render(vb);
        rtb.Freeze();
        return rtb;
    }

    // ── Background ────────────────────────────────────────────────────

    private static Brush GetSlideBrush(SlidePart slidePart, double w, double h)
    {
        try
        {
            var bg   = slidePart.Slide.CommonSlideData?.Background;
            var fill = bg?.BackgroundProperties?.GetFirstChild<A.SolidFill>();
            if (fill is not null)
            {
                var c = ResolveColor(fill, slidePart);
                if (c.HasValue) return new SolidColorBrush(c.Value);
            }
        }
        catch { /* fall through */ }
        return Brushes.White;
    }

    // ── Element dispatch ──────────────────────────────────────────────

    private static UIElement? BuildElement(OpenXmlCompositeElement element, SlidePart slidePart)
    {
        return element switch
        {
            Shape   shape   => BuildShape(shape, slidePart),
            Picture picture => BuildPicture(picture, slidePart),
            _               => null,
        };
    }

    // ── Shape (text boxes, rectangles, placeholders) ──────────────────

    private static UIElement? BuildShape(Shape shape, SlidePart slidePart)
    {
        var (left, top, width, height) = GetTransform(shape.ShapeProperties);
        if (width <= 0 || height <= 0) return null;

        var container = new Grid { Width = width, Height = height, ClipToBounds = true };

        // Fill
        var fillBrush = GetShapeFill(shape, slidePart);
        if (fillBrush is not null)
        {
            var rect = new System.Windows.Shapes.Rectangle
            {
                Width  = width,
                Height = height,
                Fill   = fillBrush,
            };
            ApplyBorder(rect, shape, slidePart);
            container.Children.Add(rect);
        }

        // Text
        if (shape.TextBody is not null)
        {
            var tb = BuildTextBlock(shape.TextBody, width, height);
            container.Children.Add(tb);
        }

        Canvas.SetLeft(container, left);
        Canvas.SetTop(container, top);
        return container;
    }

    private static Brush? GetShapeFill(Shape shape, SlidePart slidePart)
    {
        var spPr = shape.ShapeProperties;
        if (spPr is null) return null;

        var solid = spPr.GetFirstChild<A.SolidFill>();
        if (solid is not null)
        {
            var c = ResolveColor(solid, slidePart);
            if (c.HasValue) return new SolidColorBrush(c.Value);
        }

        // No-fill or unknown → transparent
        if (spPr.GetFirstChild<A.NoFill>() is not null) return null;
        return null;
    }

    private static void ApplyBorder(System.Windows.Shapes.Rectangle rect,
                                    Shape shape, SlidePart slidePart)
    {
        var ln = shape.ShapeProperties?.GetFirstChild<A.Outline>();
        if (ln is null) return;
        var solid = ln.GetFirstChild<A.SolidFill>();
        if (solid is null) return;
        var c = ResolveColor(solid, slidePart);
        if (!c.HasValue) return;
        rect.Stroke          = new SolidColorBrush(c.Value);
        rect.StrokeThickness = ln.Width.HasValue ? EmuToPx(ln.Width.Value) : 1;
    }

    // ── Text body ─────────────────────────────────────────────────────

    private static UIElement BuildTextBlock(A.TextBody body, double maxW, double maxH)
    {
        var panel = new StackPanel { Orientation = Orientation.Vertical };

        // Vertical alignment
        var anchor = body.BodyProperties?.Anchor?.Value;
        var outer  = new Grid { Width = maxW, Height = maxH };

        foreach (var para in body.Elements<A.Paragraph>())
        {
            var tb = BuildParagraph(para);
            tb.Width        = maxW;
            tb.TextWrapping = TextWrapping.Wrap;
            panel.Children.Add(tb);
        }

        outer.Children.Add(panel);
        switch (anchor)
        {
            case A.TextAnchoringTypeValues.Center:
                panel.VerticalAlignment = VerticalAlignment.Center;
                break;
            case A.TextAnchoringTypeValues.Bottom:
                panel.VerticalAlignment = VerticalAlignment.Bottom;
                break;
            default:
                panel.VerticalAlignment = VerticalAlignment.Top;
                break;
        }
        return outer;
    }

    private static TextBlock BuildParagraph(A.Paragraph para)
    {
        var tb = new TextBlock { FontFamily = new FontFamily("맑은 고딕") };

        // Alignment
        tb.TextAlignment = (para.ParagraphProperties?.Alignment?.Value) switch
        {
            A.TextAlignmentTypeValues.Center  => TextAlignment.Center,
            A.TextAlignmentTypeValues.Right   => TextAlignment.Right,
            A.TextAlignmentTypeValues.Justify => TextAlignment.Justify,
            _                                 => TextAlignment.Left,
        };

        // Runs
        foreach (var run in para.Elements<A.Run>())
        {
            var inline = BuildRun(run);
            tb.Inlines.Add(inline);
        }

        // Line break if no runs
        if (!para.Elements<A.Run>().Any())
            tb.Inlines.Add(new Run(" ") { FontSize = 14 });

        return tb;
    }

    private static Run BuildRun(A.Run run)
    {
        var text   = run.Text?.Text ?? string.Empty;
        var rProps = run.RunProperties;
        var inline = new Run(text);

        if (rProps is null) return inline;

        // Font family
        var latin = rProps.GetFirstChild<A.LatinFont>();
        if (latin?.Typeface is not null && latin.Typeface != "+mj-lt" && latin.Typeface != "+mn-lt")
            inline.FontFamily = new FontFamily(latin.Typeface);

        // Font size: OpenXml stores hundredths of a point (e.g. 2400 = 24 pt)
        if (rProps.FontSize.HasValue)
        {
            double ptSize  = rProps.FontSize.Value / 100.0;
            inline.FontSize = ptSize * WpfDpi / 72.0;   // pt → WPF pixels
        }

        if (rProps.Bold?.Value == true)         inline.FontWeight = FontWeights.Bold;
        if (rProps.Italic?.Value == true)       inline.FontStyle  = FontStyles.Italic;
        if (rProps.Underline?.Value == true)
            inline.TextDecorations = TextDecorations.Underline;

        // Color
        var solid = rProps.GetFirstChild<A.SolidFill>();
        if (solid is not null)
        {
            var c = ResolveColor(solid, null);
            if (c.HasValue) inline.Foreground = new SolidColorBrush(c.Value);
        }

        return inline;
    }

    // ── Picture ───────────────────────────────────────────────────────

    private static UIElement? BuildPicture(Picture picture, SlidePart slidePart)
    {
        var (left, top, width, height) = GetShapeTransformFromPicture(picture);
        if (width <= 0 || height <= 0) return null;

        var rId = picture.BlipFill?.Blip?.Embed?.Value;
        if (rId is null) return null;

        BitmapImage? bmp = null;
        try
        {
            var imgPart = slidePart.GetPartById(rId) as ImagePart;
            if (imgPart is not null)
            {
                using var imgStream = imgPart.GetStream();
                var ms = new MemoryStream();
                imgStream.CopyTo(ms);
                ms.Position = 0;
                bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.StreamSource  = ms;
                bmp.CacheOption   = BitmapCacheOption.OnLoad;
                bmp.EndInit();
                bmp.Freeze();
            }
        }
        catch { /* image unavailable */ }

        var img = new Image
        {
            Width   = width,
            Height  = height,
            Stretch = Stretch.Fill,
            Source  = bmp,
        };
        Canvas.SetLeft(img, left);
        Canvas.SetTop(img, top);
        return img;
    }

    // ── Transform helpers ─────────────────────────────────────────────

    private static (double left, double top, double width, double height)
        GetTransform(A.ShapeProperties? spPr)
    {
        var xfrm = spPr?.Transform2D;
        return (
            EmuToPx(xfrm?.Offset?.X ?? 0),
            EmuToPx(xfrm?.Offset?.Y ?? 0),
            EmuToPx(xfrm?.Extents?.Cx ?? 0),
            EmuToPx(xfrm?.Extents?.Cy ?? 0));
    }

    private static (double left, double top, double width, double height)
        GetShapeTransformFromPicture(Picture picture)
    {
        var xfrm = picture.ShapeProperties?.Transform2D;
        return (
            EmuToPx(xfrm?.Offset?.X ?? 0),
            EmuToPx(xfrm?.Offset?.Y ?? 0),
            EmuToPx(xfrm?.Extents?.Cx ?? 0),
            EmuToPx(xfrm?.Extents?.Cy ?? 0));
    }

    // ── Color resolution ──────────────────────────────────────────────

    private static Color? ResolveColor(A.SolidFill fill, SlidePart? _)
    {
        // RGB hex (most common)
        var hex = fill.GetFirstChild<A.RgbColorModelHex>();
        if (hex?.Val is not null)
        {
            var v = hex.Val.Value!;
            return Color.FromRgb(
                Convert.ToByte(v[..2], 16),
                Convert.ToByte(v[2..4], 16),
                Convert.ToByte(v[4..6], 16));
        }

        // Scheme color → map to reasonable default; full resolution needs theme XML
        var scheme = fill.GetFirstChild<A.SchemeColor>();
        if (scheme is not null)
        {
            return scheme.Val?.Value switch
            {
                A.SchemeColorValues.Background1 or
                A.SchemeColorValues.Light1       => Colors.White,
                A.SchemeColorValues.Text1 or
                A.SchemeColorValues.Dark1        => Colors.Black,
                A.SchemeColorValues.Accent1      => Color.FromRgb(0x46, 0x72, 0xC4),
                A.SchemeColorValues.Accent2      => Color.FromRgb(0xED, 0x7D, 0x31),
                _                                => (Color?)null,
            };
        }

        return null;
    }

    // ── Unit conversion ───────────────────────────────────────────────

    private static double EmuToPx(long emu) => emu / EmuPerInch * WpfDpi;
}
