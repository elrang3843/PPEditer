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

    public static Canvas BuildCanvas(SlidePart slidePart, long slideW, long slideH)
    {
        double nativeW = EmuToPx(slideW);
        double nativeH = EmuToPx(slideH);

        var canvas = new Canvas
        {
            Width        = nativeW,
            Height       = nativeH,
            Background   = GetSlideBrush(slidePart, nativeW, nativeH),
            ClipToBounds = true,
        };

        var tree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (tree is null) return canvas;

        int treeIndex = 0;
        foreach (var element in tree.Elements<OpenXmlCompositeElement>())
        {
            var child = BuildElement(element, slidePart);
            if (child is FrameworkElement fe)
            {
                fe.Tag = treeIndex;
                canvas.Children.Add(fe);
            }
            treeIndex++;
        }

        return canvas;
    }

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
        catch { }
        return Brushes.White;
    }

    // ── Element dispatch ──────────────────────────────────────────────

    private static UIElement? BuildElement(OpenXmlCompositeElement element, SlidePart slidePart)
    {
        return element switch
        {
            Shape      shape   => BuildShape(shape, slidePart),
            Picture    picture => BuildPicture(picture, slidePart),
            GroupShape group   => BuildGroupShape(group, slidePart),
            _                  => null,
        };
    }

    private static UIElement? BuildGroupShape(GroupShape group, SlidePart slidePart)
    {
        var xfrm   = group.GroupShapeProperties?.GetFirstChild<A.TransformGroup>();
        if (xfrm is null) return null;
        double left   = EmuToPx(xfrm.Offset?.X   ?? 0);
        double top    = EmuToPx(xfrm.Offset?.Y   ?? 0);
        double width  = EmuToPx(xfrm.Extents?.Cx ?? 0);
        double height = EmuToPx(xfrm.Extents?.Cy ?? 0);
        if (width <= 0 || height <= 0) return null;

        var container = new Canvas { Width = width, Height = height };
        foreach (var child in group.Elements<OpenXmlCompositeElement>())
        {
            var el = BuildElement(child, slidePart);
            if (el is FrameworkElement fe)
            {
                Canvas.SetLeft(fe, Canvas.GetLeft(fe) - left);
                Canvas.SetTop(fe,  Canvas.GetTop(fe)  - top);
                container.Children.Add(fe);
            }
        }
        Canvas.SetLeft(container, left);
        Canvas.SetTop(container,  top);
        return container;
    }

    // ── Shape ─────────────────────────────────────────────────────────

    private static UIElement? BuildShape(Shape shape, SlidePart slidePart)
    {
        var (left, top, width, height) = GetTransform(shape.ShapeProperties);
        if (width <= 0 || height <= 0) return null;

        var spPr      = shape.ShapeProperties;
        var fillBrush = GetShapeFill(shape, slidePart);

        System.Windows.Shapes.Shape geomShape;

        var custGeom = spPr?.GetFirstChild<A.CustomGeometry>();
        if (custGeom is not null)
        {
            geomShape = new System.Windows.Shapes.Path
            {
                Data    = BuildPathGeometry(custGeom, width, height),
                Width   = width,
                Height  = height,
                Stretch = Stretch.None,
                Fill    = fillBrush ?? Brushes.Transparent,
            };
        }
        else if (spPr?.GetFirstChild<A.PresetGeometry>()?.Preset?.InnerText == "ellipse")
        {
            geomShape = new System.Windows.Shapes.Ellipse
            {
                Width  = width,
                Height = height,
                Fill   = fillBrush ?? Brushes.Transparent,
            };
        }
        else
        {
            geomShape = new System.Windows.Shapes.Rectangle
            {
                Width  = width,
                Height = height,
                Fill   = fillBrush ?? Brushes.Transparent,
            };
        }

        ApplyStroke(geomShape, shape, slidePart);

        var container = new Grid { Width = width, Height = height, ClipToBounds = true };
        container.Children.Add(geomShape);

        if (shape.TextBody is not null)
            container.Children.Add(BuildTextBlock(shape.TextBody, width, height));

        Canvas.SetLeft(container, left);
        Canvas.SetTop(container, top);
        return container;
    }

    private static PathGeometry BuildPathGeometry(A.CustomGeometry custGeom, double wpfW, double wpfH)
    {
        var pathList = custGeom.GetFirstChild<A.PathList>();
        if (pathList is null) return new PathGeometry();

        var geometry = new PathGeometry();
        foreach (var aPath in pathList.Elements<A.Path>())
        {
            long rawW  = aPath.Width?.Value  ?? 0L;
            long rawH  = aPath.Height?.Value ?? 0L;
            double pW  = rawW > 0 ? EmuToPx(rawW) : wpfW;
            double pH  = rawH > 0 ? EmuToPx(rawH) : wpfH;
            double sx  = pW > 0 ? wpfW / pW : 1;
            double sy  = pH > 0 ? wpfH / pH : 1;

            var figure  = new PathFigure { IsClosed = false, IsFilled = true };
            bool started = false;

            foreach (var el in aPath.ChildElements)
            {
                if (el is A.MoveTo moveTo)
                {
                    var pt = moveTo.Point;
                    if (pt is null) continue;
                    double px = EmuToPx(ParsePtVal(pt.X)) * sx;
                    double py = EmuToPx(ParsePtVal(pt.Y)) * sy;
                    if (started)
                    {
                        geometry.Figures.Add(figure);
                        figure = new PathFigure { IsClosed = false, IsFilled = true };
                    }
                    figure.StartPoint = new System.Windows.Point(px, py);
                    started = true;
                }
                else if (el is A.LineTo lineTo)
                {
                    var pt = lineTo.Point;
                    if (pt is null) continue;
                    figure.Segments.Add(new LineSegment(
                        new System.Windows.Point(EmuToPx(ParsePtVal(pt.X)) * sx,
                                                 EmuToPx(ParsePtVal(pt.Y)) * sy),
                        isStroked: true));
                }
                else if (el is A.CubicBezierCurveTo cubicBez)
                {
                    var pts = cubicBez.Elements<A.Point>().ToList();
                    if (pts.Count < 3) continue;
                    figure.Segments.Add(new BezierSegment(
                        new System.Windows.Point(EmuToPx(ParsePtVal(pts[0].X)) * sx, EmuToPx(ParsePtVal(pts[0].Y)) * sy),
                        new System.Windows.Point(EmuToPx(ParsePtVal(pts[1].X)) * sx, EmuToPx(ParsePtVal(pts[1].Y)) * sy),
                        new System.Windows.Point(EmuToPx(ParsePtVal(pts[2].X)) * sx, EmuToPx(ParsePtVal(pts[2].Y)) * sy),
                        isStroked: true));
                }
                else if (el is A.CloseShapePath)
                {
                    figure.IsClosed = true;
                }
            }

            if (started) geometry.Figures.Add(figure);
        }

        return geometry;
    }

    private static long ParsePtVal(StringValue? val) =>
        val?.HasValue == true && long.TryParse(val.Value, out var n) ? n : 0L;

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

        if (spPr.GetFirstChild<A.NoFill>() is not null) return null;
        return null;
    }

    private static void ApplyStroke(System.Windows.Shapes.Shape shape,
                                    Shape xmlShape, SlidePart slidePart)
    {
        var ln    = xmlShape.ShapeProperties?.GetFirstChild<A.Outline>();
        var solid = ln?.GetFirstChild<A.SolidFill>();
        if (solid is null) return;
        var c = ResolveColor(solid, slidePart);
        if (!c.HasValue) return;
        shape.Stroke          = new SolidColorBrush(c.Value);
        shape.StrokeThickness = ln?.Width?.Value is int thick ? EmuToPx((long)thick) : 1;
    }

    // ── Text body ─────────────────────────────────────────────────────

    // Accepts Presentation.TextBody or Drawing.TextBody (both inherit OpenXmlCompositeElement)
    private static UIElement BuildTextBlock(OpenXmlCompositeElement body, double maxW, double maxH)
    {
        var panel  = new StackPanel { Orientation = Orientation.Vertical };
        var anchor = body.GetFirstChild<A.BodyProperties>()?.Anchor?.Value;
        var outer  = new Grid { Width = maxW, Height = maxH };

        foreach (var para in body.Elements<A.Paragraph>())
        {
            var tb = BuildParagraph(para);
            tb.Width        = maxW;
            tb.TextWrapping = TextWrapping.Wrap;
            panel.Children.Add(tb);
        }

        outer.Children.Add(panel);

        // SDK v3: TextAnchoringTypeValues is a struct — cannot use switch case, must use ==
        if (anchor == A.TextAnchoringTypeValues.Center)
            panel.VerticalAlignment = VerticalAlignment.Center;
        else if (anchor == A.TextAnchoringTypeValues.Bottom)
            panel.VerticalAlignment = VerticalAlignment.Bottom;
        else
            panel.VerticalAlignment = VerticalAlignment.Top;

        return outer;
    }

    private static TextBlock BuildParagraph(A.Paragraph para)
    {
        var tb = new TextBlock { FontFamily = new FontFamily("맑은 고딕") };

        // SDK v3: TextAlignmentTypeValues is a struct — cannot use switch case, must use ==
        var av = para.ParagraphProperties?.Alignment?.Value;
        if (av == A.TextAlignmentTypeValues.Center)
            tb.TextAlignment = TextAlignment.Center;
        else if (av == A.TextAlignmentTypeValues.Right)
            tb.TextAlignment = TextAlignment.Right;
        else
            tb.TextAlignment = TextAlignment.Left;

        foreach (var run in para.Elements<A.Run>())
            tb.Inlines.Add(BuildRun(run));

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

        var latin = rProps.GetFirstChild<A.LatinFont>();
        if (latin?.Typeface is not null && latin.Typeface != "+mj-lt" && latin.Typeface != "+mn-lt")
            inline.FontFamily = new FontFamily(latin.Typeface);

        if (rProps.FontSize?.HasValue == true)
            inline.FontSize = rProps.FontSize!.Value / 100.0 * WpfDpi / 72.0;

        if (rProps.Bold?.Value   == true) inline.FontWeight = FontWeights.Bold;
        if (rProps.Italic?.Value == true) inline.FontStyle  = FontStyles.Italic;

        // SDK v3: Underline is EnumValue<TextUnderlineValues>, not bool — compare via InnerText
        var uv = rProps.Underline?.InnerText;
        if (uv is { Length: > 0 } && uv != "none")
            inline.TextDecorations = TextDecorations.Underline;

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
            if (slidePart.GetPartById(rId) is ImagePart imgPart)
            {
                using var imgStream = imgPart.GetStream();
                var ms = new MemoryStream();
                imgStream.CopyTo(ms);
                ms.Position = 0;
                bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.StreamSource = ms;
                bmp.CacheOption  = BitmapCacheOption.OnLoad;
                bmp.EndInit();
                bmp.Freeze();
            }
        }
        catch { }

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

    // Accepts Presentation.ShapeProperties or Drawing.ShapeProperties (both OpenXmlCompositeElement)
    private static (double left, double top, double width, double height)
        GetTransform(OpenXmlCompositeElement? spPr)
    {
        var xfrm = spPr?.GetFirstChild<A.Transform2D>();
        return (
            EmuToPx(xfrm?.Offset?.X ?? 0),
            EmuToPx(xfrm?.Offset?.Y ?? 0),
            EmuToPx(xfrm?.Extents?.Cx ?? 0),
            EmuToPx(xfrm?.Extents?.Cy ?? 0));
    }

    private static (double left, double top, double width, double height)
        GetShapeTransformFromPicture(Picture picture)
    {
        var xfrm = picture.ShapeProperties?.GetFirstChild<A.Transform2D>();
        return (
            EmuToPx(xfrm?.Offset?.X ?? 0),
            EmuToPx(xfrm?.Offset?.Y ?? 0),
            EmuToPx(xfrm?.Extents?.Cx ?? 0),
            EmuToPx(xfrm?.Extents?.Cy ?? 0));
    }

    // ── Color resolution ──────────────────────────────────────────────

    private static Color? ResolveColor(A.SolidFill fill, SlidePart? _)
    {
        var hex = fill.GetFirstChild<A.RgbColorModelHex>();
        if (hex?.Val is not null)
        {
            var v = hex.Val.Value!;
            return Color.FromRgb(
                Convert.ToByte(v[..2], 16),
                Convert.ToByte(v[2..4], 16),
                Convert.ToByte(v[4..6], 16));
        }

        // SDK v3: SchemeColorValues is a struct — cannot use switch case, must use ==
        var scheme = fill.GetFirstChild<A.SchemeColor>();
        if (scheme is not null)
        {
            var sv = scheme.Val?.Value;
            if (sv == A.SchemeColorValues.Background1 || sv == A.SchemeColorValues.Light1)
                return Colors.White;
            if (sv == A.SchemeColorValues.Text1 || sv == A.SchemeColorValues.Dark1)
                return Colors.Black;
            if (sv == A.SchemeColorValues.Accent1)
                return Color.FromRgb(0x46, 0x72, 0xC4);
            if (sv == A.SchemeColorValues.Accent2)
                return Color.FromRgb(0xED, 0x7D, 0x31);
        }

        return null;
    }

    // ── Unit conversion ───────────────────────────────────────────────

    private static double EmuToPx(long emu) => emu / EmuPerInch * WpfDpi;
}
