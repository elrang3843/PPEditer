using System.Globalization;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using DocumentFormat.OpenXml.Packaging;
using PPEditer.Models;
using PPEditer.Rendering;

namespace PPEditer.Export;

public enum PrintLayout { SlideOnly, SlideWithNotes, Handout3 }

public static class PrintExporter
{
    public static void Print(PresentationModel model, PrintLayout layout, bool blackWhite)
    {
        var dlg = new System.Windows.Controls.PrintDialog();
        if (dlg.ShowDialog() != true) return;
        var size = new Size(dlg.PrintableAreaWidth, dlg.PrintableAreaHeight);
        dlg.PrintDocument(
            new SlidePaginator(model, layout, blackWhite, size),
            model.FileName);
    }
}

internal sealed class SlidePaginator : DocumentPaginator
{
    private readonly PresentationModel _model;
    private readonly PrintLayout       _layout;
    private readonly bool              _blackWhite;
    private readonly Size              _pageSize;

    private readonly DocProperties _docProps;

    internal SlidePaginator(PresentationModel model, PrintLayout layout,
                            bool blackWhite, Size pageSize)
    {
        _model      = model;
        _layout     = layout;
        _blackWhite = blackWhite;
        _pageSize   = pageSize;
        _docProps   = model.GetDocProperties();
    }

    public override bool                      IsPageCountValid => true;
    public override int                       PageCount =>
        _layout == PrintLayout.Handout3
            ? (_model.SlideCount + 2) / 3
            : _model.SlideCount;
    public override Size                      PageSize  { get => _pageSize; set { } }
    public override IDocumentPaginatorSource? Source    => null;

    public override DocumentPage GetPage(int pageNumber)
    {
        var v = new DrawingVisual();
        using (var dc = v.RenderOpen())
        {
            dc.DrawRectangle(Brushes.White, null, new Rect(_pageSize));
            if (_layout == PrintLayout.Handout3)
                DrawHandout(dc, pageNumber * 3);
            else
                DrawSlidePage(dc, pageNumber);
        }
        return new DocumentPage(v, _pageSize, new Rect(_pageSize), new Rect(_pageSize));
    }

    // ── Slide-only / Slide+Notes ─────────────────────────────────────

    private void DrawSlidePage(DrawingContext dc, int idx)
    {
        var part = _model.GetSlidePart(idx);
        if (part is null) return;

        double notesH  = _layout == PrintLayout.SlideWithNotes ? _pageSize.Height * 0.35 : 0.0;
        double availH  = _pageSize.Height - notesH;
        double aspect  = _model.SlideWidth / (double)_model.SlideHeight;
        var (imgW, imgH) = FitRect(_pageSize.Width, availH, aspect);

        var bmp    = RenderSlide(part, imgW, imgH);
        var imgRect = new Rect((_pageSize.Width - imgW) / 2.0, 0, imgW, imgH);
        dc.DrawImage(bmp, imgRect);

        if (_docProps.WatermarkShowOnPrint)
            WatermarkRenderer.DrawOnContext(dc, _docProps.WatermarkText,
                _docProps.WatermarkKind, imgRect);

        if (notesH > 0)
        {
            dc.DrawLine(new Pen(Brushes.LightGray, 0.5),
                        new Point(0, availH), new Point(_pageSize.Width, availH));
            string notes = _model.GetSlideNotes(idx);
            if (!string.IsNullOrWhiteSpace(notes))
                DrawNoteText(dc, notes, new Rect(8, availH + 6, _pageSize.Width - 16, notesH - 10));
        }
    }

    // ── Handout (3 slides per page) ──────────────────────────────────

    private void DrawHandout(DrawingContext dc, int startIdx)
    {
        double slotH  = _pageSize.Height / 3.0;
        double aspect = _model.SlideWidth / (double)_model.SlideHeight;
        double thumbW = Math.Min(_pageSize.Width * 0.55, (slotH - 12) * aspect);
        double thumbH = thumbW / aspect;
        double linesX = thumbW + 16;
        double linesW = _pageSize.Width - linesX - 6;

        for (int j = 0; j < 3 && startIdx + j < _model.SlideCount; j++)
        {
            var part = _model.GetSlidePart(startIdx + j);
            if (part is null) continue;

            double top     = j * slotH + (slotH - thumbH) / 2.0;
            var    bmp     = RenderSlide(part, thumbW, thumbH);
            var    thumbRc = new Rect(4, top, thumbW, thumbH);
            dc.DrawImage(bmp, thumbRc);
            if (_docProps.WatermarkShowOnPrint)
                WatermarkRenderer.DrawOnContext(dc, _docProps.WatermarkText,
                    _docProps.WatermarkKind, thumbRc);

            // Draw rule lines for handwritten notes
            const int lineCount = 6;
            double lineStep = (slotH - 12) / lineCount;
            for (int ln = 1; ln <= lineCount; ln++)
            {
                double ly = j * slotH + 6 + ln * lineStep;
                dc.DrawLine(new Pen(Brushes.LightGray, 0.5),
                            new Point(linesX, ly), new Point(linesX + linesW, ly));
            }
        }
    }

    // ── Helpers ──────────────────────────────────────────────────────

    private BitmapSource RenderSlide(SlidePart part, double w, double h)
    {
        int bmpW = Math.Max(1, (int)(w * 150.0 / 96.0));
        int bmpH = Math.Max(1, (int)(h * 150.0 / 96.0));
        var bmp  = SlideRenderer.RenderToBitmap(part, _model.SlideWidth, _model.SlideHeight, bmpW, bmpH);
        return _blackWhite ? ToGrayscale(bmp) : bmp;
    }

    private static void DrawNoteText(DrawingContext dc, string text, Rect bounds)
    {
        var ft = new FormattedText(
            text,
            CultureInfo.CurrentCulture,
            FlowDirection.LeftToRight,
            new Typeface("맑은 고딕, Segoe UI"),
            11, Brushes.Black, 1.0)
        {
            MaxTextWidth  = bounds.Width,
            MaxTextHeight = bounds.Height,
            Trimming      = TextTrimming.CharacterEllipsis,
        };
        dc.DrawText(ft, bounds.Location);
    }

    private static (double w, double h) FitRect(double maxW, double maxH, double aspect) =>
        maxW / maxH > aspect ? (maxH * aspect, maxH) : (maxW, maxW / aspect);

    private static BitmapSource ToGrayscale(BitmapSource src)
    {
        var gray = new FormatConvertedBitmap(src, PixelFormats.Gray8, BitmapPalettes.Gray256, 0);
        return new FormatConvertedBitmap(gray, PixelFormats.Bgra32, null, 0);
    }
}
