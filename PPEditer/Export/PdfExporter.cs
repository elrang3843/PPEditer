using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PPEditer.Models;
using PPEditer.Rendering;

namespace PPEditer.Export;

/// <summary>
/// Exports a presentation to PDF by rendering each slide as a bitmap and embedding it.
/// Uses PDFsharp (MIT License).
/// </summary>
public static class PdfExporter
{
    private const double EmuPerInch = 914400.0;

    /// <summary>
    /// Export all slides in <paramref name="model"/> to a PDF file.
    /// <paramref name="progress"/> receives (currentSlide, totalSlides).
    /// Returns true on success.
    /// </summary>
    public static bool Export(PresentationModel model, string outputPath,
                              IProgress<(int current, int total)>? progress = null)
    {
        if (!model.IsOpen) return false;

        // Slide physical size in inches
        double widthIn  = model.SlideWidth  / EmuPerInch;
        double heightIn = model.SlideHeight / EmuPerInch;

        // Render at 150 DPI for acceptable quality/size balance
        const double renderDpi = 150.0;
        int bmpW = (int)(widthIn  * renderDpi);
        int bmpH = (int)(heightIn * renderDpi);

        using var pdf = new PdfDocument();
        pdf.Info.Title   = model.FileName;
        pdf.Info.Creator = "PPEditer — HANDTECH";

        for (int i = 0; i < model.SlideCount; i++)
        {
            progress?.Report((i, model.SlideCount));

            var slidePart = model.GetSlidePart(i);
            if (slidePart is null) continue;

            // Render slide → BitmapSource → PNG bytes → XImage
            var bmpSrc = SlideRenderer.RenderToBitmap(
                slidePart, model.SlideWidth, model.SlideHeight, bmpW, bmpH);

            var pngBytes = EncodeToPng(bmpSrc);

            using var ms     = new MemoryStream(pngBytes);
            var xImage        = XImage.FromStream(ms);

            // PDF page — size in points (1 pt = 1/72 in)
            var page  = pdf.AddPage();
            page.Width  = new XUnit(widthIn  * 72, XGraphicsUnit.Point);
            page.Height = new XUnit(heightIn * 72, XGraphicsUnit.Point);

            using var gfx = XGraphics.FromPdfPage(page);
            gfx.DrawImage(xImage, 0, 0, page.Width.Point, page.Height.Point);
        }

        progress?.Report((model.SlideCount, model.SlideCount));
        pdf.Save(outputPath);
        return true;
    }

    private static byte[] EncodeToPng(BitmapSource src)
    {
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(src));
        using var ms = new MemoryStream();
        encoder.Save(ms);
        return ms.ToArray();
    }
}
