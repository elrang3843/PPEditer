using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace PPEditer.Services;

/// <summary>
/// Bi-directional conversion between PPTX TextBody and WPF FlowDocument.
/// Preserves font family, size, bold, italic, underline, colour, and paragraph alignment.
/// </summary>
public static class PptxConverter
{
    private const double WpfDpi    = 96.0;
    private const double PointPerInch = 72.0;

    // ── PPTX → FlowDocument ───────────────────────────────────────────

    // Accepts Presentation.TextBody or Drawing.TextBody (both inherit OpenXmlCompositeElement)
    public static FlowDocument ToFlowDocument(OpenXmlCompositeElement body)
    {
        var doc = new FlowDocument { FontFamily = new FontFamily("맑은 고딕") };

        foreach (var aPara in body.Elements<A.Paragraph>())
        {
            var para = new Paragraph { Margin = new Thickness(0) };

            // Alignment — SDK v3 struct enum: cannot use switch case, must use ==
            var av = aPara.ParagraphProperties?.Alignment?.Value;
            if (av == A.TextAlignmentTypeValues.Center)
                para.TextAlignment = TextAlignment.Center;
            else if (av == A.TextAlignmentTypeValues.Right)
                para.TextAlignment = TextAlignment.Right;
            else
                para.TextAlignment = TextAlignment.Left;

            // Spacing
            if (aPara.ParagraphProperties?.SpaceBefore?.GetFirstChild<A.SpacePercent>() is A.SpacePercent sp
                && sp.Val?.Value is int spv)
                para.Margin = para.Margin with { Top = spv / 100000.0 * 16.0 };

            bool hasRuns = false;
            foreach (var aRun in aPara.Elements<A.Run>())
            {
                para.Inlines.Add(BuildInlineRun(aRun));
                hasRuns = true;
            }
            if (!hasRuns)
                para.Inlines.Add(new Run(" ") { FontSize = DefaultFontSize(aPara) });

            doc.Blocks.Add(para);
        }

        if (!doc.Blocks.Any())
            doc.Blocks.Add(new Paragraph(new Run()));

        return doc;
    }

    private static Run BuildInlineRun(A.Run aRun)
    {
        var text   = aRun.Text?.Text ?? string.Empty;
        var rProps = aRun.RunProperties;
        var run    = new Run(text);

        if (rProps is null) return run;

        // Font family (prefer EastAsian for CJK)
        var ea = rProps.GetFirstChild<A.EastAsianFont>()?.Typeface;
        var lt = rProps.GetFirstChild<A.LatinFont>()?.Typeface;
        var ff = ea ?? lt;
        if (ff is not null && ff != "+mj-lt" && ff != "+mn-lt" &&
            ff != "+mj-ea" && ff != "+mn-ea")
            run.FontFamily = new FontFamily(ff);

        // Font size (hundredths of a point)
        if (rProps.FontSize?.HasValue == true)
            run.FontSize = rProps.FontSize!.Value / 100.0 * WpfDpi / PointPerInch;

        if (rProps.Bold?.Value == true)    run.FontWeight = FontWeights.Bold;
        if (rProps.Italic?.Value == true)  run.FontStyle  = FontStyles.Italic;
        // SDK v3: Underline is EnumValue<TextUnderlineValues>, not bool — compare via InnerText
        var uv = rProps.Underline?.InnerText;
        if (uv is { Length: > 0 } && uv != "none")
            run.TextDecorations = TextDecorations.Underline;

        // Color
        var solid = rProps.GetFirstChild<A.SolidFill>();
        if (solid is not null)
        {
            var c = ResolveHexColor(solid);
            if (c.HasValue) run.Foreground = new SolidColorBrush(c.Value);
        }

        return run;
    }

    // ── FlowDocument → PptxParagraph list ────────────────────────────

    public record PptxRun(string Text, string? FontFamily, double? FontSizePt,
                          bool Bold, bool Italic, bool Underline, Color? Color);
    public record PptxParagraph(TextAlignment Alignment, IReadOnlyList<PptxRun> Runs);

    public static List<PptxParagraph> FromFlowDocument(FlowDocument doc)
    {
        var result = new List<PptxParagraph>();

        foreach (var block in doc.Blocks.OfType<Paragraph>())
        {
            var runs = new List<PptxRun>();
            foreach (var inline in block.Inlines.OfType<Run>())
            {
                double? sizeWpf = inline.FontSize > 0 ? inline.FontSize : (double?)null;
                double? sizePt  = sizeWpf.HasValue
                    ? sizeWpf.Value * PointPerInch / WpfDpi
                    : null;

                Color? color = null;
                if (inline.Foreground is SolidColorBrush scb)
                    color = scb.Color;

                string? family = null;
                try { family = inline.FontFamily?.Source; } catch { }

                runs.Add(new PptxRun(
                    inline.Text,
                    string.IsNullOrEmpty(family) ? null : family,
                    sizePt,
                    inline.FontWeight == FontWeights.Bold,
                    inline.FontStyle  == FontStyles.Italic,
                    inline.TextDecorations?.Any() == true,
                    color == Colors.Black ? null : color));
            }
            result.Add(new PptxParagraph(block.TextAlignment, runs));
        }

        return result;
    }

    // ── Helpers ───────────────────────────────────────────────────────

    private static double DefaultFontSize(A.Paragraph para)
    {
        var first = para.Elements<A.Run>().FirstOrDefault();
        if (first?.RunProperties?.FontSize?.Value is int sz)
            return sz / 100.0 * WpfDpi / PointPerInch;
        return 18.0 * WpfDpi / PointPerInch; // 18pt default
    }

    private static Color? ResolveHexColor(A.SolidFill fill)
    {
        var hex = fill.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value;
        if (hex is null || hex.Length < 6) return null;
        return Color.FromRgb(
            Convert.ToByte(hex[..2], 16),
            Convert.ToByte(hex[2..4], 16),
            Convert.ToByte(hex[4..6], 16));
    }

    private static double EmuToWpf(long emu) => emu / 914400.0 * WpfDpi;
}
