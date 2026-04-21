using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using PPEditer.Models;

namespace PPEditer.Services;

/// <summary>
/// Bi-directional conversion between PPTX TextBody and WPF FlowDocument.
/// </summary>
public static class PptxConverter
{
    private const double WpfDpi      = 96.0;
    private const double PointPerInch = 72.0;
    private const double EmuPerInch   = 914400.0;
    private const double EmuPerCm     = EmuPerInch / 2.54;

    // ── PPTX → FlowDocument ───────────────────────────────────────────

    // Accepts Presentation.TextBody or Drawing.TextBody (both inherit OpenXmlCompositeElement)
    public static FlowDocument ToFlowDocument(OpenXmlCompositeElement body)
    {
        var doc = new FlowDocument { FontFamily = new FontFamily("맑은 고딕") };

        // Read body vertical anchor and store in doc.Tag
        var bodyPr = body.GetFirstChild<A.BodyProperties>();
        var anchor = bodyPr?.Anchor?.InnerText switch
        {
            "ctr" => VertAnchor.Middle,
            "b"   => VertAnchor.Bottom,
            _     => VertAnchor.Top,
        };
        doc.Tag = anchor;

        foreach (var aPara in body.Elements<A.Paragraph>())
        {
            var para = new Paragraph { Margin = new Thickness(0) };

            var pPr = aPara.ParagraphProperties;

            // Horizontal alignment — SDK v3 struct enum: must use == not switch
            var av = pPr?.Alignment?.Value;
            if (av == A.TextAlignmentTypeValues.Center)
                para.TextAlignment = TextAlignment.Center;
            else if (av == A.TextAlignmentTypeValues.Right)
                para.TextAlignment = TextAlignment.Right;
            else if (av?.InnerText == "dist" || av?.InnerText == "just" || av?.InnerText == "thaiDist")
                para.TextAlignment = TextAlignment.Justify;
            else
                para.TextAlignment = TextAlignment.Left;

            // Left indent (marL in EMU → WPF px)
            if (pPr?.LeftMargin?.HasValue == true)
            {
                double leftPx = pPr.LeftMargin.Value / EmuPerInch * WpfDpi;
                para.Margin = new Thickness(leftPx, para.Margin.Top, 0, para.Margin.Bottom);
            }

            // First-line indent (indent in EMU → WPF px, negative = hanging)
            if (pPr?.Indent?.HasValue == true)
                para.TextIndent = pPr.Indent.Value / EmuPerInch * WpfDpi;

            // Line spacing
            var lnSpc = pPr?.GetFirstChild<A.LineSpacing>();
            if (lnSpc?.SpacingPercent?.Val?.HasValue == true)
            {
                double pct = lnSpc.SpacingPercent.Val.Value / 1000.0;  // val is 1/1000 of %
                if (pct > 0)
                    para.LineHeight = para.FontSize > 0
                        ? para.FontSize * pct / 100.0
                        : double.NaN;
                // Store raw pct in Tag (overrides VertAnchor tag - use ExtraParaProps instead)
                para.Tag = pct;
            }

            // Space before (spcPts val = 1/100 pt)
            var spcBef = pPr?.GetFirstChild<A.SpaceBefore>();
            if (spcBef?.SpacingPoints?.Val?.HasValue == true)
            {
                double ptVal = spcBef.SpacingPoints.Val.Value / 100.0;
                double px = ptVal * WpfDpi / PointPerInch;
                para.Margin = new Thickness(para.Margin.Left, px, 0, para.Margin.Bottom);
            }

            // Space after (spcPts val = 1/100 pt)
            var spcAft = pPr?.GetFirstChild<A.SpaceAfter>();
            if (spcAft?.SpacingPoints?.Val?.HasValue == true)
            {
                double ptVal = spcAft.SpacingPoints.Val.Value / 100.0;
                double px = ptVal * WpfDpi / PointPerInch;
                para.Margin = new Thickness(para.Margin.Left, para.Margin.Top, 0, px);
            }

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
        bool hasUnderline = uv is { Length: > 0 } && uv != "none";

        // Strikethrough: check strike InnerText
        var sv = rProps.Strike?.InnerText;
        bool hasStrike = sv is "sngStrike" or "dblStrike";

        var decos = new TextDecorationCollection();
        if (hasUnderline) decos.Add(TextDecorations.Underline[0]);
        if (hasStrike)    decos.Add(TextDecorations.Strikethrough[0]);

        bool hasOutline = rProps.GetFirstChild<A.Outline>() is not null;

        if (decos.Count > 0)
            run.TextDecorations = decos;

        // Color (foreground)
        var solid = rProps.GetFirstChild<A.SolidFill>();
        if (solid is not null)
        {
            var c = ResolveHexColor(solid);
            if (c.HasValue) run.Foreground = new SolidColorBrush(c.Value);
        }

        // Background highlight
        var highlight = rProps.GetFirstChild<A.Highlight>();
        if (highlight is not null)
        {
            var hx = highlight.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value;
            if (hx is not null && hx.Length >= 6)
            {
                var hc = Color.FromRgb(
                    Convert.ToByte(hx[..2], 16),
                    Convert.ToByte(hx[2..4], 16),
                    Convert.ToByte(hx[4..6], 16));
                run.Background = new SolidColorBrush(hc);
            }
        }

        // Script (superscript/subscript via baseline attribute)
        if (rProps.Baseline?.HasValue == true)
        {
            int baseline = rProps.Baseline.Value;
            if (baseline > 0)
                System.Windows.Documents.Typography.SetVariants(run, FontVariants.Superscript);
            else if (baseline < 0)
                System.Windows.Documents.Typography.SetVariants(run, FontVariants.Subscript);
        }

        // Spacing, underline color, outline → store in Tag
        int spacing = rProps.Spacing?.Value ?? 0;
        RgbColor? ulColor = null;
        var ulFill = rProps.GetFirstChild<A.UnderlineFill>()?.GetFirstChild<A.SolidFill>();
        if (ulFill is not null)
        {
            var ux = ulFill.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value;
            if (ux is not null && ux.Length >= 6)
                ulColor = new RgbColor(
                    Convert.ToByte(ux[..2], 16),
                    Convert.ToByte(ux[2..4], 16),
                    Convert.ToByte(ux[4..6], 16));
        }

        run.Tag = new CharExtraProps
        {
            SpacingPt100   = spacing,
            UnderlineColor = ulColor,
            HasOutline     = hasOutline,
            HasOverline    = false,
        };

        return run;
    }

    // ── FlowDocument → PptxParagraph list ────────────────────────────

    public record PptxRun(
        string Text, string? FontFamily, double? FontSizePt,
        bool Bold, bool Italic, bool Underline,
        bool Strikethrough, ScriptKind Script, int SpacingPt100,
        Color? Color, Color? BackColor, Color? UnderlineColor, bool HasOutline);

    /// <summary>Carries both paragraph-level formatting and runs.</summary>
    public record PptxParagraph(
        TextAlignment Alignment,
        int    MarginLeftEmu,      // a:pPr @marL
        int    TextIndentEmu,      // a:pPr @indent (negative = hanging)
        int    LineSpacePct1000,   // a:lnSpc/spcPct @val  (0 = auto)
        int    SpaceBeforePt100,   // a:spcBef/spcPts @val (0 = auto)
        int    SpaceAfterPt100,    // a:spcAft/spcPts @val (0 = auto)
        IReadOnlyList<PptxRun> Runs);

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

                Color? backColor = null;
                if (inline.Background is SolidColorBrush bgb && bgb.Color.A > 0)
                    backColor = bgb.Color;

                string? family = null;
                try { family = inline.FontFamily?.Source; } catch { }

                var decos = inline.TextDecorations;
                bool hasUnderline  = decos?.Any(d => d.Location == TextDecorationLocation.Underline)     == true;
                bool hasStrike     = decos?.Any(d => d.Location == TextDecorationLocation.Strikethrough) == true;

                var fv = inline.GetValue(Typography.VariantsProperty) is FontVariants variants
                    ? variants : FontVariants.Normal;
                var script = fv == FontVariants.Superscript ? ScriptKind.Superscript :
                             fv == FontVariants.Subscript   ? ScriptKind.Subscript   : ScriptKind.None;

                int spacing = 0;
                Color? ulColor = null;
                bool hasOutline = false;
                if (inline.Tag is CharExtraProps ex)
                {
                    spacing    = ex.SpacingPt100;
                    hasOutline = ex.HasOutline;
                    if (ex.UnderlineColor.HasValue)
                        ulColor = Color.FromRgb(ex.UnderlineColor.Value.R,
                                                ex.UnderlineColor.Value.G,
                                                ex.UnderlineColor.Value.B);
                }

                runs.Add(new PptxRun(
                    inline.Text,
                    string.IsNullOrEmpty(family) ? null : family,
                    sizePt,
                    inline.FontWeight == FontWeights.Bold,
                    inline.FontStyle  == FontStyles.Italic,
                    hasUnderline,
                    hasStrike,
                    script,
                    spacing,
                    color == Colors.Black ? null : color,
                    backColor,
                    ulColor,
                    hasOutline));
            }

            // Paragraph-level properties
            int marLeftEmu  = (int)Math.Round(block.Margin.Left  / WpfDpi * EmuPerInch);
            int indentEmu   = (int)Math.Round(block.TextIndent   / WpfDpi * EmuPerInch);
            int spcBefPt100 = (int)Math.Round(block.Margin.Top   * PointPerInch / WpfDpi * 100);
            int spcAftPt100 = (int)Math.Round(block.Margin.Bottom * PointPerInch / WpfDpi * 100);

            // Line spacing: stored in Paragraph.Tag as pct double
            int lnSpcPct1000 = 0;
            if (block.Tag is double lsPct && lsPct > 0)
                lnSpcPct1000 = (int)Math.Round(lsPct * 1000);
            else if (block.LineHeight > 0 && !double.IsNaN(block.LineHeight))
            {
                // approximate: compare LineHeight to default (assume 14pt baseline)
                double defaultPx = 14.0 * WpfDpi / PointPerInch;
                lnSpcPct1000 = (int)Math.Round(block.LineHeight / defaultPx * 100000);
            }

            result.Add(new PptxParagraph(
                block.TextAlignment,
                marLeftEmu, indentEmu,
                lnSpcPct1000, spcBefPt100, spcAftPt100,
                runs));
        }

        return result;
    }

    // ── Helpers ───────────────────────────────────────────────────────

    private static double DefaultFontSize(A.Paragraph para)
    {
        var first = para.Elements<A.Run>().FirstOrDefault();
        if (first?.RunProperties?.FontSize?.Value is int sz)
            return sz / 100.0 * WpfDpi / PointPerInch;
        return 18.0 * WpfDpi / PointPerInch;
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

    // ── Extra character properties (stored as Run.Tag) ─────────────────

    /// <summary>Extra run properties that cannot be expressed via WPF Run properties.</summary>
    public sealed class CharExtraProps
    {
        public int       SpacingPt100   { get; set; }
        public RgbColor? UnderlineColor { get; set; }
        public bool      HasOutline     { get; set; }
        public bool      HasOverline    { get; set; }
    }
}
