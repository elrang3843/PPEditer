namespace PPEditer.Models;

public enum HorzAlign { Left, Center, Right, Justify }
public enum VertAnchor { Top, Middle, Bottom }

public sealed class ParagraphStyle
{
    public HorzAlign  HorzAlign     { get; set; } = HorzAlign.Left;
    public VertAnchor VertAnchor    { get; set; } = VertAnchor.Top;
    /// <summary>Left indent in cm (a:pPr @marL).</summary>
    public double     MarginLeftCm  { get; set; }
    /// <summary>First-line indent in cm (a:pPr @indent). Negative = hanging.</summary>
    public double     TextIndentCm  { get; set; }
    /// <summary>Line spacing in percent (100 = single, 150 = 1.5 lines).</summary>
    public double     LineSpacePct  { get; set; } = 100.0;
    /// <summary>Space before paragraph in pt.</summary>
    public double     SpaceBeforePt { get; set; }
    /// <summary>Space after paragraph in pt.</summary>
    public double     SpaceAfterPt  { get; set; }
}
