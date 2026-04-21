namespace PPEditer.Models;

public enum ScriptKind { None, Superscript, Subscript }

public sealed class CharStyle
{
    public string?    FontFamily     { get; set; }
    public double?    FontSizePt     { get; set; }
    public bool?      Bold           { get; set; }
    public bool?      Italic         { get; set; }
    public bool?      HasUnderline   { get; set; }
    public bool?      HasStrike      { get; set; }
    public bool?      HasOverline    { get; set; }  // WPF-only display
    public bool?      HasOutline     { get; set; }  // PPTX a:rPr/a:ln
    public ScriptKind Script         { get; set; } = ScriptKind.None;
    public int        SpacingPt100   { get; set; }  // PPTX spc attribute
    public RgbColor?  ForeColor      { get; set; }
    public RgbColor?  HighlightColor { get; set; }
    public RgbColor?  UnderlineColor { get; set; }
}
