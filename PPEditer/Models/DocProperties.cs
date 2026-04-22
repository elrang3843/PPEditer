namespace PPEditer.Models;

public enum WatermarkKind { None, Diagonal, Horizontal }

public sealed class DocProperties
{
    public string Title   { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public string Author  { get; set; } = string.Empty;
    public string Manager { get; set; } = string.Empty;
    public string Company { get; set; } = string.Empty;

    public WatermarkKind WatermarkKind        { get; set; } = WatermarkKind.None;
    public string        WatermarkText        { get; set; } = string.Empty;
    public bool          WatermarkShowOnPrint { get; set; } = true;
    public bool          WatermarkShowOnSlide { get; set; } = true;

    // Auto-managed — displayed read-only in the dialog
    public string    LastModifiedBy { get; set; } = string.Empty;
    public DateTime? Created        { get; set; }
    public DateTime? Modified       { get; set; }
    public int       Revision       { get; set; }
}
