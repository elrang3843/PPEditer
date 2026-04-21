namespace PPEditer.Models;

public sealed class DocProperties
{
    public string Title          { get; set; } = string.Empty;
    public string Subject        { get; set; } = string.Empty;
    public string Author         { get; set; } = string.Empty;
    public string Manager        { get; set; } = string.Empty;
    public string Company        { get; set; } = string.Empty;

    // Auto-managed — displayed read-only in the dialog
    public string    LastModifiedBy { get; set; } = string.Empty;
    public DateTime? Created        { get; set; }
    public DateTime? Modified       { get; set; }
    public int       Revision       { get; set; }
}
