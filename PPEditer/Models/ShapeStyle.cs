namespace PPEditer.Models;

public enum FillKind { None, Solid }
public enum OutlineKind { None, Solid, Dash, Dot, DashDot, DashDotDot }

public readonly record struct RgbColor(byte R, byte G, byte B)
{
    public static readonly RgbColor DefaultFill    = new(0xBD, 0xD7, 0xEE);
    public static readonly RgbColor DefaultOutline = new(0x2E, 0x74, 0xB5);
    public static readonly RgbColor Black          = new(0, 0, 0);

    public static RgbColor FromHex(string hex)
    {
        hex = hex.TrimStart('#');
        if (hex.Length < 6) return Black;
        return new RgbColor(
            Convert.ToByte(hex[0..2], 16),
            Convert.ToByte(hex[2..4], 16),
            Convert.ToByte(hex[4..6], 16));
    }

    public string ToHex() => $"{R:X2}{G:X2}{B:X2}";
}

public sealed class ShapeStyle
{
    public string     Name           { get; set; } = "";
    public bool       IsPicture      { get; set; }
    public long       X              { get; set; }
    public long       Y              { get; set; }
    public long       Cx             { get; set; }
    public long       Cy             { get; set; }
    public FillKind   FillKind       { get; set; } = FillKind.Solid;
    public RgbColor   FillColor      { get; set; } = RgbColor.DefaultFill;
    public OutlineKind OutlineKind   { get; set; } = OutlineKind.Solid;
    public RgbColor   OutlineColor   { get; set; } = RgbColor.DefaultOutline;
    public double     OutlineWidthPt { get; set; } = 1.5;
}
