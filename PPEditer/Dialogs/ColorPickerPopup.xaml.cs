using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using PPEditer.Models;

namespace PPEditer.Dialogs;

public partial class ColorPickerPopup : Window
{
    private static readonly string[] PaletteHex =
    [
        "000000", "444444", "888888", "CCCCCC", "FFFFFF", "FF0000",
        "FF8000", "FFFF00", "00CC00", "00CCCC", "0000FF", "CC00CC",
    ];

    // Mutable gradient stop for the spectrum's White→Hue horizontal gradient
    private readonly GradientStop _hueGradStop;

    private double _h, _s, _v;   // H: 0–360,  S: 0–1,  V: 0–1
    private bool   _updating;
    private bool   _specDrag, _hueDrag;

    public RgbColor SelectedColor { get; private set; }

    public ColorPickerPopup(RgbColor initial)
    {
        InitializeComponent();
        SelectedColor = initial;
        RgbToHsv(initial.R, initial.G, initial.B, out _h, out _s, out _v);

        // Build the spectrum gradient in code so the hue stop stays mutable
        _hueGradStop = new GradientStop(Colors.Red, 1.0);
        var specBrush = new LinearGradientBrush { StartPoint = new Point(0, 0), EndPoint = new Point(1, 0) };
        specBrush.GradientStops.Add(new GradientStop(Colors.White, 0));
        specBrush.GradientStops.Add(_hueGradStop);
        RectHueGrad.Fill = specBrush;

        // Sync all controls once the window is rendered (ActualWidth/Height available)
        Loaded += (_, _) => ApplyHsv();

        // Quick-palette buttons
        foreach (var hex in PaletteHex)
        {
            var c   = RgbColor.FromHex(hex);
            var btn = new Button
            {
                Background      = new SolidColorBrush(Color.FromRgb(c.R, c.G, c.B)),
                BorderBrush     = Brushes.Gray,
                BorderThickness = new Thickness(1),
                Margin          = new Thickness(1),
                ToolTip         = $"#{hex.ToUpper()}",
            };
            string h = hex;
            btn.Click += (_, _) =>
            {
                var pc = RgbColor.FromHex(h);
                RgbToHsv(pc.R, pc.G, pc.B, out _h, out _s, out _v);
                ApplyHsv();
            };
            ColorGrid.Children.Add(btn);
        }
    }

    // ── Spectrum (SV plane) mouse ────────────────────────────────────

    private void OnSpectrumDown(object s, MouseButtonEventArgs e)
    {
        _specDrag = true;
        SpectrumGrid.CaptureMouse();
        SetSvFromPoint(e.GetPosition(SpectrumGrid));
    }

    private void OnSpectrumMove(object s, MouseEventArgs e)
    {
        if (_specDrag) SetSvFromPoint(e.GetPosition(SpectrumGrid));
    }

    private void OnSpectrumUp(object s, MouseButtonEventArgs e)
    {
        _specDrag = false;
        SpectrumGrid.ReleaseMouseCapture();
    }

    private void SetSvFromPoint(Point p)
    {
        double w = SpectrumGrid.ActualWidth;
        double h = SpectrumGrid.ActualHeight;
        if (w <= 0 || h <= 0) return;
        _s = Math.Clamp(p.X / w, 0, 1);
        _v = Math.Clamp(1 - p.Y / h, 0, 1);
        ApplyHsv();
    }

    // ── Hue strip mouse ──────────────────────────────────────────────

    private void OnHueDown(object s, MouseButtonEventArgs e)
    {
        _hueDrag = true;
        HueGrid.CaptureMouse();
        SetHFromPoint(e.GetPosition(HueGrid));
    }

    private void OnHueMove(object s, MouseEventArgs e)
    {
        if (_hueDrag) SetHFromPoint(e.GetPosition(HueGrid));
    }

    private void OnHueUp(object s, MouseButtonEventArgs e)
    {
        _hueDrag = false;
        HueGrid.ReleaseMouseCapture();
    }

    private void SetHFromPoint(Point p)
    {
        double h = HueGrid.ActualHeight;
        if (h <= 0) return;
        _h = Math.Clamp(p.Y / h * 360, 0, 360);
        ApplyHsv();
    }

    // ── RGB inputs ───────────────────────────────────────────────────

    private void OnRgbSliderChanged(object s, RoutedPropertyChangedEventArgs<double> e)
    {
        if (_updating) return;
        FromRgb((byte)SlR.Value, (byte)SlG.Value, (byte)SlB.Value,
                skipSliders: true, skipTextboxes: false);
    }

    private void OnRgbTextChanged(object s, TextChangedEventArgs e)
    {
        if (_updating) return;
        if (!byte.TryParse(TbR.Text, out byte r)) return;
        if (!byte.TryParse(TbG.Text, out byte g)) return;
        if (!byte.TryParse(TbB.Text, out byte b)) return;
        FromRgb(r, g, b, skipSliders: false, skipTextboxes: true);
    }

    private void OnHexChanged(object s, TextChangedEventArgs e)
    {
        if (_updating || TbHex.Text.Length != 6) return;
        try
        {
            var c = RgbColor.FromHex(TbHex.Text);
            FromRgb(c.R, c.G, c.B, skipSliders: false, skipTextboxes: false, skipHex: true);
        }
        catch { }
    }

    private void FromRgb(byte r, byte g, byte b,
        bool skipSliders, bool skipTextboxes, bool skipHex = false)
    {
        RgbToHsv(r, g, b, out double nh, out double ns, out double nv);
        if (ns > 0.01) _h = nh;  // preserve hue for near-achromatic colors
        _s = ns;
        _v = nv;
        ApplyHsv(skipSliders, skipTextboxes, skipHex);
    }

    // ── Central sync ─────────────────────────────────────────────────

    private void ApplyHsv(bool skipSliders = false, bool skipTextboxes = false, bool skipHex = false)
    {
        HsvToRgb(_h, _s, _v, out byte r, out byte g, out byte b);
        var wpfColor = Color.FromRgb(r, g, b);

        _updating = true;
        try
        {
            SelectedColor            = new RgbColor(r, g, b);
            PreviewSwatch.Background = new SolidColorBrush(wpfColor);

            // Update spectrum's hue gradient stop
            HsvToRgb(_h, 1, 1, out byte hr, out byte hg, out byte hb);
            _hueGradStop.Color = Color.FromRgb(hr, hg, hb);

            // Spectrum crosshair cursor
            double sw = SpectrumGrid.ActualWidth;
            double sh = SpectrumGrid.ActualHeight;
            if (sw > 0 && sh > 0)
            {
                Canvas.SetLeft(SpectrumCursor, _s * sw - 6);
                Canvas.SetTop (SpectrumCursor, (1 - _v) * sh - 6);
            }

            // Hue strip cursor
            double hh = HueGrid.ActualHeight;
            if (hh > 0)
                Canvas.SetTop(HueCursor, _h / 360 * hh - 2);

            if (!skipHex)       TbHex.Text = $"{r:X2}{g:X2}{b:X2}";
            if (!skipSliders)   { SlR.Value = r; SlG.Value = g; SlB.Value = b; }
            if (!skipTextboxes) { TbR.Text = r.ToString(); TbG.Text = g.ToString(); TbB.Text = b.ToString(); }
        }
        finally { _updating = false; }
    }

    // ── HSV ↔ RGB ────────────────────────────────────────────────────

    private static void HsvToRgb(double h, double s, double v,
        out byte r, out byte g, out byte b)
    {
        if (s <= 0) { r = g = b = (byte)Math.Round(v * 255); return; }
        h = ((h % 360) + 360) % 360;
        double sec = h / 60;
        int    i   = (int)sec;
        double f   = sec - i;
        double p   = v * (1 - s);
        double q   = v * (1 - s * f);
        double t   = v * (1 - s * (1 - f));
        (double dr, double dg, double db) = (i % 6) switch
        {
            0 => (v, t, p),
            1 => (q, v, p),
            2 => (p, v, t),
            3 => (p, q, v),
            4 => (t, p, v),
            _ => (v, p, q),
        };
        r = (byte)Math.Round(dr * 255);
        g = (byte)Math.Round(dg * 255);
        b = (byte)Math.Round(db * 255);
    }

    private static void RgbToHsv(byte rB, byte gB, byte bB,
        out double h, out double s, out double v)
    {
        double r = rB / 255.0, g = gB / 255.0, b = bB / 255.0;
        double max = Math.Max(r, Math.Max(g, b));
        double min = Math.Min(r, Math.Min(g, b));
        double d   = max - min;
        v = max;
        s = max == 0 ? 0 : d / max;
        if (d == 0) { h = 0; return; }
        if      (max == r) h = 60 * ((g - b) / d % 6);
        else if (max == g) h = 60 * ((b - r) / d + 2);
        else               h = 60 * ((r - g) / d + 4);
        if (h < 0) h += 360;
    }

    // ── OK / Cancel ──────────────────────────────────────────────────

    private void OnOk(object sender, RoutedEventArgs e)
    {
        HsvToRgb(_h, _s, _v, out byte r, out byte g, out byte b);
        SelectedColor = new RgbColor(r, g, b);
        DialogResult  = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;
}
