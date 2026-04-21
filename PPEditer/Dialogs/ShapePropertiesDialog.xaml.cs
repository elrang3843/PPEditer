using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using PPEditer.Models;

namespace PPEditer.Dialogs;

public partial class ShapePropertiesDialog : Window
{
    private const double EmuPerCm = 360000.0;

    private RgbColor _fillColor;
    private RgbColor _outlineColor;
    private bool     _isPicture;

    public ShapeStyle Result { get; private set; } = new();

    public ShapePropertiesDialog(ShapeStyle initial)
    {
        InitializeComponent();
        _isPicture    = initial.IsPicture;
        _fillColor    = initial.FillColor;
        _outlineColor = initial.OutlineColor;

        TbName.Text = initial.Name;
        TbX.Text    = (initial.X  / EmuPerCm).ToString("F2", CultureInfo.InvariantCulture);
        TbY.Text    = (initial.Y  / EmuPerCm).ToString("F2", CultureInfo.InvariantCulture);
        TbCx.Text   = (initial.Cx / EmuPerCm).ToString("F2", CultureInfo.InvariantCulture);
        TbCy.Text   = (initial.Cy / EmuPerCm).ToString("F2", CultureInfo.InvariantCulture);

        PopulateFillCombo();
        CbFillKind.SelectedIndex = (int)initial.FillKind;
        UpdateFillColorButton();
        if (initial.IsPicture)
        {
            CbFillKind.IsEnabled    = false;
            BtnFillColor.IsEnabled  = false;
        }

        PopulateOutlineCombo();
        CbOutlineKind.SelectedIndex = (int)initial.OutlineKind;
        TbLineWidth.Text = initial.OutlineWidthPt.ToString("F1", CultureInfo.InvariantCulture);
        UpdateOutlineColorButton();
        SyncOutlineEnabled();
    }

    // ── Combo population ──────────────────────────────────────────────

    private void PopulateFillCombo()
    {
        CbFillKind.Items.Add(Res("Dlg_FillNone",  "없음"));
        CbFillKind.Items.Add(Res("Dlg_FillSolid", "단색"));
    }

    private void PopulateOutlineCombo()
    {
        CbOutlineKind.Items.Add(Res("Dlg_OutlineNone",       "없음"));
        CbOutlineKind.Items.Add(Res("Dlg_OutlineSolid",      "실선"));
        CbOutlineKind.Items.Add(Res("Dlg_OutlineDash",       "파선 (--)"));
        CbOutlineKind.Items.Add(Res("Dlg_OutlineDot",        "점선 (..)"));
        CbOutlineKind.Items.Add(Res("Dlg_OutlineDashDot",    "파선+점 (-.)"));
        CbOutlineKind.Items.Add(Res("Dlg_OutlineDashDotDot", "파선+점+점 (-..)"));
    }

    // ── Color button helpers ──────────────────────────────────────────

    private void UpdateFillColorButton()
    {
        BtnFillColor.Background = new SolidColorBrush(
            Color.FromRgb(_fillColor.R, _fillColor.G, _fillColor.B));
        BtnFillColor.Content = $"#{_fillColor.ToHex()}";
    }

    private void UpdateOutlineColorButton()
    {
        BtnOutlineColor.Background = new SolidColorBrush(
            Color.FromRgb(_outlineColor.R, _outlineColor.G, _outlineColor.B));
        BtnOutlineColor.Content = $"#{_outlineColor.ToHex()}";
    }

    private void SyncOutlineEnabled()
    {
        bool has = CbOutlineKind.SelectedIndex > 0;
        TbLineWidth.IsEnabled      = has;
        BtnOutlineColor.IsEnabled  = has;
    }

    // ── Event handlers ────────────────────────────────────────────────

    private void OnFillKindChanged(object sender, SelectionChangedEventArgs e)
    {
        if (BtnFillColor is null) return;
        BtnFillColor.IsEnabled = CbFillKind.SelectedIndex == (int)FillKind.Solid && !_isPicture;
    }

    private void OnOutlineKindChanged(object sender, SelectionChangedEventArgs e)
    {
        if (TbLineWidth is null) return;
        SyncOutlineEnabled();
    }

    private void OnPickFillColor(object sender, RoutedEventArgs e)
    {
        var c = PickColor(_fillColor, BtnFillColor);
        if (c.HasValue) { _fillColor = c.Value; UpdateFillColorButton(); }
    }

    private void OnPickOutlineColor(object sender, RoutedEventArgs e)
    {
        var c = PickColor(_outlineColor, BtnOutlineColor);
        if (c.HasValue) { _outlineColor = c.Value; UpdateOutlineColorButton(); }
    }

    private RgbColor? PickColor(RgbColor current, Button anchor)
    {
        var screen = anchor.PointToScreen(new Point(0, anchor.ActualHeight));
        var picker = new ColorPickerPopup(current)
        {
            Owner = this,
            Left  = screen.X,
            Top   = screen.Y,
        };
        return picker.ShowDialog() == true ? picker.SelectedColor : (RgbColor?)null;
    }

    // ── OK / Cancel ───────────────────────────────────────────────────

    private void OnOk(object sender, RoutedEventArgs e)
    {
        if (!double.TryParse(TbX.Text,  NumberStyles.Any, CultureInfo.InvariantCulture, out double x)  || x < 0 ||
            !double.TryParse(TbY.Text,  NumberStyles.Any, CultureInfo.InvariantCulture, out double y)  || y < 0 ||
            !double.TryParse(TbCx.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double cx) || cx <= 0 ||
            !double.TryParse(TbCy.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double cy) || cy <= 0)
        {
            MessageBox.Show(this,
                Res("Dlg_ShapeProps_InvalidInput", "올바른 숫자 값을 입력하세요."),
                Res("Lbl_Error", "오류"),
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!double.TryParse(TbLineWidth.Text, NumberStyles.Any,
                             CultureInfo.InvariantCulture, out double lw) || lw < 0.25)
            lw = 1.5;

        Result = new ShapeStyle
        {
            Name           = TbName.Text,
            IsPicture      = _isPicture,
            X              = (long)(x  * EmuPerCm),
            Y              = (long)(y  * EmuPerCm),
            Cx             = (long)(cx * EmuPerCm),
            Cy             = (long)(cy * EmuPerCm),
            FillKind       = (FillKind)Math.Clamp(CbFillKind.SelectedIndex,    0, 1),
            FillColor      = _fillColor,
            OutlineKind    = (OutlineKind)Math.Clamp(CbOutlineKind.SelectedIndex, 0, 5),
            OutlineColor   = _outlineColor,
            OutlineWidthPt = lw,
        };
        DialogResult = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;

    private static string Res(string key, string fallback)
        => Application.Current.TryFindResource(key) is string s ? s : fallback;
}
