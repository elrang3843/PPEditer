using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using PPEditer.Models;
using PPEditer.Services;

namespace PPEditer.Dialogs;

public partial class CharPropertiesDialog : Window
{
    private RgbColor? _foreColor;
    private RgbColor? _backColor;
    private RgbColor? _ulColor;

    public CharStyle Result { get; private set; } = new();

    public CharPropertiesDialog(CharStyle initial)
    {
        InitializeComponent();

        // Populate font combo
        CbFont.ItemsSource = FontService.GetAllFonts().Select(f => f.Name).ToList();
        CbFont.Text = initial.FontFamily ?? string.Empty;

        // Size & spacing
        TbSize.Text    = initial.FontSizePt.HasValue
            ? initial.FontSizePt.Value.ToString("F1", CultureInfo.InvariantCulture)
            : string.Empty;
        TbSpacing.Text = initial.SpacingPt100.ToString(CultureInfo.InvariantCulture);

        // Colors
        _foreColor = initial.ForeColor;
        _backColor = initial.HighlightColor;
        _ulColor   = initial.UnderlineColor;
        UpdateColorButton(BtnForeColor, _foreColor);
        UpdateColorButton(BtnBackColor, _backColor);
        UpdateColorButton(BtnUlColor,   _ulColor);

        // Style checkboxes (null → false)
        ChkBold.IsChecked      = initial.Bold      ?? false;
        ChkItalic.IsChecked    = initial.Italic     ?? false;
        ChkOutline.IsChecked   = initial.HasOutline ?? false;
        ChkUnderline.IsChecked = initial.HasUnderline ?? false;
        ChkStrike.IsChecked    = initial.HasStrike  ?? false;
        ChkOverline.IsChecked  = initial.HasOverline ?? false;

        // Script radio buttons
        RbNone.IsChecked  = initial.Script == ScriptKind.None;
        RbSuper.IsChecked = initial.Script == ScriptKind.Superscript;
        RbSub.IsChecked   = initial.Script == ScriptKind.Subscript;
    }

    // ── Color button helpers ──────────────────────────────────────────

    private static void UpdateColorButton(Button btn, RgbColor? color)
    {
        if (color.HasValue)
        {
            btn.Background = new SolidColorBrush(
                Color.FromRgb(color.Value.R, color.Value.G, color.Value.B));
            btn.Content = $"#{color.Value.ToHex()}";
        }
        else
        {
            btn.Background = SystemColors.ControlBrush;
            btn.Content    = Res("Dlg_Char_NoColor", "(없음)");
        }
    }

    private RgbColor? PickColor(RgbColor? current, Button anchor)
    {
        var initial = current ?? new RgbColor(0, 0, 0);
        var screen  = anchor.PointToScreen(new Point(0, anchor.ActualHeight));
        var picker  = new ColorPickerPopup(initial)
        {
            Owner = this,
            Left  = screen.X,
            Top   = screen.Y,
        };
        return picker.ShowDialog() == true ? picker.SelectedColor : current;
    }

    // ── Color pick event handlers ─────────────────────────────────────

    private void OnPickFore(object sender, RoutedEventArgs e)
    {
        _foreColor = PickColor(_foreColor, BtnForeColor);
        UpdateColorButton(BtnForeColor, _foreColor);
    }

    private void OnPickBack(object sender, RoutedEventArgs e)
    {
        _backColor = PickColor(_backColor, BtnBackColor);
        UpdateColorButton(BtnBackColor, _backColor);
    }

    private void OnPickUl(object sender, RoutedEventArgs e)
    {
        _ulColor = PickColor(_ulColor, BtnUlColor);
        UpdateColorButton(BtnUlColor, _ulColor);
    }

    // ── OK / Cancel ───────────────────────────────────────────────────

    private void OnOk(object sender, RoutedEventArgs e)
    {
        // Font family
        string? fontFamily = string.IsNullOrWhiteSpace(CbFont.Text) ? null : CbFont.Text.Trim();

        // Font size
        double? sizePt = null;
        if (!string.IsNullOrWhiteSpace(TbSize.Text) &&
            double.TryParse(TbSize.Text.Trim(), NumberStyles.Any,
                            CultureInfo.InvariantCulture, out double sz) && sz > 0)
            sizePt = sz;

        // Spacing
        int spacing = 0;
        if (!string.IsNullOrWhiteSpace(TbSpacing.Text) &&
            int.TryParse(TbSpacing.Text.Trim(), NumberStyles.Any,
                         CultureInfo.InvariantCulture, out int sp))
            spacing = sp;

        // Script
        var script = RbSuper.IsChecked == true ? ScriptKind.Superscript :
                     RbSub.IsChecked   == true ? ScriptKind.Subscript   :
                                                 ScriptKind.None;

        Result = new CharStyle
        {
            FontFamily     = fontFamily,
            FontSizePt     = sizePt,
            Bold           = ChkBold.IsChecked,
            Italic         = ChkItalic.IsChecked,
            HasUnderline   = ChkUnderline.IsChecked,
            HasStrike      = ChkStrike.IsChecked,
            HasOverline    = ChkOverline.IsChecked,
            HasOutline     = ChkOutline.IsChecked,
            Script         = script,
            SpacingPt100   = spacing,
            ForeColor      = _foreColor,
            HighlightColor = _backColor,
            UnderlineColor = _ulColor,
        };

        DialogResult = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;

    private static string Res(string key, string fallback)
        => Application.Current.TryFindResource(key) is string s ? s : fallback;
}
