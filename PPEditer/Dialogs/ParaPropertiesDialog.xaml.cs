using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using PPEditer.Models;

namespace PPEditer.Dialogs;

public partial class ParaPropertiesDialog : Window
{
    public ParagraphStyle Result { get; private set; } = new();

    public ParaPropertiesDialog(ParagraphStyle initial)
    {
        InitializeComponent();

        PopulateHorzAlign();
        PopulateVertAnchor();

        CbHorzAlign.SelectedIndex  = (int)initial.HorzAlign;
        CbVertAnchor.SelectedIndex = (int)initial.VertAnchor;

        TbIndent.Text      = initial.MarginLeftCm.ToString("F2", CultureInfo.InvariantCulture);
        TbHanging.Text     = (-initial.TextIndentCm).ToString("F2", CultureInfo.InvariantCulture);
        TbLineSpacing.Text = initial.LineSpacePct.ToString("F0", CultureInfo.InvariantCulture);
        TbSpaceBefore.Text = initial.SpaceBeforePt.ToString("F1", CultureInfo.InvariantCulture);
        TbSpaceAfter.Text  = initial.SpaceAfterPt.ToString("F1", CultureInfo.InvariantCulture);
    }

    private void PopulateHorzAlign()
    {
        CbHorzAlign.Items.Add(Res("Dlg_Para_AlignLeft",    "왼쪽"));
        CbHorzAlign.Items.Add(Res("Dlg_Para_AlignCenter",  "가운데"));
        CbHorzAlign.Items.Add(Res("Dlg_Para_AlignRight",   "오른쪽"));
        CbHorzAlign.Items.Add(Res("Dlg_Para_AlignJustify", "균형"));
    }

    private void PopulateVertAnchor()
    {
        CbVertAnchor.Items.Add(Res("Dlg_Para_AnchorTop",    "위"));
        CbVertAnchor.Items.Add(Res("Dlg_Para_AnchorMiddle", "가운데"));
        CbVertAnchor.Items.Add(Res("Dlg_Para_AnchorBottom", "아래"));
    }

    private void OnOk(object sender, RoutedEventArgs e)
    {
        if (!double.TryParse(TbIndent.Text,      NumberStyles.Any, CultureInfo.InvariantCulture, out double indent)      || indent < 0 ||
            !double.TryParse(TbHanging.Text,     NumberStyles.Any, CultureInfo.InvariantCulture, out double hanging)     || hanging < 0 ||
            !double.TryParse(TbLineSpacing.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double lineSpacing) || lineSpacing < 10 ||
            !double.TryParse(TbSpaceBefore.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double spaceBefore) || spaceBefore < 0 ||
            !double.TryParse(TbSpaceAfter.Text,  NumberStyles.Any, CultureInfo.InvariantCulture, out double spaceAfter)  || spaceAfter < 0)
        {
            MessageBox.Show(this,
                Res("Dlg_Para_InvalidInput", "올바른 숫자 값을 입력하세요."),
                Res("Lbl_Error", "오류"),
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        Result = new ParagraphStyle
        {
            HorzAlign     = (HorzAlign) Math.Clamp(CbHorzAlign.SelectedIndex,  0, 3),
            VertAnchor    = (VertAnchor)Math.Clamp(CbVertAnchor.SelectedIndex, 0, 2),
            MarginLeftCm  = indent,
            TextIndentCm  = -hanging,   // hanging stored negative (first-line outdent)
            LineSpacePct  = lineSpacing,
            SpaceBeforePt = spaceBefore,
            SpaceAfterPt  = spaceAfter,
        };
        DialogResult = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;

    private static string Res(string key, string fallback)
        => Application.Current.TryFindResource(key) is string s ? s : fallback;
}
