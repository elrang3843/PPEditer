using System.Globalization;
using System.Windows;

namespace PPEditer.Dialogs;

public partial class ShapePropertiesDialog : Window
{
    private const double EmuPerCm = 360000.0;

    public long ResultX  { get; private set; }
    public long ResultY  { get; private set; }
    public long ResultCx { get; private set; }
    public long ResultCy { get; private set; }

    public ShapePropertiesDialog(string name, long xEmu, long yEmu, long cxEmu, long cyEmu)
    {
        InitializeComponent();
        TbName.Text = name;
        TbX.Text    = (xEmu  / EmuPerCm).ToString("F2", CultureInfo.InvariantCulture);
        TbY.Text    = (yEmu  / EmuPerCm).ToString("F2", CultureInfo.InvariantCulture);
        TbCx.Text   = (cxEmu / EmuPerCm).ToString("F2", CultureInfo.InvariantCulture);
        TbCy.Text   = (cyEmu / EmuPerCm).ToString("F2", CultureInfo.InvariantCulture);
    }

    private void OnOk(object sender, RoutedEventArgs e)
    {
        if (!double.TryParse(TbX.Text,  NumberStyles.Any, CultureInfo.InvariantCulture, out double x)  || x < 0 ||
            !double.TryParse(TbY.Text,  NumberStyles.Any, CultureInfo.InvariantCulture, out double y)  || y < 0 ||
            !double.TryParse(TbCx.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double cx) || cx <= 0 ||
            !double.TryParse(TbCy.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double cy) || cy <= 0)
        {
            string msg = Application.Current.TryFindResource("Dlg_ShapeProps_InvalidInput") as string
                         ?? "올바른 숫자 값을 입력하세요.";
            MessageBox.Show(this, msg, "오류", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        ResultX  = (long)(x  * EmuPerCm);
        ResultY  = (long)(y  * EmuPerCm);
        ResultCx = (long)(cx * EmuPerCm);
        ResultCy = (long)(cy * EmuPerCm);
        DialogResult = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;
}
