using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using PPEditer.Models;

namespace PPEditer.Dialogs;

public partial class ColorPickerPopup : Window
{
    private static readonly string[] PaletteHex =
    [
        "000000", "404040", "808080", "BFBFBF", "FFFFFF", "FF0000",
        "FF8000", "FFFF00", "00FF00", "00FFFF", "0000FF", "8000FF",
        "FF00FF", "800000", "804000", "808000", "008000", "008080",
        "000080", "400080", "800040", "BDD7EE", "2E74B5", "FF9999",
    ];

    public RgbColor SelectedColor { get; private set; }

    public ColorPickerPopup(RgbColor initial)
    {
        InitializeComponent();
        SelectedColor = initial;
        TbHex.Text    = initial.ToHex();

        foreach (var hex in PaletteHex)
        {
            var c = RgbColor.FromHex(hex);
            var btn = new Button
            {
                Width           = 22,
                Height          = 22,
                Background      = new SolidColorBrush(Color.FromRgb(c.R, c.G, c.B)),
                BorderThickness = new Thickness(1),
                BorderBrush     = Brushes.Gray,
                Tag             = hex,
                Margin          = new Thickness(1),
                ToolTip         = $"#{hex}",
            };
            string h = hex;
            btn.Click += (_, _) => TbHex.Text = h;
            ColorGrid.Children.Add(btn);
        }
    }

    private void OnOk(object sender, RoutedEventArgs e)
    {
        SelectedColor = RgbColor.FromHex(TbHex.Text);
        DialogResult  = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;
}
