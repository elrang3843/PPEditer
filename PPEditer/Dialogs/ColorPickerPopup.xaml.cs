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
        SelectedColor            = initial;
        TbHex.Text               = initial.ToHex();
        PreviewSwatch.Background = new SolidColorBrush(Color.FromRgb(initial.R, initial.G, initial.B));

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
            // Clicking a palette cell immediately confirms the color and closes the picker.
            // The hex text box + OK button remain available for manual hex entry.
            btn.Click += (_, _) =>
            {
                SelectedColor = RgbColor.FromHex(h);
                TbHex.Text    = h;
                DialogResult  = true;
            };
            ColorGrid.Children.Add(btn);
        }
    }

    private void OnHexTextChanged(object sender, TextChangedEventArgs e)
    {
        if (TbHex.Text.Length == 6)
        {
            try
            {
                var c = RgbColor.FromHex(TbHex.Text);
                PreviewSwatch.Background = new SolidColorBrush(Color.FromRgb(c.R, c.G, c.B));
            }
            catch { }
        }
    }

    private void OnOk(object sender, RoutedEventArgs e)
    {
        SelectedColor = RgbColor.FromHex(TbHex.Text);
        DialogResult  = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;
}
