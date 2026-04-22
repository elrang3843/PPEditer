using System.Windows;
using PPEditer.Export;

namespace PPEditer.Dialogs;

public partial class PrintLayoutDialog : Window
{
    public PrintLayout SelectedLayout { get; private set; } = PrintLayout.SlideOnly;
    public bool        BlackWhite     { get; private set; } = false;

    public PrintLayoutDialog(Window owner)
    {
        InitializeComponent();
        Owner = owner;
    }

    private void OnOk(object sender, RoutedEventArgs e)
    {
        SelectedLayout = RbSlideNotes.IsChecked == true ? PrintLayout.SlideWithNotes
                       : RbHandout3.IsChecked   == true ? PrintLayout.Handout3
                       : PrintLayout.SlideOnly;
        BlackWhite   = ChkBW.IsChecked == true;
        DialogResult = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;
}
