using System.Windows;

namespace PPEditer.Dialogs;

public partial class AboutDialog : Window
{
    public AboutDialog() => InitializeComponent();
    private void OkButton_Click(object sender, RoutedEventArgs e) => Close();
}
