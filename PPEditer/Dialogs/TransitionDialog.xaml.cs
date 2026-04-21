using System.Windows;
using System.Windows.Controls;
using PPEditer.Models;

namespace PPEditer.Dialogs;

public partial class TransitionDialog : Window
{
    public TransitionKind SelectedKind { get; private set; }
    public bool           ApplyToAll   { get; private set; }

    public TransitionDialog(TransitionKind current)
    {
        InitializeComponent();
        foreach (ListBoxItem item in LbTransitions.Items)
        {
            if (item.Tag is string tag && tag == current.ToString())
            {
                LbTransitions.SelectedItem = item;
                break;
            }
        }
        if (LbTransitions.SelectedItem is null)
            LbTransitions.SelectedIndex = 0;
    }

    private void OnSelectionChanged(object sender, SelectionChangedEventArgs e) { }

    private void OnOk(object sender, RoutedEventArgs e)
    {
        SelectedKind = LbTransitions.SelectedItem is ListBoxItem li &&
                       li.Tag is string tag &&
                       Enum.TryParse<TransitionKind>(tag, out var k)
                       ? k : TransitionKind.None;
        ApplyToAll   = ChkAllSlides.IsChecked == true;
        DialogResult = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;
}
