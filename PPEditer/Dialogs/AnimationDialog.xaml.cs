using System.Windows;
using System.Windows.Controls;
using PPEditer.Models;

namespace PPEditer.Dialogs;

public partial class AnimationDialog : Window
{
    public AnimationKind SelectedKind { get; private set; }

    public AnimationDialog(AnimationKind current)
    {
        InitializeComponent();
        foreach (ListBoxItem item in LbAnimations.Items)
        {
            if (item.Tag is string tag && tag == current.ToString())
            {
                LbAnimations.SelectedItem = item;
                break;
            }
        }
        if (LbAnimations.SelectedItem is null)
            LbAnimations.SelectedIndex = 0;
    }

    private void OnOk(object sender, RoutedEventArgs e)
    {
        SelectedKind = LbAnimations.SelectedItem is ListBoxItem li &&
                       li.Tag is string tag &&
                       Enum.TryParse<AnimationKind>(tag, out var k)
                       ? k : AnimationKind.None;
        DialogResult = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;
}
