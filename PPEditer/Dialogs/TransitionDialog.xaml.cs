using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using PPEditer.Models;

namespace PPEditer.Dialogs;

public partial class TransitionDialog : Window
{
    public TransitionKind SelectedKind    { get; private set; }
    public bool           ApplyToAll      { get; private set; }
    public double         DurationSeconds { get; private set; }

    public TransitionDialog(SlideTransition current)
    {
        InitializeComponent();
        foreach (ListBoxItem item in LbTransitions.Items)
        {
            if (item.Tag is string tag && tag == current.Kind.ToString())
            {
                LbTransitions.SelectedItem = item;
                break;
            }
        }
        if (LbTransitions.SelectedItem is null)
            LbTransitions.SelectedIndex = 0;

        SetDur(current.DurationMs / 1000.0);
    }

    private void OnSelectionChanged(object sender, SelectionChangedEventArgs e) { }

    private void OnDurUp  (object s, RoutedEventArgs e) => SetDur(GetDur() + 0.1);
    private void OnDurDown(object s, RoutedEventArgs e) => SetDur(GetDur() - 0.1);

    private double GetDur() =>
        double.TryParse(TbDuration.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? v : 0.7;

    private void SetDur(double v)
    {
        v = Math.Round(Math.Max(0.1, v), 1);
        TbDuration.Text = v.ToString("0.0", CultureInfo.InvariantCulture);
    }

    private void OnOk(object sender, RoutedEventArgs e)
    {
        SelectedKind = LbTransitions.SelectedItem is ListBoxItem li &&
                       li.Tag is string tag &&
                       Enum.TryParse<TransitionKind>(tag, out var k)
                       ? k : TransitionKind.None;
        ApplyToAll      = ChkAllSlides.IsChecked == true;
        DurationSeconds = GetDur();
        DialogResult    = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;
}
