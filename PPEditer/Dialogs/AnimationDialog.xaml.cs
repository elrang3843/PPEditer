using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using PPEditer.Models;

namespace PPEditer.Dialogs;

public partial class AnimationDialog : Window
{
    public AnimationKind SelectedKind    { get; private set; }
    public double        DurationSeconds { get; private set; }
    public int           RepeatCount     { get; private set; }
    public bool          AutoPlay        { get; private set; }

    public AnimationDialog(ShapeAnimation current)
    {
        InitializeComponent();
        foreach (ListBoxItem item in LbAnimations.Items)
        {
            if (item.Tag is string tag && tag == current.Kind.ToString())
            {
                LbAnimations.SelectedItem = item;
                break;
            }
        }
        if (LbAnimations.SelectedItem is null)
            LbAnimations.SelectedIndex = 0;

        SetDur(current.DurationMs / 1000.0);
        SetRep(current.RepeatCount);
        ChkAutoPlay.IsChecked = current.AutoPlay;
    }

    private void OnDurUp  (object s, RoutedEventArgs e) => SetDur(GetDur() + 0.1);
    private void OnDurDown(object s, RoutedEventArgs e) => SetDur(GetDur() - 0.1);
    private void OnRepUp  (object s, RoutedEventArgs e) => SetRep(GetRep() + 1);
    private void OnRepDown(object s, RoutedEventArgs e) => SetRep(GetRep() - 1);

    private double GetDur() =>
        double.TryParse(TbDuration.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? v : 0.5;

    private int GetRep() =>
        int.TryParse(TbRepeat.Text, out var v) ? v : 1;

    private void SetDur(double v)
    {
        v = Math.Round(Math.Max(0.1, v), 1);
        TbDuration.Text = v.ToString("0.0", CultureInfo.InvariantCulture);
    }

    private void SetRep(int v)
    {
        TbRepeat.Text = Math.Max(0, v).ToString();
    }

    private void OnOk(object sender, RoutedEventArgs e)
    {
        SelectedKind    = LbAnimations.SelectedItem is ListBoxItem li &&
                          li.Tag is string tag &&
                          Enum.TryParse<AnimationKind>(tag, out var k)
                          ? k : AnimationKind.None;
        DurationSeconds = GetDur();
        RepeatCount     = GetRep();
        AutoPlay        = ChkAutoPlay.IsChecked == true;
        DialogResult    = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;
}
