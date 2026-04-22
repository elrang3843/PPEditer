using System.Windows;
using PPEditer.Services;

namespace PPEditer.Dialogs;

public partial class DisplaySettingsDialog : Window
{
    private readonly IReadOnlyList<MonitorInfo> _monitors;

    public DisplaySettingsDialog()
    {
        InitializeComponent();
        _monitors = ScreenHelper.GetMonitors();
        PopulateCombos();
    }

    private void PopulateCombos()
    {
        string auto = Application.Current.TryFindResource("Opt_AutoDetect") as string ?? "자동 감지";

        ShowMonitorCombo.Items.Add(auto);
        PresenterMonitorCombo.Items.Add(auto);

        foreach (var m in _monitors)
        {
            ShowMonitorCombo.Items.Add(m.DisplayName);
            PresenterMonitorCombo.Items.Add(m.DisplayName);
        }

        var s = AppSettings.Current;
        ShowMonitorCombo.SelectedIndex      = s.ShowMonitorIndex      < 0 ? 0 : s.ShowMonitorIndex + 1;
        PresenterMonitorCombo.SelectedIndex = s.PresenterMonitorIndex < 0 ? 0 : s.PresenterMonitorIndex + 1;
    }

    private void OnOK(object sender, RoutedEventArgs e)
    {
        var s = AppSettings.Current;
        s.ShowMonitorIndex      = ShowMonitorCombo.SelectedIndex - 1;
        s.PresenterMonitorIndex = PresenterMonitorCombo.SelectedIndex - 1;
        s.Save();
        DialogResult = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;
}
