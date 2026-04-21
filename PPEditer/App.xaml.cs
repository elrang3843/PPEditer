using System.IO;
using System.Windows;
using PPEditer.Services;

namespace PPEditer;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // Apply persisted theme and language before the window appears
        var settings = AppSettings.Current;
        AppTheme.Apply(settings.Theme);
        AppLanguage.Apply(settings.Language);

        var mainWindow = new MainWindow();
        MainWindow = mainWindow;
        mainWindow.Show();

        if (e.Args.Length > 0 && File.Exists(e.Args[0]))
            mainWindow.OpenFile(e.Args[0]);
    }
}
