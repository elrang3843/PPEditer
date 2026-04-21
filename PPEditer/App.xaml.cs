using System.Windows;

namespace PPEditer;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // Open file passed as command-line argument
        if (e.Args.Length > 0 && System.IO.File.Exists(e.Args[0]))
        {
            var mainWindow = new MainWindow();
            mainWindow.Show();
            mainWindow.OpenFile(e.Args[0]);
        }
    }
}
