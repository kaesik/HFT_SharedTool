using System.Windows;

namespace HFT_SharedTool;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        var mode = e.Args.Length > 0 ? e.Args[0].ToLowerInvariant() : "standalone";

        var mainWindow = new MainWindow(mode);
        mainWindow.Show();
    }
}