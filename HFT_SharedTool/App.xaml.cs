using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows;

namespace HFT_SharedTool;

public partial class App : Application {
    private static readonly string LocalDir =
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "HFT_SharedTool",
            "bin");

    protected override void OnStartup(StartupEventArgs e) {
        base.OnStartup(e);

        if (IsRunningFromLocalCopy()) {
            RunApp(e.Args);
            return;
        }

        CopyAndRelaunch(e.Args);
        Shutdown();
    }

    private static bool IsRunningFromLocalCopy() {
        var exe = Assembly.GetExecutingAssembly().Location;
        return exe.StartsWith(LocalDir, StringComparison.OrdinalIgnoreCase);
    }

    private static void CopyAndRelaunch(string[] args) {
        try {
            var sourceDir = AppDomain.CurrentDomain.BaseDirectory;
            Directory.CreateDirectory(LocalDir);

            foreach (var file in Directory.GetFiles(sourceDir, "*", SearchOption.TopDirectoryOnly)) {
                var dest = Path.Combine(LocalDir, Path.GetFileName(file));
                File.Copy(file, dest, true);
            }

            var localExe = Path.Combine(LocalDir,
                Path.GetFileName(Assembly.GetExecutingAssembly().Location));

            Process.Start(new ProcessStartInfo {
                FileName = localExe,
                Arguments = args.Length > 0 ? string.Join(" ", args) : "",
                UseShellExecute = true
            });
        }
        catch (Exception ex) {
            MessageBox.Show(
                $"Nie udało się uruchomić lokalnej kopii aplikacji:\n{ex.Message}",
                "HFT Shared Tool",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private static void RunApp(string[] args) {
        ThemeService.Initialize();
        var mode = args.Length > 0 ? args[0].ToLowerInvariant() : "standalone";
        var mainWindow = new MainWindow(mode);
        mainWindow.Show();
    }
}