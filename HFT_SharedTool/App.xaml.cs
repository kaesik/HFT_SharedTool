using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows;

namespace HFT_SharedTool;

public partial class App : Application {
    private static readonly string BaseLocalDir =
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
        return exe.StartsWith(BaseLocalDir, StringComparison.OrdinalIgnoreCase);
    }

    private static void CopyAndRelaunch(string[] args) {
        try {
            var instanceDir = Path.Combine(BaseLocalDir, Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(instanceDir);

            var sourceDir = AppDomain.CurrentDomain.BaseDirectory;
            foreach (var file in Directory.GetFiles(sourceDir, "*", SearchOption.TopDirectoryOnly))
                File.Copy(file, Path.Combine(instanceDir, Path.GetFileName(file)), true);

            var localExe = Path.Combine(instanceDir,
                Path.GetFileName(Assembly.GetExecutingAssembly().Location));

            Process.Start(new ProcessStartInfo {
                FileName = localExe,
                Arguments = args.Length > 0 ? string.Join(" ", args) : "",
                UseShellExecute = true
            });

            CleanupOldInstances(instanceDir);
        }
        catch (Exception ex) {
            MessageBox.Show(
                $"Nie udało się uruchomić lokalnej kopii aplikacji:\n{ex.Message}",
                "HFT Shared Tool",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private static void CleanupOldInstances(string currentInstanceDir) {
        try {
            if (!Directory.Exists(BaseLocalDir)) return;

            foreach (var dir in Directory.GetDirectories(BaseLocalDir)) {
                if (string.Equals(dir, currentInstanceDir, StringComparison.OrdinalIgnoreCase))
                    continue;
                try {
                    Directory.Delete(dir, true);
                }
                catch {
                    // ignored
                }
            }
        }
        catch {
            // ignored
        }
    }

    private static void RunApp(string[] args) {
        ThemeService.Initialize();
        var mode = args.Length > 0 ? args[0].ToLowerInvariant() : "standalone";
        var mainWindow = new MainWindow(mode);
        try {
            mainWindow.Show();
        }
        catch (InvalidOperationException) {
            Current.Shutdown();
        }
    }
}