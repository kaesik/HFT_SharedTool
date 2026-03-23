using System;
using System.IO;
using System.Windows;

namespace HFT_SharedTool;

internal static class ThemeService {
    private static readonly string SettingsPath =
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "HFT_SharedTool",
            "theme.txt");

    public static bool IsDark { get; private set; } = true;

    public static void Initialize() {
        if (File.Exists(SettingsPath))
            IsDark = File.ReadAllText(SettingsPath).Trim() != "Light";
        Apply();
    }

    public static void Toggle() {
        IsDark = !IsDark;
        Save();
        Apply();
    }

    private static void Apply() {
        var uri = new Uri(
            IsDark ? "Themes/DarkTheme.xaml" : "Themes/LightTheme.xaml",
            UriKind.Relative);
        Application.Current.Resources.MergedDictionaries[0] =
            new ResourceDictionary { Source = uri };
    }

    private static void Save() {
        try {
            var dir = Path.GetDirectoryName(SettingsPath);
            if (!string.IsNullOrEmpty(dir))
                Directory.CreateDirectory(dir);
            File.WriteAllText(SettingsPath, IsDark ? "Dark" : "Light");
        }
        catch { }
    }
}
