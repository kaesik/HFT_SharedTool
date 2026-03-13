using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;

namespace HFT_SharedTool;

public partial class LicenseAccountSelectionWindow {
    private const string LicenseFilePath = @"Z:\000_PMJ\Tekla\HFT_SharedTool\HFT_SharingTool_Licences.txt";
    private const string DateFormat = "yyyy-MM-dd HH:mm";

    private static readonly string[] DefaultLicenseIds = [
        "ts1@holdfort.com",
        "ts2@holdfort.com",
        "ts3@holdfort.com",
        "ts4@holdfort.com",
        "ts5@holdfort.com",
        "ts6@holdfort.com",
        "ts7@holdfort.com",
        "ts8@holdfort.com",
        "td1@holdfort.com"
    ];

    public string SelectedLicenseId { get; private set; }

    public LicenseAccountSelectionWindow(string preselectedLicenseId = null) {
        InitializeComponent();
        LoadAccounts(preselectedLicenseId);
    }

    private void LoadAccounts(string preselectedLicenseId) {
        var loginMap = LoadLoginMap(LicenseFilePath);

        var items = new List<LicenseDisplayItem>();

        foreach (var licenseId in DefaultLicenseIds) {
            loginMap.TryGetValue(licenseId, out var loggedUser);

            var userDisplay = string.IsNullOrWhiteSpace(loggedUser) || loggedUser == "-"
                ? "WOLNE"
                : loggedUser;

            items.Add(new LicenseDisplayItem {
                LicenseId = licenseId,
                LoggedUser = loggedUser,
                DisplayText = $"{licenseId} ({userDisplay})"
            });
        }

        AccountsListBox.ItemsSource = items;

        if (!string.IsNullOrWhiteSpace(preselectedLicenseId)) {
            var selectedItem = items.FirstOrDefault(item =>
                string.Equals(item.LicenseId, preselectedLicenseId, StringComparison.OrdinalIgnoreCase));

            if (selectedItem != null) {
                AccountsListBox.SelectedItem = selectedItem;
                return;
            }
        }

        if (items.Count > 0) {
            AccountsListBox.SelectedIndex = 0;
        }
    }

    private void Ok_Click(object sender, RoutedEventArgs e) {
        if (AccountsListBox.SelectedItem is not LicenseDisplayItem selectedItem) {
            MessageBox.Show(
                "Wybierz konto.",
                "Brak wyboru",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        SelectedLicenseId = selectedItem.LicenseId;
        DialogResult = true;
        Close();
    }

    private static Dictionary<string, string> LoadLoginMap(string filePath) {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var licenseId in DefaultLicenseIds) {
            result[licenseId] = null;
        }

        if (!File.Exists(filePath)) {
            return result;
        }

        var lines = File.ReadAllLines(filePath);
        string currentLicenseId = null;

        foreach (var raw in lines) {
            var line = raw.Trim();

            if (string.IsNullOrWhiteSpace(line)) {
                continue;
            }

            if (DefaultLicenseIds.Contains(line, StringComparer.OrdinalIgnoreCase)) {
                currentLicenseId = line;
                continue;
            }

            if (currentLicenseId == null) {
                continue;
            }

            var parts = line.Split([" - "], StringSplitOptions.None);

            if (parts.Length != 3) {
                continue;
            }

            var action = parts[0].Trim();
            var user = parts[1].Trim();
            var dateText = parts[2].Trim();

            if (!string.Equals(action, "LOG IN", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            if (user == "-" || string.IsNullOrWhiteSpace(user)) {
                result[currentLicenseId] = null;
                continue;
            }

            if (!DateTime.TryParseExact(
                    dateText,
                    DateFormat,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out _)) {
                continue;
            }

            result[currentLicenseId] = user;
        }

        return result;
    }

    private sealed class LicenseDisplayItem {
        public string LicenseId { get; set; }
        public string LoggedUser { get; set; }
        public string DisplayText { get; set; }
    }
}