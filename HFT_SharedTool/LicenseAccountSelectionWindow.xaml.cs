using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace HFT_SharedTool;

public partial class LicenseAccountSelectionWindow {
    public LicenseAccountSelectionWindow(string preselectedLicenseId = null) {
        InitializeComponent();
        LoadAccounts(preselectedLicenseId);
    }

    public string SelectedLicenseId { get; private set; }

    private void LoadAccounts(string preselectedLicenseId) {
        List<SharedLicenseInfo> infos;
        try {
            infos = SharedLicenseFileService.LoadAll(SharedConstants.LicenseFilePath);
        }
        catch (Exception) {
            infos = [];
        }

        var infoMap = infos.ToDictionary(
            i => i.LicenseId,
            StringComparer.OrdinalIgnoreCase);

        var items = new List<LicenseDisplayItem>();

        foreach (var licenseId in SharedConstants.DefaultLicenseIds) {
            string loggedUser = null;
            var userDisplay = "WOLNE";

            if (infoMap.TryGetValue(licenseId, out var info) && info.Logins.Count > 0) {
                loggedUser = string.Join(", ", info.Logins.Select(l => l.User));
                userDisplay = loggedUser;
            }

            items.Add(new LicenseDisplayItem {
                LicenseId = licenseId,
                LoggedUser = loggedUser,
                DisplayText = $"{licenseId} ({userDisplay})"
            });
        }

        AccountsListBox.ItemsSource = items;

        if (!string.IsNullOrWhiteSpace(preselectedLicenseId)) {
            var selectedItem = items.FirstOrDefault(item =>
                string.Equals(item.LicenseId, preselectedLicenseId,
                    StringComparison.OrdinalIgnoreCase));

            if (selectedItem != null) {
                AccountsListBox.SelectedItem = selectedItem;
                return;
            }
        }

        if (items.Count > 0)
            AccountsListBox.SelectedIndex = 0;
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

    private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) {
        DragMove();
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e) {
        Close();
    }

    private sealed class LicenseDisplayItem {
        public string LicenseId { get; set; }
        public string LoggedUser { get; set; }
        public string DisplayText { get; set; }
    }
}