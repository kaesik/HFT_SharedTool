using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Media;
using TSM = Tekla.Structures.Model;

namespace HFT_SharedTool;

public partial class MainWindow {
    private const string LicenseFilePath = @"Z:\000_PMJ\Tekla\HFT_sharing_lic.txt";
    private const string DateFormat = "yyyy-MM-dd HH:mm";

    public MainWindow() {
        var myModel = new TSM.Model();
        if (myModel.GetConnectionStatus()) {
            InitializeComponent();
            ModelDrawingLabel.Content = myModel.GetInfo().ModelName.Replace(".db1", "");
        }
        else {
            MessageBox.Show("Keine Verbindung zu Tekla Structures");
        }
        
        // TryInitModel(out var isModelConnected, out var isSharedModel);
        //
        // InitializeComponent();
        //
        // if (isModelConnected && isSharedModel) {
        //     AutoLoginCurrentUser();
        //     MessageBox.Show("IsSharedModel");
        // }
        // else if (isModelConnected) {
        //     AutoLoginCurrentUser();
        //     MessageBox.Show("IsModelConnected");
        // }
        // else {
        //     MessageBox.Show("Standalone");
        // }
    }

    private void TryInitModel(out bool isConnected, out bool isShared) {
        isConnected = false;
        isShared = false;

        var originalOut = Console.Out;
        var originalErr = Console.Error;

        try {
            Console.SetOut(TextWriter.Null);
            Console.SetError(TextWriter.Null);

            var model = new TSM.Model();
            ModelDrawingLabel.Content = model.GetInfo().ModelName.Replace(".db1", "");
            if (!model.GetConnectionStatus())
                return;

            isConnected = true;

            var info = model.GetInfo();
            if (info != null)
                isShared = info.SharedModel;
        }
        catch {
            isConnected = false;
            isShared = false;
        }
        finally {
            try {
                Console.SetOut(originalOut);
                Console.SetError(originalErr);
            }
            catch {
                // ignored
            }
        }
    }

    private void ClearLog() {
        LogTextBox?.Document.Blocks.Clear();
    }

    private void AppendLog(string text) {
        if (LogTextBox != null) {
            LogTextBox.AppendText(text + Environment.NewLine);
            LogTextBox.ScrollToEnd();
        }
    }

    private void AppendColoredStatus(string text, bool isUsable) {
        if (LogTextBox == null) {
            AppendLog(text + (isUsable ? " [USABLE]" : " [NOT USABLE]"));
            return;
        }

        var paragraph = new System.Windows.Documents.Paragraph(
            new System.Windows.Documents.Run(text)) {
            Foreground = isUsable ? Brushes.Green : Brushes.Red
        };

        LogTextBox.Document.Blocks.Add(paragraph);
        LogTextBox.ScrollToEnd();
    }

    private class SharedLicenseInfo {
        public class LoginEntry {
            public string User { get; set; }
            public DateTime Time { get; set; }
        }

        public string LicenseId { get; set; }
        public List<LoginEntry> Logins { get; } = [];
        public string ReadUser { get; set; }
        public DateTime? ReadTime { get; set; }
        public string WriteUser { get; set; }
        public DateTime? WriteTime { get; set; }
        public DateTime? NextUsable { get; set; }

        public void AddOrUpdateLogin(string user, DateTime time) {
            if (string.IsNullOrEmpty(user))
                return;

            foreach (var t in Logins) {
                if (string.Equals(t.User, user, StringComparison.OrdinalIgnoreCase)) {
                    t.Time = time;
                    return;
                }
            }

            Logins.Add(new LoginEntry {
                User = user,
                Time = time
            });
        }

        public void RemoveLogin(string user) {
            if (string.IsNullOrEmpty(user))
                return;

            Logins.RemoveAll(l => string.Equals(l.User, user, StringComparison.OrdinalIgnoreCase));
        }
    }

    private static class SharedLicenseFileService {
        public static List<SharedLicenseInfo> LoadAll(string filePath) {
            if (!File.Exists(filePath))
                return [];

            var lines = File.ReadAllLines(filePath);
            var result = new List<SharedLicenseInfo>();
            SharedLicenseInfo current = null;

            foreach (var raw in lines) {
                var line = raw.Trim();
                if (line.Length == 0)
                    continue;

                var parts = line.Split([" - "], StringSplitOptions.None);

                if (parts.Length == 1 && line.Contains("@")) {
                    current = new SharedLicenseInfo {
                        LicenseId = line
                    };
                    result.Add(current);
                    continue;
                }

                if (current == null)
                    continue;

                switch (parts.Length) {
                    case 3: {
                        var action = parts[0].Trim();
                        var user = parts[1].Trim();
                        var dateText = parts[2].Trim();
                        if (!DateTime.TryParseExact(dateText, DateFormat, CultureInfo.InvariantCulture,
                                DateTimeStyles.None, out var dt))
                            break;

                        if (string.Equals(action, "LOG IN", StringComparison.OrdinalIgnoreCase)) {
                            current.AddOrUpdateLogin(user, dt);
                        }
                        else if (string.Equals(action, "READ IN", StringComparison.OrdinalIgnoreCase)) {
                            current.ReadUser = user;
                            current.ReadTime = dt;
                        }
                        else if (string.Equals(action, "WRITE OUT", StringComparison.OrdinalIgnoreCase)) {
                            current.WriteUser = user;
                            current.WriteTime = dt;
                        }

                        break;
                    }
                    case 2: {
                        var action = parts[0].Trim();
                        var dateText = parts[1].Trim();
                        if (!DateTime.TryParseExact(dateText, DateFormat, CultureInfo.InvariantCulture,
                                DateTimeStyles.None, out var dt))
                            break;

                        if (string.Equals(action, "NEXT USABLE", StringComparison.OrdinalIgnoreCase)) {
                            current.NextUsable = dt;
                        }

                        break;
                    }
                }
            }

            return result;
        }

        public static SharedLicenseInfo LoadOrCreate(string filePath, string licenseId) {
            var infos = LoadAll(filePath);
            foreach (var info in infos) {
                if (string.Equals(info.LicenseId, licenseId, StringComparison.OrdinalIgnoreCase))
                    return info;
            }

            return new SharedLicenseInfo {
                LicenseId = licenseId
            };
        }

        public static void Save(string filePath, SharedLicenseInfo info) {
            var infos = LoadAll(filePath);
            var updated = false;

            for (var i = 0; i < infos.Count; i++) {
                if (string.Equals(infos[i].LicenseId, info.LicenseId, StringComparison.OrdinalIgnoreCase)) {
                    infos[i] = info;
                    updated = true;
                    break;
                }
            }

            if (!updated)
                infos.Add(info);

            var dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            using var writer = new StreamWriter(filePath, false);
            foreach (var lic in infos) {
                if (string.IsNullOrEmpty(lic.LicenseId))
                    continue;

                writer.WriteLine(lic.LicenseId);

                string loginUser;
                string loginDate;
                if (lic.Logins.Count > 0) {
                    var names = new string[lic.Logins.Count];
                    for (var i = 0; i < lic.Logins.Count; i++)
                        names[i] = lic.Logins[i].User;

                    loginUser = string.Join(", ", names);
                    var lastLogin = lic.Logins[lic.Logins.Count - 1];
                    loginDate = lastLogin.Time.ToString(DateFormat);
                }
                else {
                    loginUser = "-";
                    loginDate = "-";
                }

                writer.WriteLine("LOG IN - {0} - {1}", loginUser, loginDate);

                var readUser = string.IsNullOrEmpty(lic.ReadUser) ? "-" : lic.ReadUser;
                var readDate = lic.ReadTime.HasValue ? lic.ReadTime.Value.ToString(DateFormat) : "-";
                writer.WriteLine("READ IN - {0} - {1}", readUser, readDate);

                var writeUser = string.IsNullOrEmpty(lic.WriteUser) ? "-" : lic.WriteUser;
                var writeDate = lic.WriteTime.HasValue ? lic.WriteTime.Value.ToString(DateFormat) : "-";
                writer.WriteLine("WRITE OUT - {0} - {1}", writeUser, writeDate);

                var nextUsableDate = lic.NextUsable.HasValue ? lic.NextUsable.Value.ToString(DateFormat) : "-";
                writer.WriteLine("NEXT USABLE - {0}", nextUsableDate);

                writer.WriteLine();
            }
        }
    }

    private static class SharedLicenseManager {
        private static readonly TimeSpan HoldDuration = TimeSpan.FromHours(4);

        public static void Login(SharedLicenseInfo info, string userName, DateTime loginTime) {
            info.AddOrUpdateLogin(userName, loginTime);
        }

        public static void Logout(SharedLicenseInfo info, string userName) {
            info.RemoveLogin(userName);
        }

        public static void ReadIn(SharedLicenseInfo info, string userName, DateTime readTime) {
            info.ReadUser = userName;
            info.ReadTime = readTime;
            info.NextUsable = readTime + HoldDuration;
        }

        public static void WriteOut(SharedLicenseInfo info, string userName, DateTime writeTime) {
            info.WriteUser = userName;
            info.WriteTime = writeTime;
            info.NextUsable = writeTime + HoldDuration;
        }

        public static string FormatStatus(SharedLicenseInfo info, DateTime now, out bool isUsableNow) {
            isUsableNow = true;

            string activity;
            string activityUser;
            DateTime activityTime;

            var hasRead = info.ReadTime.HasValue;
            var hasWrite = info.WriteTime.HasValue;

            if (hasRead && (!hasWrite || info.ReadTime.Value > info.WriteTime.Value)) {
                activity = "READ IN";
                activityUser = $"({info.ReadUser}) ";
                activityTime = info.ReadTime.Value;
            }
            else if (hasWrite) {
                activity = "WRITE OUT";
                activityUser = $"({info.WriteUser}) ";
                activityTime = info.WriteTime.Value;
            }
            else {
                activity = "BRAK DANYCH";
                activityUser = "";
                activityTime = now;
            }

            string user;
            if (info.Logins.Count > 0) {
                var names = new string[info.Logins.Count];
                for (var i = 0; i < info.Logins.Count; i++)
                    names[i] = info.Logins[i].User;
                user = string.Join(", ", names);
            }
            else {
                user = "WOLNE";
            }

            if (info.NextUsable.HasValue)
                isUsableNow = now >= info.NextUsable.Value;

            var nextUsableText = info.NextUsable.HasValue
                ? info.NextUsable.Value.ToString(DateFormat)
                : "-";

            return
                $"{info.LicenseId} ({user}) - {activity} {activityUser}{activityTime:yyyy-MM-dd HH:mm} - USABLE {nextUsableText}";
        }
    }

    #region Event Handlers

    private void BtnStandaloneCheck_Click(object sender, RoutedEventArgs e) {
        ClearLog();

        var infos = SharedLicenseFileService.LoadAll(LicenseFilePath);
        if (infos == null || infos.Count == 0) {
            AppendLog("Brak danych licencji: " + LicenseFilePath);
            return;
        }

        foreach (var info in infos) {
            var line = SharedLicenseManager.FormatStatus(info, DateTime.Now, out var isUsableNow);
            AppendColoredStatus(line, isUsableNow);
        }
    }

    private void BtnLogin_Click(object sender, RoutedEventArgs e) {
        ClearLog();

        if (!TryGetLicenseFromLog(out var licenseId, out var loginTime)) {
            AppendLog("Nie znaleziono wpisu UserInfo w logu Tekli.");
            return;
        }

        var userName = Environment.UserName;
        var info = SharedLicenseFileService.LoadOrCreate(LicenseFilePath, licenseId);
        SharedLicenseManager.Login(info, userName, loginTime);
        SharedLicenseFileService.Save(LicenseFilePath, info);
        AppendLog($"LOG IN - {userName} - {loginTime.ToString(DateFormat)}");
    }

    private void BtnReadIn_Click(object sender, RoutedEventArgs e) {
        ClearLog();

        if (!TryGetLicenseFromLog(out var licenseId, out _)) {
            AppendLog("Nie znaleziono wpisu UserInfo w logu Tekli.");
            return;
        }

        var info = SharedLicenseFileService.LoadOrCreate(LicenseFilePath, licenseId);
        var userName = Environment.UserName;
        var now = DateTime.Now;
        SharedLicenseManager.ReadIn(info, userName, now);
        SharedLicenseFileService.Save(LicenseFilePath, info);
        AppendLog($"READ IN - {userName} - {now.ToString(DateFormat)}");
    }

    private void BtnWriteOut_Click(object sender, RoutedEventArgs e) {
        ClearLog();

        if (!TryGetLicenseFromLog(out var licenseId, out _)) {
            AppendLog("Nie znaleziono wpisu UserInfo w logu Tekli.");
            return;
        }

        var info = SharedLicenseFileService.LoadOrCreate(LicenseFilePath, licenseId);
        var userName = Environment.UserName;
        var now = DateTime.Now;
        SharedLicenseManager.WriteOut(info, userName, now);
        SharedLicenseFileService.Save(LicenseFilePath, info);
        AppendLog($"WRITE OUT - {userName} - {now.ToString(DateFormat)}");
    }

    private void BtnCheck_Click(object sender, RoutedEventArgs e) {
        ClearLog();
        BtnStandaloneCheck_Click(sender, e);
    }

    #endregion

    private static void AutoLoginCurrentUser() {
        if (!TryGetLicenseFromLog(out var licenseId, out var loginTime))
            return;

        var userName = Environment.UserName;
        var info = SharedLicenseFileService.LoadOrCreate(LicenseFilePath, licenseId);
        SharedLicenseManager.Login(info, userName, loginTime);
        SharedLicenseFileService.Save(LicenseFilePath, info);
    }

    private static void RemoveCurrentUserLogin() {
        if (!TryGetLicenseFromLog(out var licenseId, out _))
            return;

        var info = SharedLicenseFileService.LoadOrCreate(LicenseFilePath, licenseId);
        var userName = Environment.UserName;
        SharedLicenseManager.Logout(info, userName);
        SharedLicenseFileService.Save(LicenseFilePath, info);
    }

    protected override void OnClosed(EventArgs e) {
        RemoveCurrentUserLogin();
        base.OnClosed(e);
    }

    private static bool TryGetLicenseFromLog(out string licenseId, out DateTime loginTime) {
        licenseId = null;
        loginTime = DateTime.MinValue;

        var userName = Environment.UserName;
        var logPath = Path.Combine(@"C:\TeklaStructuresModels", "TeklaStructures_" + userName + ".log");
        if (!File.Exists(logPath))
            return false;

        var tempPath = Path.GetTempFileName();

        try {
            using var fs = new FileStream(logPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var dest = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None);
            fs.CopyTo(dest);
        }
        catch {
            return false;
        }

        string[] lines;
        try {
            lines = File.ReadAllLines(tempPath);
        }
        finally {
            try {
                File.Delete(tempPath);
            }
            catch {
                // ignored
            }
        }

        for (var i = lines.Length - 1; i >= 0; i--) {
            var line = lines[i];
            if (line.IndexOf("UserInfo", StringComparison.OrdinalIgnoreCase) < 0)
                continue;

            var parts = line.Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 3)
                continue;

            string lic = null;
            foreach (var p in parts) {
                if (p.Contains("@") && !p.StartsWith("http", StringComparison.OrdinalIgnoreCase)) {
                    lic = p.Trim();
                    break;
                }
            }

            if (string.IsNullOrEmpty(lic))
                continue;

            licenseId = lic;
            loginTime = DateTime.Now;
            return true;
        }

        return false;
    }
}