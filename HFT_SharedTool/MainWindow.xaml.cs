using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using TSM = Tekla.Structures.Model;
using TSMO = Tekla.Structures.Model.Operations;

namespace HFT_SharedTool;

public partial class MainWindow {
    private const string LicenseFilePath = @"Z:\000_PMJ\Tekla\HFT_SharedTool\HFT_SharingTool_Licences.txt";
    private const string DateFormat = "yyyy-MM-dd HH:mm";

    private readonly string _mode;
    private bool _autoLoginDone;
    private bool _logoutDone;
    private TSM.Events _events;

    public MainWindow() : this("standalone") { }

    public MainWindow(string mode) {
        _mode = mode?.ToLowerInvariant() ?? "standalone";

        InitializeComponent();

        if (_mode == "standalone") {
            ModelDrawingLabel.Content = "Tryb odczytu pliku licencji";
            Loaded += (_, _) => BtnCheck_Click();
            return;
        }

        if (!TryWaitForSharedModel(out var model)) {
            ModelDrawingLabel.Content = "Brak połączenia z modelem współdzielonym";
            Close();
            return;
        }

        _events = new TSM.Events();
        _events.ModelUnloading += OnModelUnloading;
        _events.TeklaStructuresExit += OnTeklaStructuresExit;
        _events.Register();

        var modelName = model.GetInfo().ModelName.Replace(".db1", "");
        ModelDrawingLabel.Content = $"Połączono z {modelName}";

        switch (_mode) {
            case "autologin":
                HideWindow();
                Loaded += (_, _) => StartAutoLoginWatcher();
                break;

            case "readin":
                HideWindow();
                Loaded += (_, _) => BtnReadIn_Click();
                break;

            case "writeout":
                HideWindow();
                Loaded += (_, _) => BtnWriteOut_Click();
                break;

            default:
                Loaded += (_, _) => BtnCheck_Click();
                break;
        }
    }

    private void HideWindow() {
        WindowState = WindowState.Minimized;
        ShowInTaskbar = false;
        Visibility = Visibility.Hidden;
    }
    
    private static bool TryWaitForSharedModel(out TSM.Model model, int maxWaitSeconds = 300, int delaySeconds = 5) {
        model = new TSM.Model();
        var timeout = TimeSpan.FromSeconds(maxWaitSeconds);
        var delay = TimeSpan.FromSeconds(delaySeconds);
        var start = DateTime.Now;

        while (DateTime.Now - start < timeout) {
            try {
                if (model.GetConnectionStatus()) {
                    var info = model.GetInfo();
                    if (info.SharedModel)
                        return true;
                }
            }
            catch {
                // ignored
            }

            System.Threading.Thread.Sleep(delay);
        }

        model = null;
        return false;
    }

    #region Logging Helpers
    
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

    #endregion

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

        public static string FormatStatus(
            SharedLicenseInfo info,
            DateTime now,
            string currentUser,
            out bool isUsableForCurrentUser)
        {
            string activity;
            string activityUserDisplay;
            DateTime activityTime;
            string latestActivityUserName = null;

            var hasRead = info.ReadTime.HasValue;
            var hasWrite = info.WriteTime.HasValue;

            if (hasRead && (!hasWrite || info.ReadTime.Value > info.WriteTime.Value)) {
                activity = "READ IN";
                activityUserDisplay = $"({info.ReadUser}) ";
                activityTime = info.ReadTime.Value;
                latestActivityUserName = info.ReadUser;
            }
            else if (hasWrite) {
                activity = "WRITE OUT";
                activityUserDisplay = $"({info.WriteUser}) ";
                activityTime = info.WriteTime.Value;
                latestActivityUserName = info.WriteUser;
            }
            else {
                activity = "BRAK DANYCH";
                activityUserDisplay = "";
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

            var isAfterHold = !info.NextUsable.HasValue || now >= info.NextUsable.Value;

            var isCurrentHolder = !string.IsNullOrEmpty(latestActivityUserName) &&
                                  !string.IsNullOrEmpty(currentUser) &&
                                  string.Equals(latestActivityUserName, currentUser, StringComparison.OrdinalIgnoreCase);

            isUsableForCurrentUser = isAfterHold || isCurrentHolder;

            var nextUsableText = info.NextUsable.HasValue
                ? info.NextUsable.Value.ToString(DateFormat)
                : "-";

            return
                $"{info.LicenseId} ({user}) - {activity} {activityUserDisplay}{activityTime:yyyy-MM-dd HH:mm} - USABLE {nextUsableText}";
        }
    }

    #region Event Handlers

    private void BtnReadIn_Click() {
        try {
            ClearLog();

            if (!TryGetLicenseFromLog(out var licenseId, out _)) return; 
            if (!TryGetModelSharingLogPath(out var modelSharingLogPath)) return;
            
            var info = SharedLicenseFileService.LoadOrCreate(LicenseFilePath, licenseId);
            var userName = Environment.UserName;
            var fromUtc = DateTime.UtcNow.AddMinutes(-5);

            ReadIn();

            _ = WaitAndFinalizeReadInAsync(modelSharingLogPath, info, userName, fromUtc);
        }
        catch (Exception ex) {
            MessageBox.Show($"READ IN: wyjątek: {ex.Message}");
        }
    }

    private void BtnWriteOut_Click() {
        try {
            ClearLog();

            if (!TryGetLicenseFromLog(out var licenseId, out _)) return;
            if (!TryGetModelSharingLogPath(out var modelSharingLogPath)) return;
            
            var info = SharedLicenseFileService.LoadOrCreate(LicenseFilePath, licenseId);
            var userName = Environment.UserName;
            var fromUtc = DateTime.UtcNow.AddMinutes(-5);

            WriteOut();
            
            _ = WaitAndFinalizeWriteOutAsync(modelSharingLogPath, info, userName, fromUtc);
        }
        catch (Exception ex) {
            MessageBox.Show($"WRITE OUT: wyjątek: {ex.Message}");
        }
    }

    private void BtnCheck_Click() {
        ClearLog();

        var infos = SharedLicenseFileService.LoadAll(LicenseFilePath);
        if (infos == null || infos.Count == 0) {
            AppendLog("Brak danych licencji: " + LicenseFilePath);
            return;
        }

        var currentUser = Environment.UserName;

        foreach (var info in infos) {
            var line = SharedLicenseManager.FormatStatus(info, DateTime.Now, currentUser, out var isUsableForCurrentUser);
            AppendColoredStatus(line, isUsableForCurrentUser);
        }
    }
    
    #endregion

    #region Helpers
    private static string SafeReadLogToTemp(string logPath) {
        if (string.IsNullOrEmpty(logPath) || !File.Exists(logPath))
            return null;

        try {
            var tempPath = Path.GetTempFileName();

            using var src = new FileStream(logPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var dst = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None);
            src.CopyTo(dst);

            return tempPath;
        }
        catch {
            return null;
        }
    }

    private async Task WaitAndFinalizeReadInAsync(
        string logPath,
        SharedLicenseInfo info,
        string userName,
        DateTime fromUtc)
    {
        var ok = await WaitForModelSharingConfirmationAsync(
            logPath,
            "Read-in result: OK.",
            fromUtc,
            maxWaitSeconds: 300,
            delaySeconds: 5).ConfigureAwait(false);

        Application.Current.Dispatcher.Invoke(() => {
            if (!ok) {
                if (_mode == "readin") Close();
                return;
            }

            var now = DateTime.Now;
            SharedLicenseManager.ReadIn(info, userName, now);
            SharedLicenseFileService.Save(LicenseFilePath, info);
            AppendLog($"READ IN - {userName} - {now.ToString(DateFormat)}");

            if (_mode == "readin") Close();
        });
    }

    private async Task WaitAndFinalizeWriteOutAsync(
        string logPath,
        SharedLicenseInfo info,
        string userName,
        DateTime fromUtc)
    {
        var ok = await WaitForModelSharingConfirmationAsync(
            logPath,
            "WriteOut OK",
            fromUtc,
            maxWaitSeconds: 300,
            delaySeconds: 5).ConfigureAwait(false);

        Application.Current.Dispatcher.Invoke(() => {
            if (!ok) {
                if (_mode == "writeout") Close();
                return;
            }

            var now = DateTime.Now;
            SharedLicenseManager.WriteOut(info, userName, now);
            SharedLicenseFileService.Save(LicenseFilePath, info);
            AppendLog($"WRITE OUT - {userName} - {now.ToString(DateFormat)}");

            if (_mode == "writeout") Close();
        });
    }

    private static async Task<bool> WaitForModelSharingConfirmationAsync(
        string logPath,
        string marker,
        DateTime fromUtc,
        int maxWaitSeconds = 300,
        int delaySeconds = 5)
    {
        var timeout = TimeSpan.FromSeconds(maxWaitSeconds);
        var delay = TimeSpan.FromSeconds(delaySeconds);
        var start = DateTime.UtcNow;

        if (fromUtc.Kind != DateTimeKind.Utc) fromUtc = fromUtc.ToUniversalTime();

        while (DateTime.UtcNow - start < timeout) {

            try {
                var tempPath = SafeReadLogToTemp(logPath);
                if (tempPath != null) {
                    bool found;
                    try {
                        found = HasConfirmationLine(tempPath, marker, fromUtc);
                    }
                    finally {
                        try {
                            File.Delete(tempPath);
                        }
                        catch {
                            // ignored
                        }
                    }

                    if (found) return true;
                }
            }
            catch {
                // ignored
            }

            await Task.Delay(delay).ConfigureAwait(false);
        }

        return false;
    }
    
    private static bool TryGetModelSharingLogPath(out string logPath) {
        logPath = null;

        try {
            var model = new TSM.Model();
            if (!model.GetConnectionStatus())
                return false;

            var info = model.GetInfo();
            if (!info.SharedModel)
                return false;

            var basePath = info.ModelPath;
            if (string.IsNullOrWhiteSpace(basePath))
                return false;

            basePath = basePath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            logPath = Path.Combine(basePath, "logs", "modelsharing.log");
            return true;
        }
        catch {
            logPath = null;
            return false;
        }
    }
    private static bool HasConfirmationLine(string logPath, string marker, DateTime fromUtc) {
        string[] lines;
        try {
            lines = File.ReadAllLines(logPath);
        }
        catch {
            return false;
        }

        if (fromUtc.Kind != DateTimeKind.Utc) fromUtc = fromUtc.ToUniversalTime();

        for (var i = lines.Length - 1; i >= 0; i--) {
            var line = lines[i];

            if (line.IndexOf(marker, StringComparison.OrdinalIgnoreCase) < 0) continue;

            var startBracket = line.IndexOf('[');
            var endBracket = line.IndexOf(']', startBracket + 1);

            if (startBracket < 0 || endBracket <= startBracket + 1) continue;
            
            var tsText = line.Substring(startBracket + 1, endBracket - startBracket - 1);

            if (!DateTime.TryParse(
                    tsText,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                    out var tsUtc))
            {
                continue;
            }

            var isNewEnough = tsUtc >= fromUtc;

            if (isNewEnough) return true;
        }

        return false;
    }

    #endregion
    
    #region Auto Login/Logout

    private async void StartAutoLoginWatcher() {
        try {
            ClearLog();
            AppendLog("Czekam na wpis UserInfo w logu Tekli...");

            var start = DateTime.Now;
            var maxWait = TimeSpan.FromMinutes(5);

            while (!_autoLoginDone && DateTime.Now - start < maxWait) {
                if (TryGetLicenseFromLog(out var licenseId, out var loginTime)) {
                    var userName = Environment.UserName;
                    var info = SharedLicenseFileService.LoadOrCreate(LicenseFilePath, licenseId);
                    SharedLicenseManager.Login(info, userName, loginTime);
                    SharedLicenseFileService.Save(LicenseFilePath, info);
                    _autoLoginDone = true;
                    AppendLog($"LOG IN - {userName} - {loginTime.ToString(DateFormat)}");
                    break;
                }

                await Task.Delay(TimeSpan.FromSeconds(5));
            }
        }
        catch {
            // ignored
        }
    }

    private void OnModelUnloading() {
        HandleModelClosed();
    }

    private void OnTeklaStructuresExit() {
        HandleModelClosed();
    }

    private void HandleModelClosed() {
        if (_mode == "autologin" && !_logoutDone) {
            RemoveCurrentUserLogin();
            _logoutDone = true;
        }

        Application.Current.Dispatcher.Invoke(Close);
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
        if (_mode == "autologin" && !_logoutDone) {
            RemoveCurrentUserLogin();
            _logoutDone = true;
        }

        if (_events != null) {
            _events.ModelUnloading -= OnModelUnloading;
            _events.TeklaStructuresExit -= OnTeklaStructuresExit;
            _events.UnRegister();
            _events = null;
        }

        base.OnClosed(e);
    }

    private static bool TryGetLicenseFromLog(out string licenseId, out DateTime loginTime) {
        licenseId = null;
        loginTime = DateTime.MinValue;

        var userName = Environment.UserName;
        var logPath = Path.Combine(@"C:\TeklaStructuresModels", "TeklaStructures_" + userName + ".log");
        var tempPath = SafeReadLogToTemp(logPath);
        if (tempPath == null)
            return false;

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

    #endregion

    #region Scripts
    
    private static void ReadIn() {
        TSMO.Operation.RunMacro(@"Z:\000_PMJ\Tekla\HFT_SharedTool\SharedTool\ReadIn.cs");
    }

    private static void WriteOut() {
        TSMO.Operation.RunMacro(@"Z:\000_PMJ\Tekla\HFT_SharedTool\SharedTool\WriteOut.cs");
    }

    #endregion

}