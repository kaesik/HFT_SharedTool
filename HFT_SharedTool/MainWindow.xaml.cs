using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
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
    private int _autoLoginWatcherStarted;
    private bool _autoLoginDone;
    private bool _logoutDone;
    private TSM.Events _events;

    public MainWindow() : this("standalone") { }

    public MainWindow(string mode) {
        DumpTeklaDiagnostics();
        _mode = (mode ?? "standalone").Trim().ToLowerInvariant();

        Dbg($"START MainWindow mode={_mode}");

        AppDomain.CurrentDomain.UnhandledException += (_, e) => {
            Dbg("UnhandledException", e.ExceptionObject as Exception);
        };

        TaskScheduler.UnobservedTaskException += (_, e) => {
            Dbg("UnobservedTaskException", e.Exception);
            e.SetObserved();
        };

        Application.Current.DispatcherUnhandledException += (_, e) => {
            Dbg("DispatcherUnhandledException", e.Exception);
        };

        InitializeComponent();

        if (_mode is "standalone" or "check" or "ckeck") {
            ModelDrawingLabel.Content = "Tryb odczytu pliku licencji";
            Dispatcher.BeginInvoke(new Action(BtnCheck_Click));
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
                Dispatcher.BeginInvoke(new Action(StartAutoLoginWatcher));
                break;

            case "readin":
                HideWindow();
                Dispatcher.BeginInvoke(new Action(BtnReadIn_Click));
                break;

            case "writeout":
                HideWindow();
                Dispatcher.BeginInvoke(new Action(BtnWriteOut_Click));
                break;

            default:
                Dispatcher.BeginInvoke(new Action(BtnCheck_Click));
                break;
        }
    }

    private void HideWindow() {
        WindowState = WindowState.Minimized;
        ShowInTaskbar = false;
        Visibility = Visibility.Hidden;
    }
    
    private static bool TryWaitForSharedModel(out TSM.Model model, int maxWaitSeconds = 300, int delaySeconds = 5) {
        Dbg($"TryWaitForSharedModel: ENTER maxWaitSeconds={maxWaitSeconds} delaySeconds={delaySeconds}");

        model = null;

        var timeout = TimeSpan.FromSeconds(maxWaitSeconds);
        var delay = TimeSpan.FromSeconds(delaySeconds);
        var start = DateTime.Now;

        var attempt = 0;

        while (DateTime.Now - start < timeout) {
            attempt++;

            TSM.Model m = null;
            try {
                m = new TSM.Model();

                bool connected;
                try {
                    connected = m.GetConnectionStatus();
                }
                catch (Exception ex) {
                    Dbg($"TryWaitForSharedModel: attempt={attempt} GetConnectionStatus EX", ex);
                    connected = false;
                }

                if (!connected) {
                    Dbg($"TryWaitForSharedModel: attempt={attempt} connected=FALSE (will retry)");
                }
                else {
                    try {
                        var info = m.GetInfo();
                        Dbg($"TryWaitForSharedModel: attempt={attempt} connected=TRUE sharedModel={info.SharedModel}");

                        if (info.SharedModel) {
                            model = m;
                            Dbg("TryWaitForSharedModel: SUCCESS -> true");
                            return true;
                        }
                    }
                    catch (Exception ex) {
                        Dbg($"TryWaitForSharedModel: attempt={attempt} GetInfo EX", ex);
                    }
                }
            }
            catch (Exception ex) {
                Dbg($"TryWaitForSharedModel: attempt={attempt} new Model EX", ex);
            }
            finally {
                if (model != m) {
                }
            }

            System.Threading.Thread.Sleep(delay);
        }

        Dbg("TryWaitForSharedModel: TIMEOUT -> false");
        model = null;
        return false;
    }

    #region Logging Helpers
    
    private void ClearLog() {
        try {
            if (Application.Current?.Dispatcher == null || Application.Current.Dispatcher.CheckAccess()) {
                LogTextBox?.Document.Blocks.Clear();
                return;
            }

            Application.Current.Dispatcher.BeginInvoke(new Action(() => {
                LogTextBox?.Document.Blocks.Clear();
            }));
        }
        catch {
            // ignored
        }
    }

    private void AppendLog(string text) {
        try {
            if (Application.Current?.Dispatcher == null || Application.Current.Dispatcher.CheckAccess()) {
                if (LogTextBox != null) {
                    LogTextBox.AppendText(text + Environment.NewLine);
                    LogTextBox.ScrollToEnd();
                }
                return;
            }

            Application.Current.Dispatcher.BeginInvoke(new Action(() => {
                if (LogTextBox != null) {
                    LogTextBox.AppendText(text + Environment.NewLine);
                    LogTextBox.ScrollToEnd();
                }
            }));
        }
        catch {
            // ignored
        }
    }

    private void AppendColoredStatus(string text, bool isUsable) {
        try {
            if (Application.Current?.Dispatcher != null && !Application.Current.Dispatcher.CheckAccess()) {
                Application.Current.Dispatcher.BeginInvoke(new Action(() => AppendColoredStatus(text, isUsable)));
                return;
            }

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
        catch {
            // ignored
        }
    }

    #endregion
    
    #region TMP Debug Logger

    private static readonly object DbgLock = new();
    private static string _dbgFilePath;

    private static string GetDbgFilePath() {
        if (!string.IsNullOrEmpty(_dbgFilePath))
            return _dbgFilePath;

        var dir = Path.Combine(Path.GetTempPath(), "HFT_SharedTool");
        if (!Directory.Exists(dir))
            Directory.CreateDirectory(dir);

        _dbgFilePath = Path.Combine(dir, $"debug_{DateTime.Now:yyyyMMdd_HHmmss}_{Environment.UserName}.txt");
        return _dbgFilePath;
    }

    private static void Dbg(string msg, Exception ex = null) {
        try {
            var path = GetDbgFilePath();
            var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [T{Environment.CurrentManagedThreadId}] {msg}";
            lock (DbgLock) {
                File.AppendAllText(path, line + Environment.NewLine);

                if (ex == null) return;

                File.AppendAllText(path, @"EX: " + ex.GetType().FullName + Environment.NewLine);
                File.AppendAllText(path, @"MSG: " + ex.Message + Environment.NewLine);
                File.AppendAllText(path, @"STACK: " + ex.StackTrace + Environment.NewLine);

                if (ex.InnerException != null) {
                    File.AppendAllText(path, @"INNER EX: " + ex.InnerException.GetType().FullName + Environment.NewLine);
                    File.AppendAllText(path, @"INNER MSG: " + ex.InnerException.Message + Environment.NewLine);
                    File.AppendAllText(path, @"INNER STACK: " + ex.InnerException.StackTrace + Environment.NewLine);
                }
            }
        }
        catch {
            // ignored
        }
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

            const int maxAttempts = 10;
            var delayMs = 150;

            for (var attempt = 1; attempt <= maxAttempts; attempt++) {
                try {
                    using var writer = new StreamWriter(filePath, false);

                    foreach (var lic in infos) {
                        if (string.IsNullOrEmpty(lic.LicenseId))
                            continue;

                        writer.WriteLine(lic.LicenseId);

                        string loginUser;
                        string loginDate;
                        if (lic.Logins.Count > 0) {
                            var names = new string[lic.Logins.Count];
                            for (var j = 0; j < lic.Logins.Count; j++)
                                names[j] = lic.Logins[j].User;

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

                    return; // SUCCESS
                }
                catch (IOException ex) {
                    Dbg($"SharedLicenseFileService.Save: IO EX attempt={attempt}/{maxAttempts} filePath={filePath}", ex);
                    System.Threading.Thread.Sleep(delayMs);
                    delayMs = Math.Min(delayMs * 2, 2000);
                }
            }

            throw new IOException($"Nie udało się zapisać pliku po {maxAttempts} próbach: {filePath}");
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
    
    private static void DumpTeklaDiagnostics() {
        try {
            Dbg("DumpTeklaDiagnostics: ENTER");

            // 1) bitness naszego EXE
            Dbg($"DumpTeklaDiagnostics: Tool Is64BitProcess={Environment.Is64BitProcess} Is64BitOS={Environment.Is64BitOperatingSystem}");

            // 2) środowisko
            var xsDataDir = Environment.GetEnvironmentVariable("XSDATADIR") ?? "(null)";
            var xsBinDir = Environment.GetEnvironmentVariable("XSBIN") ?? "(null)";
            Dbg($"DumpTeklaDiagnostics: XSDATADIR={xsDataDir}");
            Dbg($"DumpTeklaDiagnostics: XSBIN={xsBinDir}");

            // 3) skąd ładuje się Tekla.Structures.Model.dll (MEGA ważne)
            try {
                var asm = typeof(Tekla.Structures.Model.Model).Assembly;
                Dbg($"DumpTeklaDiagnostics: Tekla.Structures.Model.dll Location={asm.Location}");
                Dbg($"DumpTeklaDiagnostics: Tekla.Structures.Model.dll Version={asm.GetName().Version}");
            }
            catch (Exception ex) {
                Dbg("DumpTeklaDiagnostics: could not read Tekla.Structures.Model assembly info", ex);
            }

            // 4) czy Tekla proces w ogóle istnieje + jego bitness
            try {
                var teklaProcs = System.Diagnostics.Process.GetProcesses()
                    .Where(p => {
                        try { return p.ProcessName.IndexOf("TeklaStructures", StringComparison.OrdinalIgnoreCase) >= 0; }
                        catch { return false; }
                    })
                    .ToList();

                Dbg($"DumpTeklaDiagnostics: TeklaStructures process count={teklaProcs.Count}");

                foreach (var p in teklaProcs) {
                    try {
                        Dbg($"DumpTeklaDiagnostics: Tekla PID={p.Id} Name={p.ProcessName} SessionId={p.SessionId}");

                        // bitness procesu Tekli
                        bool isWow64;
                        if (Environment.Is64BitOperatingSystem) {
                            isWow64 = IsWow64Process(p.Handle);
                            // jeśli WOW64=true -> proces 32-bit na 64-bit OS
                            var teklaIs64 = !isWow64;
                            Dbg($"DumpTeklaDiagnostics: Tekla PID={p.Id} Is64Bit={teklaIs64} (IsWow64={isWow64})");
                        }
                    }
                    catch (Exception ex) {
                        Dbg("DumpTeklaDiagnostics: reading Tekla process info EX", ex);
                    }
                }
            }
            catch (Exception ex) {
                Dbg("DumpTeklaDiagnostics: process scan EX", ex);
            }

            Dbg("DumpTeklaDiagnostics: EXIT");
        }
        catch {
            // ignored
        }
    }

    // P/Invoke do WOW64
    [System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool IsWow64Process(IntPtr hProcess, out bool wow64Process);

    private static bool IsWow64Process(IntPtr hProcess) {
        try {
            if (!Environment.Is64BitOperatingSystem) return false;
            return IsWow64Process(hProcess, out var wow64) && wow64;
        }
        catch {
            return false;
        }
    }
    
    private static string NormalizeUserName(string input) {
        if (string.IsNullOrEmpty(input))
            return input;

        var normalized = input.Normalize(NormalizationForm.FormD);
        var sb = new StringBuilder(normalized.Length);

        foreach (var c in normalized) {
            var uc = CharUnicodeInfo.GetUnicodeCategory(c);
            if (uc != UnicodeCategory.NonSpacingMark)
                sb.Append(c);
        }

        var s = sb.ToString().Normalize(NormalizationForm.FormC);

        s = s
            .Replace('ą', 'a').Replace('Ą', 'A')
            .Replace('ć', 'c').Replace('Ć', 'C')
            .Replace('ę', 'e').Replace('Ę', 'E')
            .Replace('ł', 'l').Replace('Ł', 'L')
            .Replace('ń', 'n').Replace('Ń', 'N')
            .Replace('ó', 'o').Replace('Ó', 'O')
            .Replace('ś', 's').Replace('Ś', 'S')
            .Replace('ż', 'z').Replace('Ż', 'Z')
            .Replace('ź', 'z').Replace('Ź', 'Z');

        return s;
    }
    
    private static string SafeReadLogToTemp(string logPath) {
        Dbg($"SafeReadLogToTemp: logPath={logPath}");
        if (string.IsNullOrEmpty(logPath) || !File.Exists(logPath)) {
            Dbg("SafeReadLogToTemp: file missing");
            return null;
        }

        try {
            var tempPath = Path.GetTempFileName();
            using var src = new FileStream(logPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var dst = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None);
            src.CopyTo(dst);
            Dbg($"SafeReadLogToTemp: OK temp={tempPath}");
            return tempPath;
        }
        catch (Exception ex) {
            Dbg("SafeReadLogToTemp: EXCEPTION", ex);
            return null;
        }
    }

    private async Task WaitAndFinalizeReadInAsync(string logPath, SharedLicenseInfo info, string userName, DateTime fromUtc) {
        Dbg($"WaitAndFinalizeReadInAsync: ENTER logPath={logPath} marker=Read-in result: OK. fromUtc={fromUtc:O}");

        bool ok;
        try {
            ok = await WaitForModelSharingConfirmationAsync(logPath, "Read-in result: OK.", fromUtc).ConfigureAwait(false);
        }
        catch (Exception ex) {
            Dbg("WaitAndFinalizeReadInAsync: EXCEPTION while waiting", ex);
            ok = false;
        }

        Dbg($"WaitAndFinalizeReadInAsync: WAIT RESULT ok={ok}");

        Application.Current.Dispatcher.Invoke(() => {
            Dbg("WaitAndFinalizeReadInAsync: DISPATCHER INVOKE ENTER");
            if (!ok) {
                Dbg("WaitAndFinalizeReadInAsync: NOT OK -> exit/close if mode=readin");
                if (_mode == "readin") Close();
                return;
            }

            var now = DateTime.Now;
            Dbg($"WaitAndFinalizeReadInAsync: finalize ReadIn user={userName} now={now:yyyy-MM-dd HH:mm:ss}");

            SharedLicenseManager.ReadIn(info, userName, now);

            try {
                SharedLicenseFileService.Save(LicenseFilePath, info);
                Dbg("WaitAndFinalizeReadInAsync: Save OK");
            }
            catch (Exception ex) {
                Dbg("WaitAndFinalizeReadInAsync: Save EXCEPTION", ex);
                throw;
            }

            AppendLog($"READ IN - {userName} - {now.ToString(DateFormat)}");
            if (_mode == "readin") Close();
        });
    }

    private async Task WaitAndFinalizeWriteOutAsync(string logPath, SharedLicenseInfo info, string userName, DateTime fromUtc) {
        Dbg($"WaitAndFinalizeWriteOutAsync: ENTER logPath={logPath} marker=WriteOut OK fromUtc={fromUtc:O}");

        bool ok;
        try {
            ok = await WaitForModelSharingConfirmationAsync(logPath, "WriteOut OK", fromUtc).ConfigureAwait(false);
        }
        catch (Exception ex) {
            Dbg("WaitAndFinalizeWriteOutAsync: EXCEPTION while waiting", ex);
            ok = false;
        }

        Dbg($"WaitAndFinalizeWriteOutAsync: WAIT RESULT ok={ok}");

        Application.Current.Dispatcher.Invoke(() => {
            Dbg("WaitAndFinalizeWriteOutAsync: DISPATCHER INVOKE ENTER");
            if (!ok) {
                Dbg("WaitAndFinalizeWriteOutAsync: NOT OK -> exit/close if mode=writeout");
                if (_mode == "writeout") Close();
                return;
            }

            var now = DateTime.Now;
            Dbg($"WaitAndFinalizeWriteOutAsync: finalize WriteOut user={userName} now={now:yyyy-MM-dd HH:mm:ss}");

            SharedLicenseManager.WriteOut(info, userName, now);

            try {
                SharedLicenseFileService.Save(LicenseFilePath, info);
                Dbg("WaitAndFinalizeWriteOutAsync: Save OK");
            }
            catch (Exception ex) {
                Dbg("WaitAndFinalizeWriteOutAsync: Save EXCEPTION", ex);
                throw;
            }

            AppendLog($"WRITE OUT - {userName} - {now.ToString(DateFormat)}");
            if (_mode == "writeout") Close();
        });
    }
    
    private static async Task<bool> WaitForModelSharingConfirmationAsync(string logPath, string marker, DateTime fromUtc, int maxWaitSeconds = 300, int delaySeconds = 5) {
        Dbg($"WaitForModelSharingConfirmationAsync: ENTER marker='{marker}' fromUtc={fromUtc:O} maxWait={maxWaitSeconds}s delay={delaySeconds}s");

        var start = DateTime.UtcNow;
        var timeout = TimeSpan.FromSeconds(maxWaitSeconds);
        var delay = TimeSpan.FromSeconds(delaySeconds);

        if (fromUtc.Kind != DateTimeKind.Utc) fromUtc = fromUtc.ToUniversalTime();

        while (DateTime.UtcNow - start < timeout) {
            try {
                var tempPath = SafeReadLogToTemp(logPath);
                if (tempPath != null) {
                    bool found;
                    try {
                        found = HasConfirmationLine(tempPath, marker, fromUtc);
                        Dbg($"WaitForModelSharingConfirmationAsync: HasConfirmationLine={found}");
                    }
                    finally {
                        try { File.Delete(tempPath); } catch (Exception ex) { Dbg("WaitForModelSharingConfirmationAsync: temp delete EX", ex); }
                    }

                    if (found) {
                        Dbg("WaitForModelSharingConfirmationAsync: FOUND -> true");
                        return true;
                    }
                }
            }
            catch (Exception ex) {
                Dbg("WaitForModelSharingConfirmationAsync: LOOP EXCEPTION", ex);
            }

            await Task.Delay(delay).ConfigureAwait(false);
        }

        Dbg("WaitForModelSharingConfirmationAsync: TIMEOUT -> false");
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
        Dbg($"HasConfirmationLine: ENTER logPath={logPath} marker='{marker}' fromUtc={fromUtc:O}");
        
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
            if (System.Threading.Interlocked.Exchange(ref _autoLoginWatcherStarted, 1) == 1) {
                Dbg("StartAutoLoginWatcher: already started -> return");
                return;
            }
            Dbg($"StartAutoLoginWatcher: ENTER mode={_mode} autoLoginDone={_autoLoginDone} logoutDone={_logoutDone}");

            try {
                ClearLog();
                AppendLog("Czekam na wpis UserInfo w logu Tekli...");
                Dbg("StartAutoLoginWatcher: UI cleared + message written.");

                var start = DateTime.Now;
                var maxWait = TimeSpan.FromMinutes(5);
                var loop = 0;

                while (!_autoLoginDone && DateTime.Now - start < maxWait) {
                    loop++;
                    var elapsed = DateTime.Now - start;
                    Dbg($"StartAutoLoginWatcher: LOOP #{loop} elapsed={elapsed.TotalSeconds:0}s");

                    try {
                        if (TryGetLicenseFromLog(out var licenseId, out var loginTime)) {
                            Dbg($"StartAutoLoginWatcher: TryGetLicenseFromLog=TRUE licenseId={licenseId} loginTime={loginTime:yyyy-MM-dd HH:mm:ss}");

                            var userName = Environment.UserName;
                            Dbg($"StartAutoLoginWatcher: current userName={userName}");

                            var info = SharedLicenseFileService.LoadOrCreate(LicenseFilePath, licenseId);
                            Dbg("StartAutoLoginWatcher: LoadOrCreate OK");

                            SharedLicenseManager.Login(info, userName, loginTime);
                            Dbg("StartAutoLoginWatcher: SharedLicenseManager.Login OK");

                            SharedLicenseFileService.Save(LicenseFilePath, info);
                            Dbg($"StartAutoLoginWatcher: Save OK -> {LicenseFilePath}");

                            _autoLoginDone = true;
                            AppendLog($"LOG IN - {userName} - {loginTime.ToString(DateFormat)}");
                            Dbg("StartAutoLoginWatcher: _autoLoginDone=true + UI log appended");

                            break;
                        }

                        Dbg("StartAutoLoginWatcher: TryGetLicenseFromLog=FALSE (no UserInfo yet)");
                    }
                    catch (Exception ex) {
                        Dbg("StartAutoLoginWatcher: EXCEPTION inside loop", ex);
                    }

                    await Task.Delay(TimeSpan.FromSeconds(5)).ConfigureAwait(false);
                }

                Dbg($"StartAutoLoginWatcher: EXIT autoLoginDone={_autoLoginDone} elapsed={(DateTime.Now - start).TotalSeconds:0}s");
            }
            catch (Exception ex) {
                Dbg("StartAutoLoginWatcher: OUTER EXCEPTION", ex);  
            }
        }
        catch {
            //ignored
        }
    }

    private void OnModelUnloading() {
        Dbg("OnModelUnloading: EVENT");
        HandleModelClosed();
    }

    private void OnTeklaStructuresExit() {
        Dbg("OnTeklaStructuresExit: EVENT");
        HandleModelClosed();
    }

    private void HandleModelClosed() {
        Dbg($"HandleModelClosed: ENTER mode={_mode} logoutDone={_logoutDone}");

        try {
            if (_mode == "autologin" && !_logoutDone) {
                Dbg("HandleModelClosed: calling RemoveCurrentUserLogin()");
                RemoveCurrentUserLogin();
                _logoutDone = true;
                Dbg("HandleModelClosed: RemoveCurrentUserLogin DONE, logoutDone=true");
            }
        }
        catch (Exception ex) {
            Dbg("HandleModelClosed: EXCEPTION during logout", ex);
        }

        try {
            Dbg("HandleModelClosed: Dispatcher.Invoke(Close)...");
            Application.Current.Dispatcher.Invoke(Close);
            Dbg("HandleModelClosed: Close invoked");
        }
        catch (Exception ex) {
            Dbg("HandleModelClosed: EXCEPTION while closing", ex);
        }
    }

    private static void RemoveCurrentUserLogin() {
        Dbg("RemoveCurrentUserLogin: ENTER");

        try {
            if (!TryGetLicenseFromLog(out var licenseId, out var loginTime)) {
                Dbg("RemoveCurrentUserLogin: TryGetLicenseFromLog=FALSE -> nothing to remove");
                return;
            }

            Dbg($"RemoveCurrentUserLogin: licenseId={licenseId} loginTime={loginTime:yyyy-MM-dd HH:mm:ss}");

            var info = SharedLicenseFileService.LoadOrCreate(LicenseFilePath, licenseId);
            Dbg("RemoveCurrentUserLogin: LoadOrCreate OK");

            var userName = Environment.UserName;
            Dbg($"RemoveCurrentUserLogin: userName={userName}");

            SharedLicenseManager.Logout(info, userName);
            Dbg("RemoveCurrentUserLogin: SharedLicenseManager.Logout OK");

            SharedLicenseFileService.Save(LicenseFilePath, info);
            Dbg($"RemoveCurrentUserLogin: Save OK -> {LicenseFilePath}");
        }
        catch (Exception ex) {
            Dbg("RemoveCurrentUserLogin: EXCEPTION", ex);
            throw;
        }
    }

    protected override void OnClosed(EventArgs e) {
        Dbg($"OnClosed: ENTER mode={_mode} logoutDone={_logoutDone} eventsNull={_events == null}");

        try {
            if (_mode == "autologin" && !_logoutDone) {
                Dbg("OnClosed: autologin & logout not done -> RemoveCurrentUserLogin()");
                RemoveCurrentUserLogin();
                _logoutDone = true;
                Dbg("OnClosed: logoutDone=true");
            }
        }
        catch (Exception ex) {
            Dbg("OnClosed: EXCEPTION during RemoveCurrentUserLogin", ex);
        }

        try {
            if (_events != null) {
                Dbg("OnClosed: unregistering Tekla events...");
                _events.ModelUnloading -= OnModelUnloading;
                _events.TeklaStructuresExit -= OnTeklaStructuresExit;
                _events.UnRegister();
                _events = null;
                Dbg("OnClosed: events unregistered");
            }
        }
        catch (Exception ex) {
            Dbg("OnClosed: EXCEPTION during events unregister", ex);
        }

        base.OnClosed(e);
        Dbg("OnClosed: EXIT");
    }

    private static bool TryGetLicenseFromLog(out string licenseId, out DateTime loginTime) {
        licenseId = null;
        loginTime = DateTime.MinValue;

        var rawUser = Environment.UserName;
        var normalizedUser = NormalizeUserName(rawUser);

        Dbg($"TryGetLicenseFromLog: rawUser={rawUser} normalizedUser={normalizedUser}");

        const string dir = @"C:\TeklaStructuresModels";

        var candidates = new List<string> {
            Path.Combine(dir, $"TeklaStructures_{rawUser}.log"),
            Path.Combine(dir, $"TeklaStructures_{normalizedUser}.log")
        };

        string logPath = null;

        foreach (var c in candidates) {
            Dbg($"TryGetLicenseFromLog: checking {c}");
            if (File.Exists(c)) {
                logPath = c;
                Dbg($"TryGetLicenseFromLog: FOUND log {logPath}");
                break;
            }
        }

        if (logPath == null) {
            try {
                if (Directory.Exists(dir)) {
                    var logs = Directory.GetFiles(dir, "TeklaStructures_*.log", SearchOption.TopDirectoryOnly);
                    if (logs.Length > 0) {
                        logPath = logs.OrderByDescending(File.GetLastWriteTime).FirstOrDefault();
                        Dbg($"TryGetLicenseFromLog: fallback newest log={logPath}");
                    }
                    else {
                        Dbg("TryGetLicenseFromLog: fallback scan -> no TeklaStructures_*.log files");
                    }
                }
                else {
                    Dbg($"TryGetLicenseFromLog: dir missing: {dir}");
                }
            }
            catch (Exception ex) {
                Dbg("TryGetLicenseFromLog: fallback scan EX", ex);
            }
        }

        if (logPath == null) {
            Dbg("TryGetLicenseFromLog: no matching log file found");
            return false;
        }

        Dbg($"TryGetLicenseFromLog: ENTER logPath={logPath}");

        try {
            var tempPath = SafeReadLogToTemp(logPath);
            if (tempPath == null) {
                Dbg("TryGetLicenseFromLog: SafeReadLogToTemp returned NULL");
                return false;
            }

            Dbg($"TryGetLicenseFromLog: tempPath={tempPath}");

            string[] lines;
            try {
                lines = File.ReadAllLines(tempPath);
            }
            finally {
                try {
                    File.Delete(tempPath);
                    Dbg("TryGetLicenseFromLog: temp deleted");
                }
                catch (Exception ex) {
                    Dbg("TryGetLicenseFromLog: temp delete EX", ex);
                }
            }

            Dbg($"TryGetLicenseFromLog: lines={lines.Length}");

            for (var i = lines.Length - 1; i >= 0; i--) {
                var line = lines[i];

                if (line.IndexOf("UserInfo", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                Dbg($"TryGetLicenseFromLog: UserInfo found at lineIndex={i}");
                Dbg($"TryGetLicenseFromLog: line='{line}'");

                var parts = line.Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries);
                Dbg($"TryGetLicenseFromLog: partsCount={parts.Length}");

                if (parts.Length < 3)
                    continue;

                string lic = null;
                foreach (var p in parts) {
                    if (p.Contains("@") && !p.StartsWith("http", StringComparison.OrdinalIgnoreCase)) {
                        lic = p.Trim();
                        break;
                    }
                }

                if (string.IsNullOrEmpty(lic)) {
                    Dbg("TryGetLicenseFromLog: no license token found in UserInfo line");
                    continue;
                }

                licenseId = lic;
                loginTime = DateTime.Now;
                Dbg($"TryGetLicenseFromLog: SUCCESS licenseId={licenseId} loginTime={loginTime:yyyy-MM-dd HH:mm:ss}");
                return true;
            }

            Dbg("TryGetLicenseFromLog: END no UserInfo with license found");
            return false;
        }
        catch (Exception ex) {
            Dbg("TryGetLicenseFromLog: EXCEPTION", ex);
            return false;
        }
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