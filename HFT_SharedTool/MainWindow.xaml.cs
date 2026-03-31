using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Shapes;
using Path = System.IO.Path;
using TSM = Tekla.Structures.Model;
using TSMO = Tekla.Structures.Model.Operations;

namespace HFT_SharedTool;

public partial class MainWindow {
    #region Constants

    private static readonly string SelectedLicenseFilePath =
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "HFT_SharedTool",
            $"selected_license_{Environment.UserName}.txt");

    #endregion

    #region Shared Model Helpers

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

            basePath = basePath.TrimEnd(
                Path.DirectorySeparatorChar,
                Path.AltDirectorySeparatorChar);

            logPath = Path.Combine(basePath, "logs", "modelsharing.log");
            return true;
        }
        catch {
            logPath = null;
            return false;
        }
    }

    private static bool IsAnySharedModelStillOpen() {
        try {
            var model = new TSM.Model();
            if (!model.GetConnectionStatus()) {
                Dbg("IsAnySharedModelStillOpen: connection=FALSE");
                return false;
            }

            var info = model.GetInfo();
            var result = info.SharedModel;

            Dbg($"IsAnySharedModelStillOpen: connection=TRUE sharedModel={result}");
            return result;
        }
        catch (Exception ex) {
            Dbg("IsAnySharedModelStillOpen: EXCEPTION", ex);
            return false;
        }
    }

    private static async Task<bool> IsAnySharedModelStillOpenWithRetryAsync(
        int attempts = 6,
        int delayMilliseconds = 1000) {
        for (var attempt = 1; attempt <= attempts; attempt++) {
            try {
                var result = await Application.Current.Dispatcher.InvokeAsync(IsAnySharedModelStillOpen);
                Dbg(
                    $"IsAnySharedModelStillOpenWithRetryAsync: attempt={attempt}/{attempts} result={result}");

                if (result)
                    return true;
            }
            catch (Exception ex) {
                Dbg($"IsAnySharedModelStillOpenWithRetryAsync: attempt={attempt} EXCEPTION", ex);
            }

            if (attempt < attempts)
                await Task.Delay(delayMilliseconds);
        }

        Dbg("IsAnySharedModelStillOpenWithRetryAsync: no shared model detected");
        return false;
    }

    #endregion

    #region Fields

    private readonly string _mode;
    private readonly string _trimbleEmail;
    private bool _autoLoginDone;
    private int _autoLoginWatcherStarted;
    private string _autoLoginLicenseId;
    private TSM.Events _events;
    private int _logoutDone;
    private string _selectedLicenseId;
    private int _readInStarted;
    private int _writeOutStarted;

    #endregion

    #region TMP Debug Logger Fields

    private static readonly object DbgLock = new();
    private static string _dbgFilePath;

    #endregion

    #region Constructors

    public MainWindow() : this("standalone") {
    }

    public MainWindow(string mode) {
        _mode = (mode ?? "standalone").Trim().ToLowerInvariant();
        _trimbleEmail = TeklaAccountService.GetTrimbleEmail();

        Opacity = 0;
        ShowInTaskbar = false;

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

        Loaded += async (_, _) => await InitializeAfterLoadAsync();
    }

    private async Task InitializeAfterLoadAsync() {
        UpdateThemeToggleIconColor();

        if (_mode is "standalone" or "check" or "ckeck") {
            Opacity = 1;
            ShowInTaskbar = true;
            RefreshButton.Visibility = Visibility.Visible;
            ModelDrawingLabel.Text = "Tryb odczytu pliku licencji";
            Dispatcher.BeginInvoke(new Action(BtnCheck_Click));
            return;
        }

        var myModel = await WaitForSharedModelOnUiThreadAsync();

        if (myModel == null) {
            Dbg("InitializeAfterLoadAsync: model=null -> Close");
            Close();
            return;
        }

        TSM.ModelInfo modelInfo;

        try {
            modelInfo = myModel.GetInfo();
            Dbg($"InitializeAfterLoadAsync: sharedModel={modelInfo.SharedModel}");
        }
        catch (Exception ex) {
            Dbg("InitializeAfterLoadAsync: GetInfo EX", ex);
            Close();
            return;
        }

        if (!modelInfo.SharedModel) {
            Dbg("InitializeAfterLoadAsync: sharedModel=FALSE -> Close");
            Close();
            return;
        }

        if (!EnsureSelectedLicenseId()) {
            Close();
            return;
        }

        _events = new TSM.Events();
        _events.ModelUnloading += OnModelUnloading;
        _events.TeklaStructuresExit += OnTeklaStructuresExit;
        _events.Register();

        var modelName = modelInfo.ModelName.Replace(".db1", "");
        ModelDrawingLabel.Text = $"Połączono z {modelName}";
        Dbg($"InitializeAfterLoadAsync: modelName={modelName}");

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
                Opacity = 1;
                ShowInTaskbar = true;
                Dispatcher.BeginInvoke(new Action(BtnCheck_Click));
                break;
        }
    }

    private async Task<TSM.Model> WaitForSharedModelOnUiThreadAsync(
        int maxWaitSeconds = 300,
        int delaySeconds = 5) {
        Dbg($"WaitForSharedModelOnUiThreadAsync: ENTER maxWaitSeconds={maxWaitSeconds} delaySeconds={delaySeconds}");

        var timeout = TimeSpan.FromSeconds(maxWaitSeconds);
        var delay = TimeSpan.FromSeconds(delaySeconds);
        var start = DateTime.Now;
        var attempt = 0;

        while (DateTime.Now - start < timeout) {
            attempt++;

            TSM.Model model = null;
            var connected = false;
            var sharedModel = false;

            try {
                await Dispatcher.InvokeAsync(() => {
                    try {
                        model = new TSM.Model();
                        connected = model.GetConnectionStatus();

                        if (!connected) {
                            Dbg($"WaitForSharedModelOnUiThreadAsync: attempt={attempt} connected=FALSE");
                            return;
                        }

                        var info = model.GetInfo();
                        sharedModel = info.SharedModel;

                        Dbg(
                            $"WaitForSharedModelOnUiThreadAsync: attempt={attempt} connected=TRUE sharedModel={sharedModel}");
                    }
                    catch (Exception ex) {
                        Dbg($"WaitForSharedModelOnUiThreadAsync: attempt={attempt} EX", ex);
                        model = null;
                        connected = false;
                        sharedModel = false;
                    }
                });
            }
            catch (Exception ex) {
                Dbg($"WaitForSharedModelOnUiThreadAsync: attempt={attempt} Dispatcher EX", ex);
            }

            if (connected && sharedModel && model != null) {
                Dbg("WaitForSharedModelOnUiThreadAsync: SUCCESS");
                return model;
            }

            await Task.Delay(delay);
        }

        Dbg("WaitForSharedModelOnUiThreadAsync: TIMEOUT");
        return null;
    }

    #endregion

    #region Window Helpers

    private void HideWindow() {
        WindowState = WindowState.Minimized;
        ShowInTaskbar = false;
        Visibility = Visibility.Hidden;
    }

    private static void RunOnUi(Action action) {
        try {
            if (Application.Current?.Dispatcher == null || Application.Current.Dispatcher.CheckAccess()) {
                action();
                return;
            }

            Application.Current.Dispatcher.BeginInvoke(action);
        }
        catch {
            // ignored
        }
    }

    #endregion

    #region TMP Debug Logger

    private static string GetDbgFilePath() {
        if (!string.IsNullOrEmpty(_dbgFilePath))
            return _dbgFilePath;

        var dir = Path.Combine(Path.GetTempPath(), "HFT_SharedTool");
        if (!Directory.Exists(dir))
            Directory.CreateDirectory(dir);

        _dbgFilePath = Path.Combine(dir,
            $"debug_{DateTime.Now:yyyyMMdd_HHmmss}_{Environment.UserName}.txt");
        return _dbgFilePath;
    }

    private static void Dbg(string msg, Exception ex = null) {
        try {
            var path = GetDbgFilePath();
            var line =
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [T{Environment.CurrentManagedThreadId}] {msg}";

            lock (DbgLock) {
                File.AppendAllText(path, line + Environment.NewLine);

                if (ex == null) return;

                File.AppendAllText(path, @"EX: " + ex.GetType().FullName + Environment.NewLine);
                File.AppendAllText(path, @"MSG: " + ex.Message + Environment.NewLine);
                File.AppendAllText(path, @"STACK: " + ex.StackTrace + Environment.NewLine);

                if (ex.InnerException != null) {
                    File.AppendAllText(path,
                        @"INNER EX: " + ex.InnerException.GetType().FullName + Environment.NewLine);
                    File.AppendAllText(path,
                        @"INNER MSG: " + ex.InnerException.Message + Environment.NewLine);
                    File.AppendAllText(path,
                        @"INNER STACK: " + ex.InnerException.StackTrace + Environment.NewLine);
                }
            }
        }
        catch {
            // ignored
        }
    }

    #endregion

    #region Logging Helpers

    private void ClearLog() {
        try {
            RunOnUi(() => LogPanel?.Children.Clear());
        }
        catch {
            // ignored
        }
    }

    private void AppendLog(string text) {
        try {
            RunOnUi(() => {
                if (LogPanel == null) return;

                LogPanel.Children.Add(new TextBlock {
                    Text = text,
                    FontSize = 13,
                    Foreground = Application.Current.FindResource("TextPrimaryBrush") as Brush,
                    Margin = new Thickness(0, 3, 0, 3)
                });

                LogScrollViewer.ScrollToEnd();
            });
        }
        catch {
            // ignored
        }
    }

    private void AppendStatusRow(SharedLicenseInfo info, bool isUsable) {
        try {
            RunOnUi(() => {
                if (LogPanel == null) return;

                var hasRead = info.ReadTime.HasValue;
                var hasWrite = info.WriteTime.HasValue;

                string actionUser;
                DateTime? actionTime;

                if (hasRead && (!hasWrite || info.ReadTime.Value > info.WriteTime.Value)) {
                    actionUser = info.ReadUser ?? "";
                    actionTime = info.ReadTime;
                }
                else if (hasWrite) {
                    actionUser = info.WriteUser ?? "";
                    actionTime = info.WriteTime;
                }
                else {
                    actionUser = "";
                    actionTime = null;
                }

                var currentUser = Environment.UserName;

                var loggedUsers = info.Logins
                    .Where(x => x != null && !string.IsNullOrWhiteSpace(x.User))
                    .Select(x => x.User)
                    .ToList();

                var isCurrentUserLoggedHere = loggedUsers.Any(user =>
                    string.Equals(user, currentUser, StringComparison.OrdinalIgnoreCase));

                var tooltipText = info.NextUsable.HasValue
                    ? $"Dostępna od: {info.NextUsable.Value:yyyy-MM-dd HH:mm}"
                    : "Dostępna teraz";

                var isDark = ThemeService.IsDark;
                var dotOk = isDark ? Color.FromRgb(0x6E, 0xC9, 0x6E) : Color.FromRgb(0x16, 0xA3, 0x4A);
                var dotActive = isDark ? Color.FromRgb(0x7C, 0xFF, 0x9A) : Color.FromRgb(0x63, 0xF5, 0x87);
                var dotErr = isDark ? Color.FromRgb(0xE2, 0x4B, 0x4A) : Color.FromRgb(0xDC, 0x26, 0x26);

                var infoFg = isDark
                    ? new SolidColorBrush(Color.FromArgb(0xAA, 0xB8, 0xB8, 0xB8))
                    : new SolidColorBrush(Color.FromRgb(0x6B, 0x72, 0x80));

                var row = new Grid {
                    Margin = new Thickness(0, 4, 0, 4)
                };

                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(18) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(110) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                var dotColor = isUsable ? dotOk : dotErr;

                var dot = new Ellipse {
                    Width = 10,
                    Height = 10,
                    Fill = new SolidColorBrush(dotColor),
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Left,
                    ToolTip = tooltipText
                };

                if (isCurrentUserLoggedHere) {
                    dot.Effect = new DropShadowEffect {
                        Color = dotActive,
                        BlurRadius = 12,
                        ShadowDepth = 0,
                        Opacity = 1
                    };
                    dot.Fill = new SolidColorBrush(dotActive);
                }

                Grid.SetColumn(dot, 0);
                row.Children.Add(dot);

                var licenseText = new TextBlock {
                    Text = info.LicenseId,
                    FontSize = 13,
                    Foreground = Application.Current.FindResource("TextPrimaryBrush") as Brush,
                    VerticalAlignment = VerticalAlignment.Center
                };
                Grid.SetColumn(licenseText, 1);
                row.Children.Add(licenseText);

                var loggedUsersPanel = new WrapPanel {
                    Orientation = Orientation.Horizontal,
                    VerticalAlignment = VerticalAlignment.Center
                };

                if (loggedUsers.Count == 0)
                    loggedUsersPanel.Children.Add(MakeBadge("WOLNE", true, false));
                else
                    foreach (var user in loggedUsers) {
                        var isCurrentUser = string.Equals(user, currentUser, StringComparison.OrdinalIgnoreCase);
                        loggedUsersPanel.Children.Add(MakeBadge(user, false, isCurrentUser));
                    }

                Grid.SetColumn(loggedUsersPanel, 2);
                row.Children.Add(loggedUsersPanel);

                if (actionTime.HasValue && !string.IsNullOrWhiteSpace(actionUser)) {
                    var actionPanel = new StackPanel {
                        Orientation = Orientation.Horizontal,
                        VerticalAlignment = VerticalAlignment.Center,
                        Margin = new Thickness(6, 0, 0, 0)
                    };

                    actionPanel.Children.Add(new TextBlock {
                        Text = $"{FormatTimeAgo(actionTime.Value)} przez ",
                        FontSize = 11,
                        Foreground = infoFg,
                        VerticalAlignment = VerticalAlignment.Center
                    });

                    var isActionByCurrentUser =
                        string.Equals(actionUser, currentUser, StringComparison.OrdinalIgnoreCase);

                    actionPanel.Children.Add(MakeBadge(actionUser, false, isActionByCurrentUser));

                    Grid.SetColumn(actionPanel, 3);
                    row.Children.Add(actionPanel);
                }

                LogPanel.Children.Add(row);
                LogScrollViewer.ScrollToEnd();
            });
        }
        catch {
            // ignored
        }
    }

    private static string FormatTimeAgo(DateTime dateTime) {
        var now = DateTime.Now;
        var diff = now - dateTime;

        if (diff.TotalMinutes < 0)
            diff = TimeSpan.Zero;

        switch (diff.TotalDays) {
            case >= 2:
                return $"{(int)diff.TotalDays} dni temu";
            case >= 1:
                return "1 dzień temu";
        }

        var totalHours = (int)diff.TotalHours;
        var minutes = diff.Minutes;

        return $"{totalHours}:{minutes:00}h temu";
    }

    private static Border MakeBadge(string text, bool isFree, bool isCurrentUser) {
        Color bgColor;
        Color fgColor;

        if (isFree) {
            bgColor = Color.FromArgb(0x24, 0x90, 0x90, 0x90);
            fgColor = Color.FromArgb(0xCC, 0xA0, 0xA0, 0xA0);
        }
        else if (isCurrentUser) {
            if (ThemeService.IsDark) {
                bgColor = Color.FromArgb(0x32, 0x4C, 0xD9, 0x6B);
                fgColor = Color.FromRgb(0x7C, 0xFF, 0x9A);
            }
            else {
                bgColor = Color.FromArgb(0x2C, 0x22, 0xC5, 0x5E);
                fgColor = Color.FromRgb(0x16, 0xA3, 0x4A);
            }
        }
        else {
            if (ThemeService.IsDark) {
                bgColor = Color.FromArgb(0x30, 0x3D, 0x7E, 0xA6);
                fgColor = Color.FromRgb(0x7D, 0xD3, 0xFC);
            }
            else {
                bgColor = Color.FromArgb(0x22, 0x00, 0x5F, 0xB8);
                fgColor = Color.FromRgb(0x00, 0x5F, 0xB8);
            }
        }

        var border = new Border {
            Background = new SolidColorBrush(bgColor),
            CornerRadius = new CornerRadius(10),
            Padding = new Thickness(7, 1, 7, 2),
            Margin = new Thickness(0, 0, 6, 0),
            VerticalAlignment = VerticalAlignment.Center,
            Child = new TextBlock {
                Text = text,
                FontSize = 11,
                Foreground = new SolidColorBrush(fgColor)
            }
        };

        return border;
    }

    #endregion

    #region Main Event Handlers

    private void BtnReadIn_Click() {
        try {
            if (Interlocked.Exchange(ref _readInStarted, 1) == 1) {
                Dbg("BtnReadIn_Click: already started -> return");
                return;
            }

            ClearLog();

            if (!EnsureSelectedLicenseId()) {
                Interlocked.Exchange(ref _readInStarted, 0);
                return;
            }

            SaveSelectedLicenseId(_selectedLicenseId);

            if (!TryGetModelSharingLogPath(out var modelSharingLogPath)) {
                Interlocked.Exchange(ref _readInStarted, 0);
                return;
            }

            var info = SharedLicenseFileService.LoadOrCreate(
                SharedConstants.LicenseFilePath, _selectedLicenseId);
            var userName = Environment.UserName;

            _ = WaitAndFinalizeReadInAsync(modelSharingLogPath, info, userName);
        }
        catch (Exception ex) {
            Interlocked.Exchange(ref _readInStarted, 0);
            MessageBox.Show($"READ IN: wyjątek: {ex.Message}");
        }
    }

    private void BtnWriteOut_Click() {
        try {
            if (Interlocked.Exchange(ref _writeOutStarted, 1) == 1) {
                Dbg("BtnWriteOut_Click: already started -> return");
                return;
            }

            ClearLog();

            if (!EnsureSelectedLicenseId()) {
                Interlocked.Exchange(ref _writeOutStarted, 0);
                return;
            }

            SaveSelectedLicenseId(_selectedLicenseId);

            if (!TryGetModelSharingLogPath(out var modelSharingLogPath)) {
                Interlocked.Exchange(ref _writeOutStarted, 0);
                return;
            }

            var info = SharedLicenseFileService.LoadOrCreate(
                SharedConstants.LicenseFilePath, _selectedLicenseId);
            var userName = Environment.UserName;

            _ = WaitAndFinalizeWriteOutAsync(modelSharingLogPath, info, userName);
        }
        catch (Exception ex) {
            Interlocked.Exchange(ref _writeOutStarted, 0);
            MessageBox.Show($"WRITE OUT: wyjątek: {ex.Message}");
        }
    }

    private void BtnCheck_Click() {
        ClearLog();

        var infos = SharedLicenseFileService.LoadAll(SharedConstants.LicenseFilePath);
        if (infos == null || infos.Count == 0) {
            AppendLog("Brak danych licencji: " + SharedConstants.LicenseFilePath);
            return;
        }

        var currentUser = Environment.UserName;
        var now = DateTime.Now;

        foreach (var info in infos) {
            SharedLicenseManager.FormatStatus(info, now, currentUser, out var isUsable);
            AppendStatusRow(info, isUsable);
        }
    }

    #endregion

    #region License Selection Helpers

    private static void DeleteSelectedLicenseId() {
        try {
            if (!File.Exists(SelectedLicenseFilePath)) {
                Dbg("DeleteSelectedLicenseId: file does not exist");
                return;
            }

            File.Delete(SelectedLicenseFilePath);
            Dbg("DeleteSelectedLicenseId: deleted");
        }
        catch (Exception ex) {
            Dbg("DeleteSelectedLicenseId: EXCEPTION", ex);
        }
    }

    private static void SaveSelectedLicenseId(string licenseId) {
        try {
            if (string.IsNullOrWhiteSpace(licenseId))
                return;

            var dir = Path.GetDirectoryName(SelectedLicenseFilePath);
            if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            File.WriteAllText(SelectedLicenseFilePath, licenseId.Trim());
            Dbg($"SaveSelectedLicenseId: saved={licenseId}");
        }
        catch (Exception ex) {
            Dbg("SaveSelectedLicenseId: EXCEPTION", ex);
        }
    }

    private static string LoadSelectedLicenseId() {
        try {
            if (!File.Exists(SelectedLicenseFilePath)) {
                Dbg("LoadSelectedLicenseId: file does not exist");
                return null;
            }

            var value = File.ReadAllText(SelectedLicenseFilePath).Trim();

            if (string.IsNullOrWhiteSpace(value)) {
                Dbg("LoadSelectedLicenseId: value is empty");
                return null;
            }

            Dbg($"LoadSelectedLicenseId: loaded={value}");
            return value;
        }
        catch (Exception ex) {
            Dbg("LoadSelectedLicenseId: EXCEPTION", ex);
            return null;
        }
    }

    private static List<string> FindLicenseIdsByCurrentUserInLicenseFile() {
        try {
            var currentUser = Environment.UserName;

            if (string.IsNullOrWhiteSpace(currentUser)) {
                Dbg("FindLicenseIdsByCurrentUserInLicenseFile: current user is empty");
                return [];
            }

            var infos = SharedLicenseFileService.LoadAll(SharedConstants.LicenseFilePath);
            var result = (
                from info in infos
                where info?.Logins != null && !string.IsNullOrWhiteSpace(info.LicenseId)
                let hasUser = info.Logins.Any(login =>
                    login != null &&
                    !string.IsNullOrWhiteSpace(login.User) &&
                    string.Equals(login.User, currentUser, StringComparison.OrdinalIgnoreCase))
                where hasUser
                select info.LicenseId
            ).ToList();

            Dbg(
                $"FindLicenseIdsByCurrentUserInLicenseFile: found {result.Count} license(s) for user={currentUser}");
            return result;
        }
        catch (Exception ex) {
            Dbg("FindLicenseIdsByCurrentUserInLicenseFile: EXCEPTION", ex);
            return [];
        }
    }

    private bool ResolveSelectedLicenseIdForAction() {
        if (!string.IsNullOrWhiteSpace(_selectedLicenseId)) {
            Dbg($"ResolveSelectedLicenseIdForAction: already in memory={_selectedLicenseId}");
            return true;
        }

        if (!string.IsNullOrWhiteSpace(_trimbleEmail) &&
            SharedConstants.DefaultLicenseIds.Contains(_trimbleEmail, StringComparer.OrdinalIgnoreCase)) {
            _selectedLicenseId = _trimbleEmail;
            SaveSelectedLicenseId(_selectedLicenseId);
            Dbg($"ResolveSelectedLicenseIdForAction: resolved from Trimble registry={_selectedLicenseId}");
            return true;
        }

        var licenseIdsFromFile = FindLicenseIdsByCurrentUserInLicenseFile();
        var savedLicenseId = LoadSelectedLicenseId();

        if (licenseIdsFromFile.Count == 1) {
            _selectedLicenseId = licenseIdsFromFile[0];
            SaveSelectedLicenseId(_selectedLicenseId);
            Dbg($"ResolveSelectedLicenseIdForAction: resolved from file={_selectedLicenseId}");
            return true;
        }

        if (!string.IsNullOrWhiteSpace(savedLicenseId) &&
            licenseIdsFromFile.Contains(savedLicenseId, StringComparer.OrdinalIgnoreCase)) {
            _selectedLicenseId = savedLicenseId;
            Dbg($"ResolveSelectedLicenseIdForAction: resolved from saved among many={_selectedLicenseId}");
            return true;
        }

        if (!string.IsNullOrWhiteSpace(savedLicenseId) && licenseIdsFromFile.Count == 0) {
            _selectedLicenseId = savedLicenseId;
            Dbg($"ResolveSelectedLicenseIdForAction: resolved from saved file={_selectedLicenseId}");
            return true;
        }

        if (!string.IsNullOrWhiteSpace(_trimbleEmail) &&
            !SharedConstants.DefaultLicenseIds.Contains(_trimbleEmail, StringComparer.OrdinalIgnoreCase))
            MessageBox.Show(
                $"Twoje konto Tekla ({_trimbleEmail}) nie znajduje się na liście licencji.\nWybierz właściwe konto ręcznie.",
                "Nieznane konto",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);

        var window = new LicenseAccountSelectionWindow(savedLicenseId);
        var result = window.ShowDialog();

        if (result != true || string.IsNullOrWhiteSpace(window.SelectedLicenseId))
            return false;

        _selectedLicenseId = window.SelectedLicenseId;
        SaveSelectedLicenseId(_selectedLicenseId);
        Dbg($"ResolveSelectedLicenseIdForAction: selected manually={_selectedLicenseId}");
        return true;
    }

    private bool EnsureSelectedLicenseId() {
        if (!string.IsNullOrWhiteSpace(_selectedLicenseId)) {
            Dbg($"EnsureSelectedLicenseId: already in memory={_selectedLicenseId}");
            return true;
        }

        if (!string.IsNullOrWhiteSpace(_trimbleEmail) &&
            SharedConstants.DefaultLicenseIds.Contains(_trimbleEmail, StringComparer.OrdinalIgnoreCase)) {
            _selectedLicenseId = _trimbleEmail;
            SaveSelectedLicenseId(_selectedLicenseId);
            Dbg($"EnsureSelectedLicenseId: resolved from Trimble registry={_selectedLicenseId}");
            return true;
        }

        if (_mode is "readin" or "writeout")
            return ResolveSelectedLicenseIdForAction();

        var licenseIdsFromFile = FindLicenseIdsByCurrentUserInLicenseFile();

        if (licenseIdsFromFile.Count == 1) {
            _selectedLicenseId = licenseIdsFromFile[0];
            SaveSelectedLicenseId(_selectedLicenseId);
            Dbg($"EnsureSelectedLicenseId: resolved from main file={_selectedLicenseId}");
            return true;
        }

        var savedLicenseId = LoadSelectedLicenseId();

        if (!string.IsNullOrWhiteSpace(savedLicenseId) &&
            licenseIdsFromFile.Contains(savedLicenseId, StringComparer.OrdinalIgnoreCase)) {
            _selectedLicenseId = savedLicenseId;
            Dbg($"EnsureSelectedLicenseId: resolved from saved among matches={_selectedLicenseId}");
            return true;
        }

        if (!string.IsNullOrWhiteSpace(_trimbleEmail) &&
            !SharedConstants.DefaultLicenseIds.Contains(_trimbleEmail, StringComparer.OrdinalIgnoreCase))
            MessageBox.Show(
                $"Twoje konto Tekla ({_trimbleEmail}) nie znajduje się na liście licencji.\nWybierz właściwe konto ręcznie.",
                "Nieznane konto",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);

        var window = new LicenseAccountSelectionWindow(savedLicenseId);
        var result = window.ShowDialog();

        if (result != true || string.IsNullOrWhiteSpace(window.SelectedLicenseId))
            return false;

        _selectedLicenseId = window.SelectedLicenseId;
        SaveSelectedLicenseId(_selectedLicenseId);
        Dbg($"EnsureSelectedLicenseId: selected manually={_selectedLicenseId}");
        return true;
    }

    #endregion

    #region Model Sharing Confirmation Helpers

    private static string SafeReadLogToTemp(string logPath) {
        Dbg($"SafeReadLogToTemp: logPath={logPath}");

        if (string.IsNullOrEmpty(logPath) || !File.Exists(logPath)) {
            Dbg("SafeReadLogToTemp: file missing");
            return null;
        }

        try {
            var tempPath = Path.GetTempFileName();

            using var src = new FileStream(
                logPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var dst = new FileStream(
                tempPath, FileMode.Create, FileAccess.Write, FileShare.None);

            src.CopyTo(dst);
            Dbg($"SafeReadLogToTemp: OK temp={tempPath}");
            return tempPath;
        }
        catch (Exception ex) {
            Dbg("SafeReadLogToTemp: EXCEPTION", ex);
            return null;
        }
    }

    private static bool HasConfirmationLine(string logPath, string[] markers, DateTime fromUtc) {
        Dbg(
            $"HasConfirmationLine: ENTER logPath={logPath} markers=[{string.Join(", ", markers)}] fromUtc={fromUtc:O}");

        string[] lines;
        try {
            lines = File.ReadAllLines(logPath);
        }
        catch (Exception ex) {
            Dbg("HasConfirmationLine: ReadAllLines EXCEPTION", ex);
            return false;
        }

        if (fromUtc.Kind != DateTimeKind.Utc)
            fromUtc = fromUtc.ToUniversalTime();

        for (var i = lines.Length - 1; i >= 0; i--) {
            var line = lines[i];

            var markerMatched = markers.Any(marker =>
                !string.IsNullOrWhiteSpace(marker) &&
                line.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0);

            if (!markerMatched)
                continue;

            Dbg($"HasConfirmationLine: candidate line='{line}'");

            var startBracket = line.IndexOf('[');
            var endBracket = line.IndexOf(']', startBracket + 1);

            if (startBracket < 0 || endBracket <= startBracket + 1) {
                Dbg("HasConfirmationLine: candidate without timestamp in brackets -> ACCEPT");
                return true;
            }

            var tsText = line.Substring(startBracket + 1, endBracket - startBracket - 1);

            if (!DateTime.TryParse(
                    tsText,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                    out var tsUtc)) {
                Dbg($"HasConfirmationLine: timestamp parse failed for '{tsText}' -> ACCEPT");
                return true;
            }

            Dbg($"HasConfirmationLine: parsed tsUtc={tsUtc:O}");

            if (tsUtc >= fromUtc) {
                Dbg("HasConfirmationLine: timestamp >= fromUtc -> ACCEPT");
                return true;
            }

            Dbg("HasConfirmationLine: timestamp < fromUtc -> reject candidate");
        }

        Dbg("HasConfirmationLine: no matching line found");
        return false;
    }

    private static async Task<bool> WaitForModelSharingConfirmationAsync(
        string logPath,
        string[] markers,
        DateTime fromUtc,
        int maxWaitSeconds = 300,
        int delaySeconds = 5) {
        Dbg(
            $"WaitForModelSharingConfirmationAsync: ENTER markers=[{string.Join(", ", markers)}] fromUtc={fromUtc:O} maxWait={maxWaitSeconds}s delay={delaySeconds}s");

        var start = DateTime.UtcNow;
        var timeout = TimeSpan.FromSeconds(maxWaitSeconds);
        var delay = TimeSpan.FromSeconds(delaySeconds);

        if (fromUtc.Kind != DateTimeKind.Utc)
            fromUtc = fromUtc.ToUniversalTime();

        while (DateTime.UtcNow - start < timeout) {
            try {
                var tempPath = SafeReadLogToTemp(logPath);

                if (tempPath != null) {
                    bool found;
                    try {
                        found = HasConfirmationLine(tempPath, markers, fromUtc);
                        Dbg($"WaitForModelSharingConfirmationAsync: HasConfirmationLine={found}");
                    }
                    finally {
                        try {
                            File.Delete(tempPath);
                        }
                        catch (Exception ex) {
                            Dbg("WaitForModelSharingConfirmationAsync: temp delete EX", ex);
                        }
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

    private async Task WaitAndFinalizeReadInAsync(
        string logPath,
        SharedLicenseInfo info,
        string userName) {
        Dbg($"WaitAndFinalizeReadInAsync: ENTER logPath={logPath}");

        const int waitMaxSecondsPerCycle = 60;

        try {
            await Application.Current.Dispatcher.InvokeAsync(() => {
                AppendLog("READ IN: rozpoczęto...");
                ReadIn();
            });
        }
        catch (Exception ex) {
            Dbg("WaitAndFinalizeReadInAsync: ReadIn invoke EX", ex);
            Interlocked.Exchange(ref _readInStarted, 0);
            return;
        }

        var readInMarkers = new[] {
            "Read-in result: OK.",
            "Read in result: OK.",
            "Read-in result: OK",
            "Read in result: OK",
            "Read-in OK",
            "Read in OK"
        };

        var fromUtc = DateTime.UtcNow.AddSeconds(-10);

        while (true) {
            bool ok;

            try {
                ok = await WaitForModelSharingConfirmationAsync(
                    logPath,
                    readInMarkers,
                    fromUtc,
                    waitMaxSecondsPerCycle
                ).ConfigureAwait(false);
            }
            catch (Exception ex) {
                Dbg("WaitAndFinalizeReadInAsync: EXCEPTION while waiting", ex);
                ok = false;
            }

            Dbg($"WaitAndFinalizeReadInAsync: WAIT RESULT ok={ok}");

            if (ok) {
                var now = DateTime.Now;
                SharedLicenseManager.ReadIn(info, userName, now);
                SharedLicenseFileService.Save(SharedConstants.LicenseFilePath, info);

                await Application.Current.Dispatcher.InvokeAsync(() => {
                    if (_mode == "readin" && !IsLoaded)
                        return;

                    AppendLog($"READ IN - {userName} - {now.ToString(SharedConstants.DateFormat)}");

                    if (_mode == "readin")
                        Close();
                });

                Interlocked.Exchange(ref _readInStarted, 0);
                Dbg("WaitAndFinalizeReadInAsync: SUCCESS");
                return;
            }

            await Application.Current.Dispatcher.InvokeAsync(() => {
                AppendLog("READ IN: brak potwierdzenia po 60 sekundach.");
            });

            var shouldContinueWaiting = await AskUserIfReadInSucceededAsync().ConfigureAwait(false);

            if (!shouldContinueWaiting) {
                Dbg("WaitAndFinalizeReadInAsync: user selected NO -> stop without save");

                await Application.Current.Dispatcher.InvokeAsync(() => {
                    AppendLog("READ IN: anulowano zapis do bazy danych.");
                    if (_mode == "readin")
                        Close();
                });

                Interlocked.Exchange(ref _readInStarted, 0);
                return;
            }

            Dbg("WaitAndFinalizeReadInAsync: user selected YES -> continue waiting");
        }
    }

    private async Task WaitAndFinalizeWriteOutAsync(
        string logPath,
        SharedLicenseInfo info,
        string userName) {
        Dbg($"WaitAndFinalizeWriteOutAsync: ENTER logPath={logPath}");

        const int maxAttempts = 3;
        const int attemptDelaySeconds = 10;
        const int waitMaxSecondsPerAttempt = 5;

        var writeOutMarkers = new[] {
            "WriteOut OK",
            "Write out OK",
            "Write-out OK",
            "WriteOut result: OK",
            "Write out result: OK",
            "Write-out result: OK"
        };

        for (var attempt = 1; attempt <= maxAttempts; attempt++) {
            Dbg($"WaitAndFinalizeWriteOutAsync: ATTEMPT {attempt}/{maxAttempts}");

            try {
                var currentAttempt = attempt;
                await Application.Current.Dispatcher.InvokeAsync(() => {
                    AppendLog($"WRITE OUT: próba {currentAttempt}/{maxAttempts}...");
                    WriteOut();
                });
            }
            catch (Exception ex) {
                Dbg("WaitAndFinalizeWriteOutAsync: WriteOut invoke EX", ex);
            }

            var attemptFromUtc = DateTime.UtcNow.AddSeconds(-10);

            bool ok;
            try {
                ok = await WaitForModelSharingConfirmationAsync(
                    logPath,
                    writeOutMarkers,
                    attemptFromUtc,
                    waitMaxSecondsPerAttempt
                ).ConfigureAwait(false);
            }
            catch (Exception ex) {
                Dbg("WaitAndFinalizeWriteOutAsync: EXCEPTION while waiting", ex);
                ok = false;
            }

            Dbg($"WaitAndFinalizeWriteOutAsync: WAIT RESULT ok={ok} attempt={attempt}");

            if (ok) {
                var now = DateTime.Now;
                SharedLicenseManager.WriteOut(info, userName, now);
                SharedLicenseFileService.Save(SharedConstants.LicenseFilePath, info);

                await Application.Current.Dispatcher.InvokeAsync(() => {
                    if (_mode == "writeout" && !IsLoaded)
                        return;

                    AppendLog($"WRITE OUT - {userName} - {now.ToString(SharedConstants.DateFormat)}");

                    if (_mode == "writeout")
                        Close();
                });

                Interlocked.Exchange(ref _writeOutStarted, 0);
                Dbg("WaitAndFinalizeWriteOutAsync: SUCCESS");
                return;
            }

            if (attempt < maxAttempts)
                await Task.Delay(TimeSpan.FromSeconds(attemptDelaySeconds)).ConfigureAwait(false);
        }

        Dbg("WaitAndFinalizeWriteOutAsync: ALL ATTEMPTS FAILED");

        await Application.Current.Dispatcher.InvokeAsync(() => {
            AppendLog($"WRITE OUT: nie udało się po {maxAttempts} próbach.");
            if (_mode == "writeout")
                Close();
        });

        Interlocked.Exchange(ref _writeOutStarted, 0);
    }

    private static async Task<bool> AskUserIfReadInSucceededAsync() {
        try {
            var result = await Application.Current.Dispatcher.InvokeAsync(() => {
                var window = new ReadInConfirmationWindow();

                var mainWindow = Application.Current?.MainWindow;
                if (mainWindow != null &&
                    mainWindow != window &&
                    mainWindow.IsVisible)
                    window.Owner = mainWindow;

                var dialogResult = window.ShowDialog();
                var confirmed = dialogResult == true && window.UserConfirmed;

                Dbg($"AskUserIfReadInSucceededAsync: dialogResult={dialogResult} confirmed={confirmed}");
                return confirmed;
            });

            return result;
        }
        catch (Exception ex) {
            Dbg("AskUserIfReadInSucceededAsync: EXCEPTION", ex);
            return false;
        }
    }

    #endregion

    #region Auto Login And Lifecycle

    private void StartAutoLoginWatcher() {
        try {
            if (Interlocked.Exchange(ref _autoLoginWatcherStarted, 1) == 1) {
                Dbg("StartAutoLoginWatcher: already started -> return");
                return;
            }

            Dbg(
                $"StartAutoLoginWatcher: ENTER mode={_mode} autoLoginDone={_autoLoginDone} logoutDone={_logoutDone}");

            try {
                ClearLog();
                AppendLog("Czekam na zapis LOG IN...");
                Dbg("StartAutoLoginWatcher: UI cleared + message written.");

                if (!EnsureSelectedLicenseId()) {
                    Dbg("StartAutoLoginWatcher: no selected license -> return");
                    return;
                }

                SaveSelectedLicenseId(_selectedLicenseId);

                var loginTime = DateTime.Now;
                var userName = Environment.UserName;

                var info = SharedLicenseFileService.LoadOrCreate(
                    SharedConstants.LicenseFilePath, _selectedLicenseId);
                Dbg("StartAutoLoginWatcher: LoadOrCreate OK");

                SharedLicenseManager.Logout(info, userName);
                Dbg("StartAutoLoginWatcher: removed stale login before adding fresh one");

                SharedLicenseManager.Login(info, userName, loginTime);
                Dbg("StartAutoLoginWatcher: SharedLicenseManager.Login OK");

                SharedLicenseFileService.Save(SharedConstants.LicenseFilePath, info);
                Dbg($"StartAutoLoginWatcher: Save OK -> {SharedConstants.LicenseFilePath}");

                _autoLoginLicenseId = _selectedLicenseId;
                _autoLoginDone = true;
                AppendLog(
                    $"LOG IN - {userName} - {loginTime.ToString(SharedConstants.DateFormat)}");
                Dbg(
                    $"StartAutoLoginWatcher: _autoLoginDone=true _autoLoginLicenseId={_autoLoginLicenseId} + UI log appended");
            }
            catch (Exception ex) {
                Dbg("StartAutoLoginWatcher: OUTER EXCEPTION", ex);
            }
        }
        catch {
            // ignored
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
        try {
            Dbg($"HandleModelClosed: ENTER mode={_mode} autoLoginDone={_autoLoginDone}");

            try {
                if (_mode == "autologin" &&
                    _autoLoginDone &&
                    Interlocked.CompareExchange(ref _logoutDone, 1, 0) == 0) {
                    Dbg("HandleModelClosed: calling RemoveCurrentUserLogin()");
                    RemoveCurrentUserLogin();
                    Dbg("HandleModelClosed: RemoveCurrentUserLogin DONE");
                }
            }
            catch (Exception ex) {
                Dbg("HandleModelClosed: EXCEPTION during logout", ex);
                Interlocked.Exchange(ref _logoutDone, 0);
            }

            try {
                Dbg("HandleModelClosed: Dispatcher.Invoke(Close)...");
                Application.Current.Dispatcher.Invoke(Close);
            }
            catch (Exception ex) {
                Dbg("HandleModelClosed: EXCEPTION while closing", ex);
            }
        }
        catch (Exception ex) {
            Dbg("HandleModelClosed: OUTER EXCEPTION", ex);
        }
    }

    private void RemoveCurrentUserLogin() {
        Dbg("RemoveCurrentUserLogin: ENTER");

        try {
            var userName = Environment.UserName;
            var licenseIdToRemove = _autoLoginLicenseId;

            if (string.IsNullOrWhiteSpace(userName)) {
                Dbg("RemoveCurrentUserLogin: userName is empty -> nothing to remove");
                return;
            }

            if (string.IsNullOrWhiteSpace(licenseIdToRemove)) {
                Dbg("RemoveCurrentUserLogin: autoLogin license is empty -> nothing to remove");
                return;
            }

            var info = SharedLicenseFileService.LoadOrCreate(
                SharedConstants.LicenseFilePath,
                licenseIdToRemove);

            if (info == null) {
                Dbg($"RemoveCurrentUserLogin: LoadOrCreate returned null for licenseId={licenseIdToRemove}");
                return;
            }

            var userExists = info.Logins.Any(login =>
                login != null &&
                !string.IsNullOrWhiteSpace(login.User) &&
                string.Equals(login.User, userName, StringComparison.OrdinalIgnoreCase));

            if (!userExists) {
                Dbg($"RemoveCurrentUserLogin: user={userName} not found in licenseId={licenseIdToRemove}");
                return;
            }

            SharedLicenseManager.Logout(info, userName);
            Dbg($"RemoveCurrentUserLogin: removed user={userName} from licenseId={licenseIdToRemove}");

            SharedLicenseFileService.Save(SharedConstants.LicenseFilePath, info);
            Dbg($"RemoveCurrentUserLogin: saved licenseId={licenseIdToRemove}");

            DeleteSelectedLicenseId();
            _selectedLicenseId = null;
            _autoLoginLicenseId = null;
            Dbg("RemoveCurrentUserLogin: local license data cleared");
        }
        catch (Exception ex) {
            Dbg("RemoveCurrentUserLogin: EXCEPTION", ex);
            throw;
        }
    }

    protected override void OnClosed(EventArgs e) {
        try {
            Dbg($"OnClosed: ENTER mode={_mode} autoLoginDone={_autoLoginDone} logoutDone={_logoutDone}");

            try {
                if (_mode == "autologin" &&
                    _autoLoginDone &&
                    Interlocked.CompareExchange(ref _logoutDone, 1, 0) == 0) {
                    Dbg("OnClosed: autologin & logout not done -> RemoveCurrentUserLogin()");
                    RemoveCurrentUserLogin();
                    Dbg("OnClosed: logout done");
                }
            }
            catch (Exception ex) {
                Dbg("OnClosed: EXCEPTION during RemoveCurrentUserLogin", ex);
                Interlocked.Exchange(ref _logoutDone, 0);
            }

            try {
                if (_events != null) {
                    _events.ModelUnloading -= OnModelUnloading;
                    _events.TeklaStructuresExit -= OnTeklaStructuresExit;
                    _events.UnRegister();
                    _events = null;
                }
            }
            catch (Exception ex) {
                Dbg("OnClosed: EXCEPTION during events unregister", ex);
            }

            base.OnClosed(e);
            Dbg("OnClosed: EXIT");
        }
        catch (Exception ex) {
            Dbg("OnClosed: OUTER EXCEPTION", ex);
        }
    }

    #endregion

    #region Scripts

    private static void ReadIn() {
        TSMO.Operation.RunMacro(
            @"Z:\000_PMJ\Tekla\HFT_SharedTool\SharedTool\ReadIn.cs");
    }

    private static void WriteOut() {
        TSMO.Operation.RunMacro(
            @"Z:\000_PMJ\Tekla\HFT_SharedTool\SharedTool\WriteOut.cs");
    }

    #endregion

    #region Title Bar

    private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) {
        DragMove();
    }

    private void ThemeToggle_Click(object sender, RoutedEventArgs e) {
        try {
            ThemeService.Toggle();
            UpdateThemeToggleIconColor();

            if (_mode is "standalone" or "check" or "ckeck")
                BtnCheck_Click();
        }
        catch (Exception ex) {
            Dbg("ThemeToggle_Click: EXCEPTION", ex);
            MessageBox.Show($"Błąd zmiany motywu: {ex.Message}");
        }
    }

    private void UpdateThemeToggleIconColor() {
        try {
            if (ThemeToggleIcon == null)
                return;

            if (ThemeService.IsDark)
                ThemeToggleIcon.Foreground =
                    Application.Current.FindResource("TextSecondaryBrush") as Brush;
            else
                ThemeToggleIcon.Foreground = new SolidColorBrush(Color.FromRgb(0x00, 0x5F, 0xB8));
        }
        catch (Exception ex) {
            Dbg("UpdateThemeToggleIconColor: EXCEPTION", ex);
        }
    }

    private void BtnRefresh_Click(object sender, RoutedEventArgs e) {
        try {
            if (RefreshIconRotate != null) {
                var animation = new DoubleAnimation {
                    From = 0,
                    To = 360,
                    Duration = TimeSpan.FromMilliseconds(500)
                };

                RefreshIconRotate.BeginAnimation(RotateTransform.AngleProperty, animation);
            }

            BtnCheck_Click();
        }
        catch (Exception ex) {
            Dbg("BtnRefresh_Click: EXCEPTION", ex);
            MessageBox.Show($"Błąd odświeżania: {ex.Message}");
        }
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e) {
        Close();
    }

    #endregion
}