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

            basePath = basePath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            logPath = Path.Combine(basePath, "logs", "modelsharing.log");
            return true;
        }
        catch {
            logPath = null;
            return false;
        }
    }

    #endregion

    #region Constants

    private static readonly string SelectedLicenseFilePath =
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "HFT_SharedTool",
            $"selected_license_{Environment.UserName}.txt");

    private const string MinTimeBetweenFilePath =
        @"Z:\000_PMJ\Tekla\HFT_SharedTool\HFT_SharingTool_MinTimeBetween.txt";

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

        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
            Dbg("UnhandledException", e.ExceptionObject as Exception);

        TaskScheduler.UnobservedTaskException += (_, e) => {
            Dbg("UnobservedTaskException", e.Exception);
            e.SetObserved();
        };

        Application.Current.DispatcherUnhandledException += (_, e) =>
            Dbg("DispatcherUnhandledException", e.Exception);

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

        RegisterModelEvents();

        if (_mode == "autologin") {
            HideWindow();
            await TryHandleAutoLoginForCurrentModelAsync();
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

        var modelName = modelInfo.ModelName.Replace(".db1", "");
        ModelDrawingLabel.Text = $"Połączono z {modelName}";
        Dbg($"InitializeAfterLoadAsync: modelName={modelName}");

        switch (_mode) {
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

    private void RegisterModelEvents() {
        try {
            if (_events != null) {
                Dbg("RegisterModelEvents: already registered");
                return;
            }

            _events = new TSM.Events();
            _events.ModelLoad += OnModelLoad;
            _events.ModelUnloading += OnModelUnloading;
            _events.TeklaStructuresExit += OnTeklaStructuresExit;
            _events.Register();

            Dbg("RegisterModelEvents: registered");
        }
        catch (Exception ex) {
            Dbg("RegisterModelEvents: EXCEPTION", ex);
        }
    }

    private async Task TryHandleAutoLoginForCurrentModelAsync() {
        try {
            if (_mode != "autologin") {
                Dbg("TryHandleAutoLoginForCurrentModelAsync: mode is not autologin");
                return;
            }

            if (_autoLoginDone) {
                Dbg("TryHandleAutoLoginForCurrentModelAsync: already done");
                return;
            }

            await Dispatcher.InvokeAsync(() => {
                try {
                    var model = new TSM.Model();

                    if (!model.GetConnectionStatus()) {
                        Dbg("TryHandleAutoLoginForCurrentModelAsync: no model connection");
                        return;
                    }

                    var info = model.GetInfo();

                    Dbg(
                        $"TryHandleAutoLoginForCurrentModelAsync: modelName={info.ModelName}, sharedModel={info.SharedModel}");

                    if (!info.SharedModel) {
                        Dbg("TryHandleAutoLoginForCurrentModelAsync: current model is not sharing");
                        return;
                    }

                    StartAutoLoginWatcher();
                }
                catch (Exception ex) {
                    Dbg("TryHandleAutoLoginForCurrentModelAsync: dispatcher EXCEPTION", ex);
                }
            });
        }
        catch (Exception ex) {
            Dbg("TryHandleAutoLoginForCurrentModelAsync: OUTER EXCEPTION", ex);
        }
    }

    private void OnModelLoad() {
        Dbg("OnModelLoad: EVENT");

        try {
            if (_mode != "autologin") {
                Dbg("OnModelLoad: mode is not autologin -> ignore");
                return;
            }

            if (_autoLoginDone) {
                Dbg("OnModelLoad: auto login already done -> ignore");
                return;
            }

            RunOnUi(async () => { await TryHandleAutoLoginForCurrentModelAsync(); });
        }
        catch (Exception ex) {
            Dbg("OnModelLoad: EXCEPTION", ex);
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
                    File.AppendAllText(path,
                        @"INNER EX: " + ex.InnerException.GetType().FullName + Environment.NewLine);
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

    #region Min Time Between Helpers

    private static readonly object MinTimeBetweenLock = new();

    // Zapisuje minimalny czas między operacjami różnych użytkowników (READ IN / WRITE OUT).
    private static void UpdateMinTimeBetweenFile(
        string currentAction,
        SharedLicenseInfo info,
        string currentUser,
        DateTime currentTime) {
        try {
            if (string.IsNullOrWhiteSpace(currentAction) ||
                info == null ||
                string.IsNullOrWhiteSpace(currentUser))
                return;

            // Wyznacz poprzednią operację (READ IN lub WRITE OUT, nowsza wygrywa)
            var hasRead = info.ReadTime.HasValue && !string.IsNullOrWhiteSpace(info.ReadUser);
            var hasWrite = info.WriteTime.HasValue && !string.IsNullOrWhiteSpace(info.WriteUser);

            if (!hasRead && !hasWrite)
                return;

            string previousUser;
            DateTime previousTime;

            if (hasRead && (!hasWrite || info.ReadTime.Value >= info.WriteTime.Value)) {
                previousUser = info.ReadUser;
                previousTime = info.ReadTime.Value;
            }
            else {
                previousUser = info.WriteUser;
                previousTime = info.WriteTime.Value;
            }

            // Pomiń, jeśli ten sam użytkownik
            if (string.Equals(previousUser, currentUser, StringComparison.OrdinalIgnoreCase)) {
                Dbg($"UpdateMinTimeBetweenFile: skip same user={currentUser}");
                return;
            }

            var diff = currentTime - previousTime;

            if (diff < TimeSpan.Zero) {
                Dbg($"UpdateMinTimeBetweenFile: skip negative diff={diff}");
                return;
            }

            lock (MinTimeBetweenLock) {
                var values = LoadMinTimeBetweenValues();

                if (!values.TryGetValue(currentAction, out var existing) || diff < existing) {
                    values[currentAction] = diff;
                    SaveMinTimeBetweenValues(values);
                    Dbg(
                        $"UpdateMinTimeBetweenFile: updated action={currentAction} diff={diff} prev={previousUser} curr={currentUser}");
                }
                else
                    Dbg(
                        $"UpdateMinTimeBetweenFile: not updated action={currentAction} diff={diff} existing={existing}");
            }
        }
        catch (Exception ex) {
            Dbg("UpdateMinTimeBetweenFile: EXCEPTION", ex);
        }
    }

    private static Dictionary<string, TimeSpan> LoadMinTimeBetweenValues() {
        var values = new Dictionary<string, TimeSpan>(StringComparer.OrdinalIgnoreCase);

        try {
            if (!File.Exists(MinTimeBetweenFilePath))
                return values;

            foreach (var raw in File.ReadAllLines(MinTimeBetweenFilePath)) {
                var line = raw?.Trim();
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                if (TryParseMinTimeLine(line, out var action, out var value))
                    values[action] = value;
            }

            Dbg($"LoadMinTimeBetweenValues: loaded count={values.Count}");
        }
        catch (Exception ex) {
            Dbg("LoadMinTimeBetweenValues: EXCEPTION", ex);
        }

        return values;
    }

    private static void SaveMinTimeBetweenValues(Dictionary<string, TimeSpan> values) {
        try {
            var dir = Path.GetDirectoryName(MinTimeBetweenFilePath);
            if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            // Zapisz tylko znane akcje w ustalonej kolejności
            var lines = new[] { "READ IN", "WRITE OUT" }
                .Where(values.ContainsKey)
                .Select(action => $"{action} {FormatMinTimeBetween(values[action])}");

            File.WriteAllLines(MinTimeBetweenFilePath, lines);
            Dbg($"SaveMinTimeBetweenValues: saved path={MinTimeBetweenFilePath}");
        }
        catch (Exception ex) {
            Dbg("SaveMinTimeBetweenValues: EXCEPTION", ex);
        }
    }

    private static bool TryParseMinTimeLine(string line, out string action, out TimeSpan value) {
        action = null;
        value = default;

        try {
            const string readInPrefix = "READ IN ";
            const string writeOutPrefix = "WRITE OUT ";

            if (line.StartsWith(readInPrefix, StringComparison.OrdinalIgnoreCase)) {
                action = "READ IN";
                return TryParseMinTimeValue(line.Substring(readInPrefix.Length), out value);
            }

            if (line.StartsWith(writeOutPrefix, StringComparison.OrdinalIgnoreCase)) {
                action = "WRITE OUT";
                return TryParseMinTimeValue(line.Substring(writeOutPrefix.Length), out value);
            }

            return false;
        }
        catch (Exception ex) {
            Dbg("TryParseMinTimeLine: EXCEPTION", ex);
            action = null;
            value = default;
            return false;
        }
    }

    private static bool TryParseMinTimeValue(string text, out TimeSpan value) {
        value = default;

        if (string.IsNullOrWhiteSpace(text))
            return false;

        var normalized = text.Trim().TrimEnd('h', 'H').Trim();
        var parts = normalized.Split(':');

        if (parts.Length != 2 ||
            !int.TryParse(parts[0], out var hours) ||
            !int.TryParse(parts[1], out var minutes) ||
            hours < 0 || minutes < 0 || minutes > 59)
            return false;

        value = new TimeSpan(hours, minutes, 0);
        return true;
    }

    private static string FormatMinTimeBetween(TimeSpan value) {
        if (value < TimeSpan.Zero) value = TimeSpan.Zero;
        return $"{(int)value.TotalHours:D2}:{value.Minutes:D2}h";
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

                var isCurrentUserLoggedHere = loggedUsers.Any(u =>
                    string.Equals(u, currentUser, StringComparison.OrdinalIgnoreCase));

                var isLastActionByCurrentUser =
                    !string.IsNullOrWhiteSpace(actionUser) &&
                    string.Equals(actionUser, currentUser, StringComparison.OrdinalIgnoreCase);

                var isTakenByAnotherUser =
                    isCurrentUserLoggedHere &&
                    !string.IsNullOrWhiteSpace(actionUser) &&
                    !isLastActionByCurrentUser;

                var tooltipText = info.NextUsable.HasValue
                    ? $"Dostępna od: {info.NextUsable.Value:yyyy-MM-dd HH:mm}"
                    : "Dostępna teraz";

                var isDark = ThemeService.IsDark;

                var dotOk = isDark ? Color.FromRgb(0x6E, 0xC9, 0x6E) : Color.FromRgb(0x16, 0xA3, 0x4A);
                var dotActive = isDark ? Color.FromRgb(0x7C, 0xFF, 0x9A) : Color.FromRgb(0x63, 0xF5, 0x87);
                var dotErr = isDark ? Color.FromRgb(0xE2, 0x4B, 0x4A) : Color.FromRgb(0xDC, 0x26, 0x26);
                var dotNotActive = isDark ? Color.FromRgb(0xFF, 0x7A, 0x7A) : Color.FromRgb(0xFF, 0x5A, 0x5A);

                var infoFg = new SolidColorBrush(isDark
                    ? Color.FromArgb(0xAA, 0xB8, 0xB8, 0xB8)
                    : Color.FromRgb(0x6B, 0x72, 0x80));

                var row = new Grid { Margin = new Thickness(0, 4, 0, 4) };
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(32) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(110) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                var dotColor = dotErr;
                var shouldGlow = false;

                if (isTakenByAnotherUser) {
                    dotColor = dotNotActive;
                    shouldGlow = true;
                }
                else if (isCurrentUserLoggedHere && isLastActionByCurrentUser) {
                    dotColor = dotActive;
                    shouldGlow = true;
                }
                else if (isUsable)
                    dotColor = dotOk;
                else
                    dotColor = dotErr;

                var dot = new Ellipse {
                    Width = 10,
                    Height = 10,
                    Fill = new SolidColorBrush(dotColor),
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    ToolTip = tooltipText
                };

                if (shouldGlow)
                    dot.Effect = new DropShadowEffect {
                        Color = dotColor,
                        BlurRadius = 12,
                        ShadowDepth = 0,
                        Opacity = 1
                    };

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
                    foreach (var user in loggedUsers)
                        loggedUsersPanel.Children.Add(MakeBadge(
                            user, false, string.Equals(user, currentUser, StringComparison.OrdinalIgnoreCase)));

                Grid.SetColumn(loggedUsersPanel, 2);
                row.Children.Add(loggedUsersPanel);

                if (actionTime.HasValue && !string.IsNullOrWhiteSpace(actionUser)) {
                    var isActionByCurrentUser =
                        string.Equals(actionUser, currentUser, StringComparison.OrdinalIgnoreCase);

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
        var diff = DateTime.Now - dateTime;
        if (diff < TimeSpan.Zero) diff = TimeSpan.Zero;

        return diff.TotalDays switch {
            >= 2 => $"{(int)diff.TotalDays} dni temu",
            >= 1 => "1 dzień temu",
            _ => $"{(int)diff.TotalHours}:{diff.Minutes:D2}h temu"
        };
    }

    private static Border MakeBadge(string text, bool isFree, bool isCurrentUser) {
        Color bg, fg;

        if (isFree) {
            bg = Color.FromArgb(0x24, 0x90, 0x90, 0x90);
            fg = Color.FromArgb(0xCC, 0xA0, 0xA0, 0xA0);
        }
        else if (isCurrentUser) {
            bg = ThemeService.IsDark
                ? Color.FromArgb(0x32, 0x4C, 0xD9, 0x6B)
                : Color.FromArgb(0x2C, 0x22, 0xC5, 0x5E);
            fg = ThemeService.IsDark
                ? Color.FromRgb(0x7C, 0xFF, 0x9A)
                : Color.FromRgb(0x16, 0xA3, 0x4A);
        }
        else {
            bg = ThemeService.IsDark
                ? Color.FromArgb(0x30, 0x3D, 0x7E, 0xA6)
                : Color.FromArgb(0x22, 0x00, 0x5F, 0xB8);
            fg = ThemeService.IsDark
                ? Color.FromRgb(0x7D, 0xD3, 0xFC)
                : Color.FromRgb(0x00, 0x5F, 0xB8);
        }

        return new Border {
            Background = new SolidColorBrush(bg),
            CornerRadius = new CornerRadius(10),
            Padding = new Thickness(7, 1, 7, 2),
            Margin = new Thickness(0, 0, 6, 0),
            VerticalAlignment = VerticalAlignment.Center,
            Child = new TextBlock {
                Text = text,
                FontSize = 11,
                Foreground = new SolidColorBrush(fg)
            }
        };
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

            var info = SharedLicenseFileService.LoadOrCreate(SharedConstants.LicenseFilePath, _selectedLicenseId);
            _ = WaitAndFinalizeReadInAsync(modelSharingLogPath, info, Environment.UserName);
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

            var info = SharedLicenseFileService.LoadOrCreate(SharedConstants.LicenseFilePath, _selectedLicenseId);
            _ = WaitAndFinalizeWriteOutAsync(modelSharingLogPath, info, Environment.UserName);
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

        var now = DateTime.Now;
        var currentUser = Environment.UserName;

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
            var result = infos
                .Where(info => info?.Logins != null && !string.IsNullOrWhiteSpace(info.LicenseId))
                .Where(info => info.Logins.Any(login =>
                    login != null &&
                    !string.IsNullOrWhiteSpace(login.User) &&
                    string.Equals(login.User, currentUser, StringComparison.OrdinalIgnoreCase)))
                .Select(info => info.LicenseId)
                .ToList();

            Dbg($"FindLicenseIdsByCurrentUserInLicenseFile: found {result.Count} license(s) for user={currentUser}");
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

    private static bool HasConfirmationLine(string logPath, string[] markers, DateTime fromUtc) {
        Dbg($"HasConfirmationLine: ENTER logPath={logPath} markers=[{string.Join(", ", markers)}] fromUtc={fromUtc:O}");

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

            var markerMatched = markers.Any(m =>
                !string.IsNullOrWhiteSpace(m) &&
                line.IndexOf(m, StringComparison.OrdinalIgnoreCase) >= 0);

            if (!markerMatched)
                continue;

            Dbg($"HasConfirmationLine: candidate line='{line}'");

            var startBracket = line.IndexOf('[');
            var endBracket = line.IndexOf(']', startBracket + 1);

            if (startBracket < 0 || endBracket <= startBracket + 1) {
                Dbg("HasConfirmationLine: no timestamp brackets -> ACCEPT");
                return true;
            }

            var tsText = line.Substring(startBracket + 1, endBracket - startBracket - 1);

            if (!DateTime.TryParse(
                    tsText,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                    out var tsUtc)) {
                Dbg($"HasConfirmationLine: timestamp parse failed '{tsText}' -> ACCEPT");
                return true;
            }

            Dbg($"HasConfirmationLine: tsUtc={tsUtc:O}");

            if (tsUtc >= fromUtc) {
                Dbg("HasConfirmationLine: timestamp >= fromUtc -> ACCEPT");
                return true;
            }

            Dbg("HasConfirmationLine: timestamp < fromUtc -> reject");
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

    private async Task WaitAndFinalizeReadInAsync(string logPath, SharedLicenseInfo info, string userName) {
        Dbg($"WaitAndFinalizeReadInAsync: ENTER logPath={logPath}");

        const int waitMaxSecondsPerCycle = 60;

        var readInMarkers = new[] {
            "Read-in result: OK.",
            "Read in result: OK.",
            "Read-in result: OK",
            "Read in result: OK",
            "Read-in OK",
            "Read in OK"
        };

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

        var fromUtc = DateTime.UtcNow.AddSeconds(-10);

        while (true) {
            bool ok;
            try {
                ok = await WaitForModelSharingConfirmationAsync(
                    logPath, readInMarkers, fromUtc, waitMaxSecondsPerCycle).ConfigureAwait(false);
            }
            catch (Exception ex) {
                Dbg("WaitAndFinalizeReadInAsync: EXCEPTION while waiting", ex);
                ok = false;
            }

            Dbg($"WaitAndFinalizeReadInAsync: WAIT RESULT ok={ok}");

            if (ok) {
                var now = DateTime.Now;
                UpdateMinTimeBetweenFile("READ IN", info, userName, now);
                SharedLicenseManager.ReadIn(info, userName, now);
                SharedLicenseFileService.Save(SharedConstants.LicenseFilePath, info);

                await Application.Current.Dispatcher.InvokeAsync(() => {
                    if (_mode == "readin" && !IsLoaded) return;
                    AppendLog($"READ IN - {userName} - {now.ToString(SharedConstants.DateFormat)}");
                    if (_mode == "readin") Close();
                });

                Interlocked.Exchange(ref _readInStarted, 0);
                Dbg("WaitAndFinalizeReadInAsync: SUCCESS");
                return;
            }

            await Application.Current.Dispatcher.InvokeAsync(() =>
                AppendLog("READ IN: brak potwierdzenia po 60 sekundach."));

            var shouldContinue = await AskUserIfReadInSucceededAsync().ConfigureAwait(false);

            if (!shouldContinue) {
                Dbg("WaitAndFinalizeReadInAsync: user selected NO -> stop without save");
                await Application.Current.Dispatcher.InvokeAsync(() => {
                    AppendLog("READ IN: anulowano zapis do bazy danych.");
                    if (_mode == "readin") Close();
                });

                Interlocked.Exchange(ref _readInStarted, 0);
                return;
            }

            Dbg("WaitAndFinalizeReadInAsync: user selected YES -> continue waiting");
        }
    }

    private async Task WaitAndFinalizeWriteOutAsync(string logPath, SharedLicenseInfo info, string userName) {
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
                var current = attempt;
                await Application.Current.Dispatcher.InvokeAsync(() => {
                    AppendLog($"WRITE OUT: próba {current}/{maxAttempts}...");
                    WriteOut();
                });
            }
            catch (Exception ex) {
                Dbg("WaitAndFinalizeWriteOutAsync: WriteOut invoke EX", ex);
            }

            bool ok;
            try {
                ok = await WaitForModelSharingConfirmationAsync(
                        logPath, writeOutMarkers, DateTime.UtcNow.AddSeconds(-10), waitMaxSecondsPerAttempt)
                    .ConfigureAwait(false);
            }
            catch (Exception ex) {
                Dbg("WaitAndFinalizeWriteOutAsync: EXCEPTION while waiting", ex);
                ok = false;
            }

            Dbg($"WaitAndFinalizeWriteOutAsync: WAIT RESULT ok={ok} attempt={attempt}");

            if (ok) {
                var now = DateTime.Now;
                UpdateMinTimeBetweenFile("WRITE OUT", info, userName, now);
                SharedLicenseManager.WriteOut(info, userName, now);
                SharedLicenseFileService.Save(SharedConstants.LicenseFilePath, info);

                await Application.Current.Dispatcher.InvokeAsync(() => {
                    if (_mode == "writeout" && !IsLoaded) return;
                    AppendLog($"WRITE OUT - {userName} - {now.ToString(SharedConstants.DateFormat)}");
                    if (_mode == "writeout") Close();
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
            if (_mode == "writeout") Close();
        });

        Interlocked.Exchange(ref _writeOutStarted, 0);
    }

    private static async Task<bool> AskUserIfReadInSucceededAsync() {
        try {
            return await Application.Current.Dispatcher.InvokeAsync(() => {
                var window = new ReadInConfirmationWindow();

                var main = Application.Current?.MainWindow;
                if (main != null && main != window && main.IsVisible)
                    window.Owner = main;

                var dialogResult = window.ShowDialog();
                var confirmed = dialogResult == true && window.UserConfirmed;

                Dbg($"AskUserIfReadInSucceededAsync: dialogResult={dialogResult} confirmed={confirmed}");
                return confirmed;
            });
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
            if (_mode != "autologin") {
                Dbg("StartAutoLoginWatcher: mode is not autologin -> return");
                return;
            }

            if (_autoLoginDone) {
                Dbg("StartAutoLoginWatcher: already logged in -> return");
                return;
            }

            if (Interlocked.Exchange(ref _autoLoginWatcherStarted, 1) == 1) {
                Dbg("StartAutoLoginWatcher: already started -> return");
                return;
            }

            Dbg($"StartAutoLoginWatcher: ENTER mode={_mode} autoLoginDone={_autoLoginDone} logoutDone={_logoutDone}");

            try {
                ClearLog();
                AppendLog("Wykryto model sharing. Zapisuję LOG IN...");

                if (!EnsureSelectedLicenseId()) {
                    Dbg("StartAutoLoginWatcher: no selected license -> return");
                    Interlocked.Exchange(ref _autoLoginWatcherStarted, 0);
                    return;
                }

                SaveSelectedLicenseId(_selectedLicenseId);

                var loginTime = DateTime.Now;
                var userName = Environment.UserName;

                var info = SharedLicenseFileService.LoadOrCreate(SharedConstants.LicenseFilePath, _selectedLicenseId);

                SharedLicenseManager.Logout(info, userName);
                SharedLicenseManager.Login(info, userName, loginTime);
                SharedLicenseFileService.Save(SharedConstants.LicenseFilePath, info);
                Dbg($"StartAutoLoginWatcher: Save OK -> {SharedConstants.LicenseFilePath}");

                _autoLoginLicenseId = _selectedLicenseId;
                _autoLoginDone = true;
                Interlocked.Exchange(ref _logoutDone, 0);

                AppendLog($"LOG IN - {userName} - {loginTime.ToString(SharedConstants.DateFormat)}");
                Dbg($"StartAutoLoginWatcher: _autoLoginDone=true _autoLoginLicenseId={_autoLoginLicenseId}");
            }
            catch (Exception ex) {
                Dbg("StartAutoLoginWatcher: OUTER EXCEPTION", ex);
                Interlocked.Exchange(ref _autoLoginWatcherStarted, 0);
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
                _autoLoginDone = false;
                _autoLoginLicenseId = null;
                _selectedLicenseId = null;
                Interlocked.Exchange(ref _autoLoginWatcherStarted, 0);
                Dbg("HandleModelClosed: autologin state cleared");
            }
            catch (Exception ex) {
                Dbg("HandleModelClosed: EXCEPTION while clearing state", ex);
            }

            try {
                if (_mode != "autologin") {
                    Dbg("HandleModelClosed: Dispatcher.Invoke(Close)...");
                    Application.Current.Dispatcher.Invoke(Close);
                }
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

            var info = SharedLicenseFileService.LoadOrCreate(SharedConstants.LicenseFilePath, licenseIdToRemove);

            var userExists = info.Logins.Any(login =>
                login != null &&
                !string.IsNullOrWhiteSpace(login.User) &&
                string.Equals(login.User, userName, StringComparison.OrdinalIgnoreCase));

            if (!userExists) {
                Dbg($"RemoveCurrentUserLogin: user={userName} not found in licenseId={licenseIdToRemove}");
                return;
            }

            SharedLicenseManager.Logout(info, userName);
            SharedLicenseFileService.Save(SharedConstants.LicenseFilePath, info);
            Dbg($"RemoveCurrentUserLogin: removed user={userName} from licenseId={licenseIdToRemove}");

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
                    _events.ModelLoad -= OnModelLoad;
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
        TSMO.Operation.RunMacro(@"Z:\000_PMJ\Tekla\HFT_SharedTool\SharedTool\ReadIn.cs");
    }

    private static void WriteOut() {
        TSMO.Operation.RunMacro(@"Z:\000_PMJ\Tekla\HFT_SharedTool\SharedTool\WriteOut.cs");
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
            if (ThemeToggleIcon == null) return;

            ThemeToggleIcon.Foreground = ThemeService.IsDark
                ? Application.Current.FindResource("TextSecondaryBrush") as Brush
                : new SolidColorBrush(Color.FromRgb(0x00, 0x5F, 0xB8));
        }
        catch (Exception ex) {
            Dbg("UpdateThemeToggleIconColor: EXCEPTION", ex);
        }
    }

    private void BtnRefresh_Click(object sender, RoutedEventArgs e) {
        try {
            if (RefreshIconRotate != null) {
                var animation = new DoubleAnimation {
                    From = 0, To = 360, Duration = TimeSpan.FromMilliseconds(500)
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