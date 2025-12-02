using System;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Media;

namespace HFT_SharedTool
{
    public partial class MainWindow {
        private const string LicenseFilePath = @"Z:\000_PMJ\Tekla\HFT_sharing_lic.txt";
        private const string DateFormat = "yyyy-MM-dd HH:mm";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void AppendLog(string text)
        {
            if (LogTextBox != null)
            {
                LogTextBox.AppendText(text + Environment.NewLine);
                LogTextBox.ScrollToEnd();
            }
        }

        private void AppendColoredStatus(string text, bool isUsable)
        {
            if (LogTextBox == null)
            {
                AppendLog(text + (isUsable ? " [USABLE]" : " [NOT USABLE]"));
                return;
            }

            var range = new System.Windows.Documents.TextRange(LogTextBox.Document.ContentEnd, LogTextBox.Document.ContentEnd)
            {
                Text = text + Environment.NewLine
            };
            range.ApplyPropertyValue(System.Windows.Documents.TextElement.ForegroundProperty,
                isUsable ? Brushes.Green : Brushes.Red);
            LogTextBox.ScrollToEnd();
        }

        private class SharedLicenseInfo
        {
            public string LicenseId { get; set; }
            public string LoginUser { get; set; }
            public DateTime? LoginTime { get; set; }
            public string ReadUser { get; set; }
            public DateTime? ReadTime { get; set; }
            public string WriteUser { get; set; }
            public DateTime? WriteTime { get; set; }
            public DateTime? NextUsable { get; set; }
        }

        private static class SharedLicenseFileService
        {
            public static SharedLicenseInfo Load(string filePath)
            {
                if (!File.Exists(filePath))
                    return null;

                var lines = File.ReadAllLines(filePath);
                if (lines.Length == 0)
                    return null;

                var info = new SharedLicenseInfo {
                    LicenseId = lines[0].Trim()
                };

                for (var i = 1; i < lines.Length; i++)
                {
                    var line = lines[i].Trim();
                    if (line.Length == 0)
                        continue;

                    var parts = line.Split(new[] { " - " }, StringSplitOptions.None);
                    switch (parts.Length) {
                        case 3: {
                            var action = parts[0].Trim();
                            var user = parts[1].Trim();
                            var dateText = parts[2].Trim();
                            if (!DateTime.TryParseExact(dateText, DateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                                continue;

                            if (string.Equals(action, "LOG IN", StringComparison.OrdinalIgnoreCase))
                            {
                                info.LoginUser = user;
                                info.LoginTime = dt;
                            }
                            else if (string.Equals(action, "READ IN", StringComparison.OrdinalIgnoreCase))
                            {
                                info.ReadUser = user;
                                info.ReadTime = dt;
                            }
                            else if (string.Equals(action, "WRITE OUT", StringComparison.OrdinalIgnoreCase))
                            {
                                info.WriteUser = user;
                                info.WriteTime = dt;
                            }

                            break;
                        }
                        case 2: {
                            var action = parts[0].Trim();
                            var dateText = parts[1].Trim();
                            if (!DateTime.TryParseExact(dateText, DateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                                continue;

                            if (string.Equals(action, "NEXT USABLE", StringComparison.OrdinalIgnoreCase))
                            {
                                info.NextUsable = dt;
                            }

                            break;
                        }
                    }
                }

                return info;
            }

            public static SharedLicenseInfo LoadOrCreate(string filePath, string licenseId)
            {
                var info = Load(filePath);
                if (info == null)
                {
                    info = new SharedLicenseInfo
                    {
                        LicenseId = licenseId
                    };
                }
                else if (string.IsNullOrEmpty(info.LicenseId))
                {
                    info.LicenseId = licenseId;
                }

                return info;
            }

            public static void Save(string filePath, SharedLicenseInfo info)
            {
                var dir = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                using (var writer = new StreamWriter(filePath, false))
                {
                    writer.WriteLine(info.LicenseId ?? string.Empty);

                    if (info.LoginTime.HasValue && !string.IsNullOrEmpty(info.LoginUser))
                        writer.WriteLine("LOG IN - {0} - {1}", info.LoginUser, info.LoginTime.Value.ToString(DateFormat));

                    if (info.ReadTime.HasValue && !string.IsNullOrEmpty(info.ReadUser))
                        writer.WriteLine("READ IN - {0} - {1}", info.ReadUser, info.ReadTime.Value.ToString(DateFormat));

                    if (info.WriteTime.HasValue && !string.IsNullOrEmpty(info.WriteUser))
                        writer.WriteLine("WRITE OUT - {0} - {1}", info.WriteUser, info.WriteTime.Value.ToString(DateFormat));

                    if (info.NextUsable.HasValue)
                        writer.WriteLine("NEXT USABLE - {0}", info.NextUsable.Value.ToString(DateFormat));
                }
            }
        }

        private static class SharedLicenseManager
        {
            private static readonly TimeSpan HoldDuration = TimeSpan.FromHours(4);

            public static void Login(SharedLicenseInfo info, string userName, DateTime loginTime)
            {
                info.LoginUser = userName;
                info.LoginTime = loginTime;
            }

            public static void ReadIn(SharedLicenseInfo info, string userName, DateTime readTime)
            {
                info.ReadUser = userName;
                info.ReadTime = readTime;
                info.NextUsable = readTime + HoldDuration;
            }

            public static void WriteOut(SharedLicenseInfo info, string userName, DateTime writeTime)
            {
                info.WriteUser = userName;
                info.WriteTime = writeTime;
                info.NextUsable = writeTime + HoldDuration;
            }

            public static string FormatStatus(SharedLicenseInfo info, DateTime now, out bool isUsableNow)
            {
                isUsableNow = true;

                string activity;
                DateTime activityTime;

                var hasRead = info.ReadTime.HasValue;
                var hasWrite = info.WriteTime.HasValue;

                if (hasRead && (!hasWrite || info.ReadTime.Value > info.WriteTime.Value))
                {
                    activity = "READ IN";
                    activityTime = info.ReadTime.Value;
                }
                else if (hasWrite)
                {
                    activity = "WRITE OUT";
                    activityTime = info.WriteTime.Value;
                }
                else
                {
                    activity = "BRAK DANYCH";
                    activityTime = now;
                }

                var user = !string.IsNullOrEmpty(info.LoginUser) ? info.LoginUser : "WOLNE";

                if (info.NextUsable.HasValue)
                    isUsableNow = now >= info.NextUsable.Value;

                var nextUsableText = info.NextUsable.HasValue
                    ? info.NextUsable.Value.ToString(DateFormat)
                    : "-";

                return
                    $"{info.LicenseId} ({user}) - {activity} {activityTime:yyyy-MM-dd HH:mm} - USABLE {nextUsableText}";
            }
        }

        private void BtnStandaloneCheck_Click(object sender, RoutedEventArgs e)
        {
            var info = SharedLicenseFileService.Load(LicenseFilePath);
            if (info == null)
            {
                AppendLog("Brak danych licencji: " + LicenseFilePath);
                return;
            }

            var line = SharedLicenseManager.FormatStatus(info, DateTime.Now, out var isUsableNow);
            AppendColoredStatus(line, isUsableNow);
        }

        private void BtnLogin_Click(object sender, RoutedEventArgs e)
        {
            if (!TryGetLicenseFromLog(out var licenseId, out var loginTime))
            {
                AppendLog("Nie znaleziono wpisu UserInfo w logu Tekli.");
                return;
            }

            var userName = Environment.UserName;
            var info = SharedLicenseFileService.LoadOrCreate(LicenseFilePath, licenseId);
            SharedLicenseManager.Login(info, userName, loginTime);
            SharedLicenseFileService.Save(LicenseFilePath, info);
            AppendLog($"LOG IN - {userName} - {loginTime.ToString(DateFormat)}");
        }

        private void BtnReadIn_Click(object sender, RoutedEventArgs e)
        {
            var info = SharedLicenseFileService.Load(LicenseFilePath);
            if (info == null)
            {
                AppendLog("Brak danych licencji do READ IN.");
                return;
            }

            var userName = Environment.UserName;
            var now = DateTime.Now;
            SharedLicenseManager.ReadIn(info, userName, now);
            SharedLicenseFileService.Save(LicenseFilePath, info);
            AppendLog($"READ IN - {userName} - {now.ToString(DateFormat)}");
        }

        private void BtnWriteOut_Click(object sender, RoutedEventArgs e)
        {
            var info = SharedLicenseFileService.Load(LicenseFilePath);
            if (info == null)
            {
                AppendLog("Brak danych licencji do WRITE OUT.");
                return;
            }

            var userName = Environment.UserName;
            var now = DateTime.Now;
            SharedLicenseManager.WriteOut(info, userName, now);
            SharedLicenseFileService.Save(LicenseFilePath, info);
            AppendLog($"WRITE OUT - {userName} - {now.ToString(DateFormat)}");
        }

        private void BtnCheck_Click(object sender, RoutedEventArgs e)
        {
            BtnStandaloneCheck_Click(sender, e);
        }

        private static bool TryGetLicenseFromLog(out string licenseId, out DateTime loginTime)
        {
            licenseId = null;
            loginTime = DateTime.MinValue;

            var userName = Environment.UserName;
            var logPath = Path.Combine(@"C:\TeklaStructuresModels", "TeklaStructures_" + userName + ".log");
            if (!File.Exists(logPath))
                return false;

            var lines = File.ReadAllLines(logPath);
            for (var i = lines.Length - 1; i >= 0; i--)
            {
                var line = lines[i];
                if (line.IndexOf("Info: UserInfo", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                var parts = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length < 6)
                    continue;

                var datePart = parts[0];
                var timePart = parts[1];
                var dateTimeText = datePart + " " + timePart;

                if (!DateTime.TryParseExact(dateTimeText, DateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                    continue;

                var lic = parts[4].Trim();
                if (lic.Length == 0)
                    continue;

                loginTime = dt;
                licenseId = lic;
                return true;
            }

            return false;
        }
    }
}