using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;

namespace HFT_SharedTool;

internal static class SharedLicenseFileService {
    private static LockFileHandle AcquireFileLock(string lockPath, int timeoutMs = 8000) {
        var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);

        while (DateTime.UtcNow < deadline)
            try {
                var fs = new FileStream(
                    lockPath,
                    FileMode.OpenOrCreate,
                    FileAccess.ReadWrite,
                    FileShare.None);

                return new LockFileHandle(fs, lockPath);
            }
            catch (IOException) {
                Thread.Sleep(200);
            }

        return null;
    }

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

            if (parts.Length == 1 && line.Contains('@')) {
                current = new SharedLicenseInfo { LicenseId = line };
                result.Add(current);
                continue;
            }

            if (current == null)
                continue;

            switch (parts.Length) {
                case 3: {
                    var action = parts[0].Trim();
                    var userText = parts[1].Trim();
                    var dateText = parts[2].Trim();

                    if (dateText == "-") {
                        if (string.Equals(action, "LOG IN", StringComparison.OrdinalIgnoreCase))
                            current.Logins.Clear();
                        else if (string.Equals(action, "READ IN", StringComparison.OrdinalIgnoreCase)) {
                            current.ReadUser = null;
                            current.ReadTime = null;
                        }
                        else if (string.Equals(action, "WRITE OUT", StringComparison.OrdinalIgnoreCase)) {
                            current.WriteUser = null;
                            current.WriteTime = null;
                        }

                        break;
                    }

                    if (!DateTime.TryParseExact(
                            dateText,
                            SharedConstants.DateFormat,
                            CultureInfo.InvariantCulture,
                            DateTimeStyles.None,
                            out var dt))
                        break;

                    if (string.Equals(action, "LOG IN", StringComparison.OrdinalIgnoreCase)) {
                        current.Logins.Clear();

                        if (!string.IsNullOrWhiteSpace(userText) && userText != "-") {
                            var users = userText
                                .Split([','], StringSplitOptions.RemoveEmptyEntries)
                                .Select(u => u.Trim())
                                .Where(u => !string.IsNullOrWhiteSpace(u))
                                .Distinct(StringComparer.OrdinalIgnoreCase);

                            foreach (var u in users)
                                current.AddOrUpdateLogin(u, dt);
                        }
                    }
                    else if (string.Equals(action, "READ IN", StringComparison.OrdinalIgnoreCase)) {
                        current.ReadUser = userText == "-" ? null : userText;
                        current.ReadTime = dt;
                    }
                    else if (string.Equals(action, "WRITE OUT", StringComparison.OrdinalIgnoreCase)) {
                        current.WriteUser = userText == "-" ? null : userText;
                        current.WriteTime = dt;
                    }

                    break;
                }
                case 2: {
                    var action = parts[0].Trim();
                    var dateText = parts[1].Trim();

                    if (dateText == "-") {
                        if (string.Equals(action, "NEXT USABLE", StringComparison.OrdinalIgnoreCase))
                            current.NextUsable = null;
                        break;
                    }

                    if (!DateTime.TryParseExact(
                            dateText,
                            SharedConstants.DateFormat,
                            CultureInfo.InvariantCulture,
                            DateTimeStyles.None,
                            out var dt))
                        break;

                    if (string.Equals(action, "NEXT USABLE", StringComparison.OrdinalIgnoreCase))
                        current.NextUsable = dt;

                    break;
                }
            }
        }

        return result;
    }

    public static SharedLicenseInfo LoadOrCreate(string filePath, string licenseId) {
        var infos = LoadAll(filePath);

        foreach (var info in infos.Where(i =>
                     string.Equals(i.LicenseId, licenseId, StringComparison.OrdinalIgnoreCase)))
            return info;

        return new SharedLicenseInfo { LicenseId = licenseId };
    }

    public static void Save(string filePath, SharedLicenseInfo info) {
        var lockPath = filePath + ".lock";
        using var fileLock = AcquireFileLock(lockPath);
        var infos = LoadAll(filePath);
        var updated = false;

        for (var i = 0; i < infos.Count; i++) {
            if (!string.Equals(infos[i].LicenseId, info.LicenseId,
                    StringComparison.OrdinalIgnoreCase)) continue;

            infos[i] = info;
            updated = true;
            break;
        }

        if (!updated)
            infos.Add(info);

        WriteAll(filePath, infos);
    }

    public static void SaveMany(string filePath, IEnumerable<SharedLicenseInfo> modifiedInfos) {
        var lockPath = filePath + ".lock";
        using var fileLock = AcquireFileLock(lockPath);

        var infos = LoadAll(filePath);

        foreach (var modified in modifiedInfos) {
            var replaced = false;
            for (var i = 0; i < infos.Count; i++) {
                if (!string.Equals(infos[i].LicenseId, modified.LicenseId,
                        StringComparison.OrdinalIgnoreCase)) continue;
                infos[i] = modified;
                replaced = true;
                break;
            }

            if (!replaced)
                infos.Add(modified);
        }

        WriteAll(filePath, infos);
    }

    private static void WriteAll(string filePath, List<SharedLicenseInfo> infos) {
        var dir = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            Directory.CreateDirectory(dir);

        const int maxAttempts = 10;
        var delayMs = 150;

        for (var attempt = 1; attempt <= maxAttempts; attempt++)
            try {
                using var writer = new StreamWriter(filePath, false);

                foreach (var lic in infos.Where(l => !string.IsNullOrEmpty(l.LicenseId))) {
                    writer.WriteLine(lic.LicenseId);

                    string loginUser, loginDate;

                    if (lic.Logins.Count > 0) {
                        loginUser = string.Join(", ", lic.Logins.Select(x => x.User));
                        loginDate = lic.Logins[lic.Logins.Count - 1].Time.ToString(SharedConstants.DateFormat);
                    }
                    else {
                        loginUser = "-";
                        loginDate = "-";
                    }

                    writer.WriteLine("LOG IN - {0} - {1}", loginUser, loginDate);

                    var readUser = string.IsNullOrEmpty(lic.ReadUser) ? "-" : lic.ReadUser;
                    var readDate = lic.ReadTime.HasValue
                        ? lic.ReadTime.Value.ToString(SharedConstants.DateFormat)
                        : "-";
                    writer.WriteLine("READ IN - {0} - {1}", readUser, readDate);

                    var writeUser = string.IsNullOrEmpty(lic.WriteUser) ? "-" : lic.WriteUser;
                    var writeDate = lic.WriteTime.HasValue
                        ? lic.WriteTime.Value.ToString(SharedConstants.DateFormat)
                        : "-";
                    writer.WriteLine("WRITE OUT - {0} - {1}", writeUser, writeDate);

                    var nextUsableDate = lic.NextUsable.HasValue
                        ? lic.NextUsable.Value.ToString(SharedConstants.DateFormat)
                        : "-";
                    writer.WriteLine("NEXT USABLE - {0}", nextUsableDate);

                    writer.WriteLine();
                }

                return;
            }
            catch (IOException) {
                Thread.Sleep(delayMs);
                delayMs = Math.Min(delayMs * 2, 2000);
            }

        throw new IOException(
            $"Nie udało się zapisać pliku po {maxAttempts} próbach: {filePath}");
    }

    private sealed class LockFileHandle(FileStream fs, string lockPath) : IDisposable {
        private bool _disposed;

        public void Dispose() {
            if (_disposed) return;
            _disposed = true;
            try {
                fs.Dispose();
            }
            catch {
                // ignored
            }

            try {
                File.Delete(lockPath);
            }
            catch {
                // ignored
            }
        }
    }
}