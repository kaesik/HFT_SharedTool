using System;
using System.Linq;

namespace HFT_SharedTool;

internal static class SharedLicenseManager {
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
        out bool isUsableForCurrentUser) {

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

        var user = info.Logins.Count > 0
            ? string.Join(", ", info.Logins.Select(x => x.User))
            : "WOLNE";

        var isAfterHold = !info.NextUsable.HasValue || now >= info.NextUsable.Value;

        var isCurrentHolder =
            !string.IsNullOrEmpty(latestActivityUserName) &&
            !string.IsNullOrEmpty(currentUser) &&
            string.Equals(latestActivityUserName, currentUser, StringComparison.OrdinalIgnoreCase);

        isUsableForCurrentUser = isAfterHold || isCurrentHolder;

        var nextUsableText = info.NextUsable.HasValue
            ? info.NextUsable.Value.ToString(SharedConstants.DateFormat)
            : "-";

        return
            $"{info.LicenseId} ({user}) - {activity} {activityUserDisplay}{activityTime:yyyy-MM-dd HH:mm} - USABLE {nextUsableText}";
    }
}
