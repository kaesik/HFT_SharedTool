using System;
using System.Collections.Generic;
using System.Linq;

namespace HFT_SharedTool;

internal class SharedLicenseInfo {
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

        foreach (var entry in Logins.Where(e =>
                     string.Equals(e.User, user, StringComparison.OrdinalIgnoreCase))) {
            entry.Time = time;
            return;
        }

        Logins.Add(new LoginEntry { User = user, Time = time });
    }

    public void RemoveLogin(string user) {
        if (string.IsNullOrEmpty(user))
            return;

        Logins.RemoveAll(l => string.Equals(l.User, user, StringComparison.OrdinalIgnoreCase));
    }

    public class LoginEntry {
        public string User { get; set; }
        public DateTime Time { get; set; }
    }
}
