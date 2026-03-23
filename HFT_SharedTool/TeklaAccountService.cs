using System;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Win32;

namespace HFT_SharedTool;

internal static class TeklaAccountService {
    private const string RegistryKeyPath = @"Software\Trimble\Tekla Account";

    public static string GetTrimbleEmail() {
        try {
            using var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath);
            if (key == null) return "";

            foreach (var valueName in key.GetValueNames()) {
                if (key.GetValue(valueName) is not byte[] encrypted)
                    continue;

                try {
                    var decrypted = ProtectedData.Unprotect(
                        encrypted, null, DataProtectionScope.CurrentUser);

                    var json = Encoding.UTF8.GetString(decrypted);
                    var email = ExtractEmailFromUserIdentity(json);
                    if (!string.IsNullOrWhiteSpace(email))
                        return email;
                }
                catch {
                    // ignored
                }
            }
        }
        catch {
            // ignored
        }

        return "";
    }

    private static string ExtractEmailFromUserIdentity(string json) {
        const string identityKey = "\"UserIdentity\"";
        var identityIdx = json.IndexOf(identityKey, StringComparison.Ordinal);
        if (identityIdx < 0) return "";

        var blockStart = json.IndexOf('{', identityIdx + identityKey.Length);
        if (blockStart < 0) return "";

        var blockEnd = json.IndexOf('}', blockStart);
        if (blockEnd < 0) return "";

        var block = json.Substring(blockStart, blockEnd - blockStart);

        const string emailKey = "\"email\"";
        var emailIdx = block.IndexOf(emailKey, StringComparison.Ordinal);
        if (emailIdx < 0) return "";

        var colon = block.IndexOf(':', emailIdx + emailKey.Length);
        if (colon < 0) return "";

        var open = block.IndexOf('"', colon + 1);
        if (open < 0) return "";

        var close = block.IndexOf('"', open + 1);
        return close < 0 ? "" : block.Substring(open + 1, close - open - 1).Trim();
    }
}