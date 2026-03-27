using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Win32;

namespace HFT_SharedTool;

internal static class TeklaAccountService {
    private const string RegistryKeyPath = @"Software\Trimble\Tekla Account";

    public static string GetTrimbleEmail() {
        try {
            using var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath);
            if (key == null)
                return "";

            var emails = new List<string>();

            foreach (var valueName in key.GetValueNames()) {
                if (key.GetValue(valueName) is not byte[] encrypted)
                    continue;

                try {
                    var decrypted = ProtectedData.Unprotect(
                        encrypted,
                        null,
                        DataProtectionScope.CurrentUser);

                    var json = Encoding.UTF8.GetString(decrypted);
                    var email = ExtractEmailFromUserIdentity(json);

                    if (!string.IsNullOrWhiteSpace(email))
                        emails.Add(email.Trim());
                }
                catch {
                    // ignored
                }
            }

            if (emails.Count == 0)
                return "";

            var distinctEmails = emails
                .Where(email => !string.IsNullOrWhiteSpace(email))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var matchedDefaultLicense = distinctEmails.FirstOrDefault(email =>
                SharedConstants.DefaultLicenseIds.Contains(email, StringComparer.OrdinalIgnoreCase));

            return !string.IsNullOrWhiteSpace(matchedDefaultLicense) ? matchedDefaultLicense : distinctEmails[0];
        }
        catch {
            return "";
        }
    }

    private static string ExtractEmailFromUserIdentity(string json) {
        const string identityKey = "\"UserIdentity\"";
        var identityIndex = json.IndexOf(identityKey, StringComparison.OrdinalIgnoreCase);
        if (identityIndex < 0)
            return "";

        var blockStart = json.IndexOf('{', identityIndex + identityKey.Length);
        if (blockStart < 0)
            return "";

        var depth = 0;
        var blockEnd = -1;

        for (var i = blockStart; i < json.Length; i++)
            if (json[i] == '{')
                depth++;
            else if (json[i] == '}') {
                depth--;
                if (depth == 0) {
                    blockEnd = i;
                    break;
                }
            }

        if (blockEnd < 0)
            return "";

        var block = json.Substring(blockStart, blockEnd - blockStart + 1);

        const string emailKey = "\"email\"";
        var emailIndex = block.IndexOf(emailKey, StringComparison.OrdinalIgnoreCase);
        if (emailIndex < 0)
            return "";

        var colonIndex = block.IndexOf(':', emailIndex + emailKey.Length);
        if (colonIndex < 0)
            return "";

        var openQuoteIndex = block.IndexOf('"', colonIndex + 1);
        if (openQuoteIndex < 0)
            return "";

        var closeQuoteIndex = block.IndexOf('"', openQuoteIndex + 1);
        return closeQuoteIndex < 0
            ? ""
            : block.Substring(openQuoteIndex + 1, closeQuoteIndex - openQuoteIndex - 1).Trim();
    }
}