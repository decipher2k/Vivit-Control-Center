using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;
using System.Security.Cryptography;
using System.Text;

namespace Vivit_Control_Center.Settings
{
    [Serializable]
    public class AppSettings
    {
        public string Theme { get; set; } = "Dark"; // Dark or Light
        public string DefaultLocalPath { get; set; } = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        public string DefaultScriptsPath { get; set; } = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        public List<string> DisabledModules { get; set; } = new List<string>();

        private string _officeSuite = "MSOffice"; // always MSOffice now
        public string OfficeSuite
        {
            get => string.IsNullOrWhiteSpace(_officeSuite) ? "MSOffice" : _officeSuite;
            set => _officeSuite = "MSOffice"; // ignore external changes, force default
        }

        [Obsolete("LibreOffice wird nicht mehr unterstützt.")]
        public string LibreOfficeProgramPath { get; set; } = string.Empty;

        public string AiService { get; set; } = "ChatGPT";
        public string MessengerService { get; set; } = "WhatsApp";
        public string ChatService { get; set; } = "Discord";
        public string NewsMode { get; set; } = "Webnews";
        public List<string> RssFeeds { get; set; } = new List<string>();
        public int RssMaxArticles { get; set; } = 60;

        public List<CustomWebModuleUrl> CustomWebModuleUrls { get; set; } = new List<CustomWebModuleUrl>();

        public string CustomFediverseUrl { get; set; } = "https://mastodon.social"; // override for Fediverse in Social module
        public string SocialLastNetwork { get; set; } = "Fediverse"; // persists last selected social network

        // NEW: UI language (IETF code e.g. en, de, fr, es, ru, zh, ja, eo)
        public string Language { get; set; } = "en";

        // Email accounts
        public List<EmailAccount> EmailAccounts { get; set; } = new List<EmailAccount>();

        // Legacy list (paths only) kept for migration
        public List<string> ExternalPrograms { get; set; } = new List<string>();
        // New detailed list with captions
        public List<ExternalProgram> ExternalProgramsDetailed { get; set; } = new List<ExternalProgram>();

        private static readonly List<string> DefaultFeeds = new List<string>
        {
            "https://rss.dw.com/rdf/rss-en-world",
            "https://feeds.bbci.co.uk/news/world/rss.xml"
        };

        public static string GetSettingsPath()
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VivitControlCenter");
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
            return Path.Combine(dir, "settings.xml");
        }

        public static AppSettings Load()
        {
            try
            {
                var path = GetSettingsPath();
                if (!File.Exists(path))
                {
                    return new AppSettings { RssFeeds = new List<string>(DefaultFeeds), EmailAccounts = new List<EmailAccount>() };
                }
                var ser = new XmlSerializer(typeof(AppSettings));
                using (var fs = File.OpenRead(path))
                {
                    var loaded = (AppSettings)ser.Deserialize(fs);
                    if (loaded.RssFeeds == null) loaded.RssFeeds = new List<string>();
                    if (loaded.CustomWebModuleUrls == null) loaded.CustomWebModuleUrls = new List<CustomWebModuleUrl>();
                    if (loaded.ExternalPrograms == null) loaded.ExternalPrograms = new List<string>();
                    if (loaded.ExternalProgramsDetailed == null) loaded.ExternalProgramsDetailed = new List<ExternalProgram>();
                    if (loaded.EmailAccounts == null) loaded.EmailAccounts = new List<EmailAccount>();
                    // Migration
                    if (loaded.ExternalProgramsDetailed.Count == 0 && loaded.ExternalPrograms.Count > 0)
                    {
                        foreach (var p in loaded.ExternalPrograms)
                        {
                            if (string.IsNullOrWhiteSpace(p)) continue;
                            loaded.ExternalProgramsDetailed.Add(new ExternalProgram { Path = p, Caption = System.IO.Path.GetFileNameWithoutExtension(p) });
                        }
                    }
                    if (string.IsNullOrWhiteSpace(loaded.SocialLastNetwork)) loaded.SocialLastNetwork = "Fediverse";
                    if (string.IsNullOrWhiteSpace(loaded.Language)) loaded.Language = "en";
                    loaded.OfficeSuite = "MSOffice";
                    loaded.LibreOfficeProgramPath = string.Empty;

                    // Decrypt sensitive fields if needed
                    foreach (var acc in loaded.EmailAccounts)
                    {
                        acc.Password = SecretProtector.UnprotectIfNeeded(acc.Password);
                        acc.OAuthAccessToken = SecretProtector.UnprotectIfNeeded(acc.OAuthAccessToken);
                        acc.OAuthRefreshToken = SecretProtector.UnprotectIfNeeded(acc.OAuthRefreshToken);
                        acc.OAuthClientSecret = SecretProtector.UnprotectIfNeeded(acc.OAuthClientSecret);
                    }

                    return loaded;
                }
            }
            catch
            {
                return new AppSettings { RssFeeds = new List<string>(), CustomWebModuleUrls = new List<CustomWebModuleUrl>(), ExternalPrograms = new List<string>(), ExternalProgramsDetailed = new List<ExternalProgram>(), EmailAccounts = new List<EmailAccount>() };
            }
        }

        public void Save()
        {
            try
            {
                OfficeSuite = "MSOffice";
                LibreOfficeProgramPath = string.Empty;
                if (string.IsNullOrWhiteSpace(SocialLastNetwork)) SocialLastNetwork = "Fediverse";
                if (string.IsNullOrWhiteSpace(Language)) Language = "en";
                if (ExternalProgramsDetailed == null) ExternalProgramsDetailed = new List<ExternalProgram>();
                if (EmailAccounts == null) EmailAccounts = new List<EmailAccount>();
                ExternalPrograms = new List<string>();
                foreach (var p in ExternalProgramsDetailed)
                    if (!string.IsNullOrWhiteSpace(p?.Path)) ExternalPrograms.Add(p.Path);

                // Temporarily encrypt sensitive fields before serialization, then restore
                var backups = new List<(EmailAccount acc, string pwd, string at, string rt, string cs)>();
                foreach (var acc in EmailAccounts)
                {
                    backups.Add((acc, acc.Password, acc.OAuthAccessToken, acc.OAuthRefreshToken, acc.OAuthClientSecret));
                    acc.Password = SecretProtector.ProtectIfNeeded(acc.Password);
                    acc.OAuthAccessToken = SecretProtector.ProtectIfNeeded(acc.OAuthAccessToken);
                    acc.OAuthRefreshToken = SecretProtector.ProtectIfNeeded(acc.OAuthRefreshToken);
                    acc.OAuthClientSecret = SecretProtector.ProtectIfNeeded(acc.OAuthClientSecret);
                }

                try
                {
                    var path = GetSettingsPath();
                    var ser = new XmlSerializer(typeof(AppSettings));
                    using (var fs = File.Create(path))
                        ser.Serialize(fs, this);
                }
                finally
                {
                    // Restore plain values in memory
                    foreach (var b in backups)
                    {
                        b.acc.Password = b.pwd;
                        b.acc.OAuthAccessToken = b.at;
                        b.acc.OAuthRefreshToken = b.rt;
                        b.acc.OAuthClientSecret = b.cs;
                    }
                }
            }
            catch { }
        }

        private static class SecretProtector
        {
            private const string Prefix = "enc:";

            public static string ProtectIfNeeded(string value)
            {
                if (string.IsNullOrEmpty(value)) return value;
                if (value.StartsWith(Prefix, StringComparison.Ordinal)) return value; // already protected
                try
                {
                    var bytes = Encoding.UTF8.GetBytes(value);
                    var protectedBytes = ProtectedData.Protect(bytes, null, DataProtectionScope.CurrentUser);
                    return Prefix + Convert.ToBase64String(protectedBytes);
                }
                catch { return value; }
            }

            public static string UnprotectIfNeeded(string value)
            {
                if (string.IsNullOrEmpty(value)) return value;
                if (!value.StartsWith(Prefix, StringComparison.Ordinal)) return value;
                try
                {
                    var b64 = value.Substring(Prefix.Length);
                    var protectedBytes = Convert.FromBase64String(b64);
                    var bytes = ProtectedData.Unprotect(protectedBytes, null, DataProtectionScope.CurrentUser);
                    return Encoding.UTF8.GetString(bytes ?? new byte[0]);
                }
                catch { return value; }
            }
        }
    }

    [Serializable]
    public class CustomWebModuleUrl { public string Tag { get; set; } public string Url { get; set; } }

    [Serializable]
    public class ExternalProgram { public string Path { get; set; } public string Caption { get; set; } }

    [Serializable]
    public class EmailAccount
    {
        public string DisplayName { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public string ImapHost { get; set; }
        public int ImapPort { get; set; } = 993;
        public bool ImapUseSsl { get; set; } = true;
        public string SmtpHost { get; set; }
        public int SmtpPort { get; set; } = 587;
        public bool SmtpUseSsl { get; set; } = true;

        // OAuth2 configuration
        public string AuthMethod { get; set; } = "Password"; // Password or OAuth2
        public string OAuthProvider { get; set; } = ""; // Google, Microsoft, Custom
        public string OAuthAccessToken { get; set; }
        public string OAuthRefreshToken { get; set; }
        public DateTime OAuthTokenExpiryUtc { get; set; }
        public string OAuthClientId { get; set; }
        public string OAuthClientSecret { get; set; }
        public string OAuthTenant { get; set; } // for Microsoft (common, organizations, consumers or GUID)
        public string OAuthScope { get; set; } // space-separated scopes
        public string OAuthTokenEndpoint { get; set; } // override token endpoint
    }
}
