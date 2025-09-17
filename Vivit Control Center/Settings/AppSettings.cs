using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

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
                    return new AppSettings { RssFeeds = new List<string>(DefaultFeeds) };
                }
                var ser = new XmlSerializer(typeof(AppSettings));
                using (var fs = File.OpenRead(path))
                {
                    var loaded = (AppSettings)ser.Deserialize(fs);
                    if (loaded.RssFeeds == null) loaded.RssFeeds = new List<string>();
                    if (loaded.CustomWebModuleUrls == null) loaded.CustomWebModuleUrls = new List<CustomWebModuleUrl>();
                    if (loaded.ExternalPrograms == null) loaded.ExternalPrograms = new List<string>();
                    if (loaded.ExternalProgramsDetailed == null) loaded.ExternalProgramsDetailed = new List<ExternalProgram>();
                    // Migration: if detailed list empty but legacy list has entries
                    if (loaded.ExternalProgramsDetailed.Count == 0 && loaded.ExternalPrograms.Count > 0)
                    {
                        foreach (var p in loaded.ExternalPrograms)
                        {
                            if (string.IsNullOrWhiteSpace(p)) continue;
                            loaded.ExternalProgramsDetailed.Add(new ExternalProgram
                            {
                                Path = p,
                                Caption = System.IO.Path.GetFileNameWithoutExtension(p)
                            });
                        }
                    }
                    if (string.IsNullOrWhiteSpace(loaded.SocialLastNetwork)) loaded.SocialLastNetwork = "Fediverse";
                    if (string.IsNullOrWhiteSpace(loaded.Language)) loaded.Language = "en";
                    loaded.OfficeSuite = "MSOffice";
                    loaded.LibreOfficeProgramPath = string.Empty;
                    return loaded;
                }
            }
            catch
            {
                return new AppSettings { RssFeeds = new List<string>(), CustomWebModuleUrls = new List<CustomWebModuleUrl>(), ExternalPrograms = new List<string>(), ExternalProgramsDetailed = new List<ExternalProgram>() };
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
                // Keep legacy list in sync (paths only) for backward compatibility
                ExternalPrograms = new List<string>();
                foreach (var p in ExternalProgramsDetailed)
                {
                    if (!string.IsNullOrWhiteSpace(p?.Path)) ExternalPrograms.Add(p.Path);
                }
                var path = GetSettingsPath();
                var ser = new XmlSerializer(typeof(AppSettings));
                using (var fs = File.Create(path))
                {
                    ser.Serialize(fs, this);
                }
            }
            catch { }
        }
    }

    [Serializable]
    public class CustomWebModuleUrl
    {
        public string Tag { get; set; }
        public string Url { get; set; }
    }

    [Serializable]
    public class ExternalProgram
    {
        public string Path { get; set; }
        public string Caption { get; set; }
    }
}
