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

        // Custom Web Module URLs (user overrides) excluding AI/Messenger/Chat/News
        public List<CustomWebModuleUrl> CustomWebModuleUrls { get; set; } = new List<CustomWebModuleUrl>();

        public string CustomFediverseUrl { get; set; } = "https://mastodon.social"; // override for Fediverse in Social module
        public string SocialLastNetwork { get; set; } = "Fediverse"; // persists last selected social network

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
                    if (string.IsNullOrWhiteSpace(loaded.SocialLastNetwork)) loaded.SocialLastNetwork = "Fediverse";
                    loaded.OfficeSuite = "MSOffice";
                    loaded.LibreOfficeProgramPath = string.Empty;
                    return loaded;
                }
            }
            catch
            {
                return new AppSettings { RssFeeds = new List<string>(), CustomWebModuleUrls = new List<CustomWebModuleUrl>() };
            }
        }

        public void Save()
        {
            try
            {
                OfficeSuite = "MSOffice";
                LibreOfficeProgramPath = string.Empty;
                if (string.IsNullOrWhiteSpace(SocialLastNetwork)) SocialLastNetwork = "Fediverse";
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
}
