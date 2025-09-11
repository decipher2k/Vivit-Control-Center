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
        public string OfficeSuite { get; set; } = "MSOffice"; // MSOffice or LibreOffice
        public string LibreOfficeProgramPath { get; set; } = string.Empty; // e.g. C:\Program Files\LibreOffice\program

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
                if (!File.Exists(path)) return new AppSettings();
                var ser = new XmlSerializer(typeof(AppSettings));
                using (var fs = File.OpenRead(path))
                {
                    return (AppSettings)ser.Deserialize(fs);
                }
            }
            catch { return new AppSettings(); }
        }

        public void Save()
        {
            try
            {
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
}
