using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;

namespace Vivit_Control_Center.Localization
{
    public static class LocalizationManager
    {
        private static readonly HashSet<string> Supported = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { "en","de","fr","es","ru","zh","ja","eo" };

        public static void ApplyLanguage(string code)
        {
            if (string.IsNullOrWhiteSpace(code) || !Supported.Contains(code)) code = "en";
            try
            {
                // Remove old localization dictionaries
                var toRemove = new List<ResourceDictionary>();
                foreach (var rd in Application.Current.Resources.MergedDictionaries)
                {
                    if (rd.Source != null && rd.Source.OriginalString.Contains("/Localization/Strings."))
                        toRemove.Add(rd);
                }
                foreach (var r in toRemove) Application.Current.Resources.MergedDictionaries.Remove(r);

                // Always add English base for fallback
                var baseUri = new Uri($"/Vivit Control Center;component/Localization/Strings.en.xaml", UriKind.Relative);
                Application.Current.Resources.MergedDictionaries.Add(new ResourceDictionary { Source = baseUri });

                if (!string.Equals(code, "en", StringComparison.OrdinalIgnoreCase))
                {
                    var langUri = new Uri($"/Vivit Control Center;component/Localization/Strings.{code}.xaml", UriKind.Relative);
                    Application.Current.Resources.MergedDictionaries.Add(new ResourceDictionary { Source = langUri });
                }

                try { CultureInfo.CurrentUICulture = new CultureInfo(code); } catch { }
            }
            catch
            {
                if (!string.Equals(code, "en", StringComparison.OrdinalIgnoreCase))
                {
                    ApplyLanguage("en");
                }
            }
        }

        public static string GetString(string key, string fallback = null)
        {
            try
            {
                var obj = Application.Current.TryFindResource(key);
                if (obj is string s && !string.IsNullOrEmpty(s)) return s;
            }
            catch { }
            return fallback ?? key;
        }
    }
}
