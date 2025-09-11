using System;
using System.Threading.Tasks;
using System.Windows.Controls;
using Vivit_Control_Center.Settings;
using Vivit_Control_Center.Views; // for WebViewModule

namespace Vivit_Control_Center.Views.Modules
{
    public partial class SocialModule : BaseSimpleModule
    {
        private AppSettings _settings;
        private WebViewModule _webViewModule;
        private bool _initialized;
        private bool _navigated;

        public SocialModule()
        {
            InitializeComponent();
            Loaded += (_, __) => Init();
        }

        private async void Init()
        {
            if (_initialized) return;
            _initialized = true;
            _settings = AppSettings.Load();
            _webViewModule = new WebViewModule("Social");
            BrowserHost.Content = _webViewModule.View;
            // Set selection based on persisted setting (default Fediverse)
            var last = _settings.SocialLastNetwork ?? "Fediverse";
            int index = 0;
            for (int i = 0; i < cmbSocial.Items.Count; i++)
            {
                var item = cmbSocial.Items[i] as ComboBoxItem;
                if (item != null && string.Equals(item.Content?.ToString(), last, StringComparison.OrdinalIgnoreCase)) { index = i; break; }
            }
            cmbSocial.SelectedIndex = index;
            await EnsureNavigateAsync();
        }

        private string ResolveSocialUrl(string key)
        {
            switch (key)
            {
                case "Facebook": return "https://facebook.com";
                case "X.com": return "http://x.com";
                case "Bluesky": return "https://bsky.app";
                case "Fediverse": return string.IsNullOrWhiteSpace(_settings.CustomFediverseUrl) ? "https://mastodon.social" : _settings.CustomFediverseUrl.Trim();
                default: return "https://facebook.com";
            }
        }

        private async Task EnsureNavigateAsync()
        {
            if (!_initialized) return;
            var sel = (cmbSocial.SelectedItem as ComboBoxItem)?.Content?.ToString();
            if (string.IsNullOrWhiteSpace(sel)) return;
            var url = ResolveSocialUrl(sel);
            _settings.SocialLastNetwork = sel; // persist selection first
            _settings.Save();
            await _webViewModule.NavigateToAsync(url);
            _navigated = true;
        }

        private async void cmbSocial_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_initialized) return;
            await EnsureNavigateAsync();
        }

        public override async Task PreloadAsync()
        {
            if (_webViewModule != null && !_navigated)
                await _webViewModule.PreloadAsync();
            await base.PreloadAsync();
        }
    }
}
