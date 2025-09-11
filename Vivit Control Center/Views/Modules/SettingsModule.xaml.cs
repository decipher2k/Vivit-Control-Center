using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using Vivit_Control_Center.Settings;

namespace Vivit_Control_Center.Views.Modules
{
    public partial class SettingsModule : BaseSimpleModule
    {
        private AppSettings _settings;
        private static readonly string[] AllTags = new[]
        {
            "AI","News","Messenger","Chat","Explorer","Office","Notes","Media Player","Steam",
            "Webbrowser","Order Food","eBay","Temu","Terminal","Scripting","SSH","SFTP","Settings"
        };

        public SettingsModule()
        {
            InitializeComponent();
            Loaded += (_, __) => Init();
        }

        private void Init()
        {
            _settings = AppSettings.Load();
            if (string.IsNullOrWhiteSpace(_settings.OfficeSuite)) _settings.OfficeSuite = "MSOffice"; // enforce default
            if (_settings.OfficeSuite != "MSOffice") _settings.OfficeSuite = "MSOffice"; // force MSOffice always
            if (!string.IsNullOrEmpty(_settings.LibreOfficeProgramPath)) _settings.LibreOfficeProgramPath = string.Empty; // clear obsolete path

            // Theme items
            cmbTheme.Items.Clear();
            cmbTheme.Items.Add("Dark");
            cmbTheme.Items.Add("Light");
            cmbTheme.SelectedItem = string.Equals(_settings.Theme, "Light", StringComparison.OrdinalIgnoreCase) ? "Light" : "Dark";

            // Parameter selections
            SelectComboItem(aiServiceSelector, _settings.AiService, "ChatGPT");
            SelectComboItem(messengerServiceSelector, _settings.MessengerService, "WhatsApp");
            SelectComboItem(chatServiceSelector, _settings.ChatService, "Discord");
            SelectComboItem(newsModeSelector, _settings.NewsMode, "Webnews");
            btnEditFeeds.Visibility = IsRssSelected() ? Visibility.Visible : Visibility.Collapsed;
            if (txtRssMax != null) txtRssMax.Text = (_settings.RssMaxArticles > 0 ? _settings.RssMaxArticles : 60).ToString();

            // Paths
            txtLocalPath.Text = _settings.DefaultLocalPath ?? string.Empty;
            txtScriptsPath.Text = _settings.DefaultScriptsPath ?? string.Empty;

            // Module checkboxes
            modulesPanel.Children.Clear();
            foreach (var tag in AllTags)
            {
                if (string.Equals(tag, "Settings", StringComparison.OrdinalIgnoreCase)) continue; // always visible
                var cb = new CheckBox
                {
                    Content = tag,
                    Margin = new Thickness(0, 2, 0, 2),
                    IsChecked = !_settings.DisabledModules.Contains(tag, StringComparer.OrdinalIgnoreCase)
                };
                modulesPanel.Children.Add(cb);
            }
        }

        private bool IsRssSelected() => (newsModeSelector.SelectedItem as ComboBoxItem)?.Content?.ToString().Equals("RSS", StringComparison.OrdinalIgnoreCase) == true;

        private void SelectComboItem(ComboBox combo, string value, string fallback)
        {
            if (combo == null) return;
            string target = string.IsNullOrWhiteSpace(value) ? fallback : value;
            combo.SelectedIndex = -1;
            foreach (var item in combo.Items.OfType<ComboBoxItem>())
            {
                if (string.Equals(item.Content?.ToString(), target, StringComparison.OrdinalIgnoreCase))
                {
                    combo.SelectedItem = item;
                    return;
                }
            }
            if (combo.Items.Count > 0) combo.SelectedIndex = 0;
        }

        private void btnPickLocalPath_Click(object sender, RoutedEventArgs e) { var p = PickFolder(); if (!string.IsNullOrEmpty(p)) txtLocalPath.Text = p; }
        private void btnPickScriptsPath_Click(object sender, RoutedEventArgs e) { var p = PickFolder(); if (!string.IsNullOrEmpty(p)) txtScriptsPath.Text = p; }

        private string PickFolder()
        {
            var dlg = new OpenFileDialog { CheckFileExists = false, CheckPathExists = true, ValidateNames = false, FileName = "Ordner auswaehlen" };
            var res = dlg.ShowDialog();
            if (res == true) { try { return Path.GetDirectoryName(dlg.FileName); } catch { } }
            return null;
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            _settings.Theme = cmbTheme.SelectedItem?.ToString() ?? "Dark";
            _settings.AiService = (aiServiceSelector.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? _settings.AiService;
            _settings.MessengerService = (messengerServiceSelector.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? _settings.MessengerService;
            _settings.ChatService = (chatServiceSelector.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? _settings.ChatService;
            _settings.NewsMode = (newsModeSelector.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? _settings.NewsMode;
            _settings.OfficeSuite = "MSOffice"; // force
            _settings.LibreOfficeProgramPath = string.Empty; // clear
            int max; if (txtRssMax != null && int.TryParse(txtRssMax.Text, out max) && max > 0) _settings.RssMaxArticles = max; else if (_settings.RssMaxArticles <= 0) _settings.RssMaxArticles = 60;
            _settings.DefaultLocalPath = txtLocalPath.Text ?? string.Empty;
            _settings.DefaultScriptsPath = txtScriptsPath.Text ?? string.Empty;

            var enabled = new List<string>();
            foreach (var child in modulesPanel.Children) if (child is CheckBox cb && cb.IsChecked == true) enabled.Add(cb.Content.ToString());
            var disabled = AllTags.Where(t => !enabled.Contains(t, StringComparer.OrdinalIgnoreCase)).ToList();
            disabled.RemoveAll(t => string.Equals(t, "Settings", StringComparison.OrdinalIgnoreCase));
            _settings.DisabledModules = disabled;
            _settings.Save();
            MessageBox.Show("Einstellungen gespeichert. Neustart empfohlen.");
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e) => Init();
        private void newsModeSelector_SelectionChanged(object sender, SelectionChangedEventArgs e) => btnEditFeeds.Visibility = IsRssSelected() ? Visibility.Visible : Visibility.Collapsed;
        private void btnEditFeeds_Click(object sender, RoutedEventArgs e)
        {
            var current = string.Join(Environment.NewLine, _settings.RssFeeds ?? new List<string>());
            var dlg = new RssFeedsDialog(current);
            if (dlg.ShowDialog() == true)
            {
                var lines = (dlg.FeedsText ?? string.Empty)
                    .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(l => l.Trim())
                    .Where(l => l.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
                _settings.RssFeeds = lines;
                _settings.Save();
            }
        }
    }
}
