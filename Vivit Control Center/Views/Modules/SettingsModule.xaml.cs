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

            // Theme items
            cmbTheme.Items.Clear();
            cmbTheme.Items.Add("Dark");
            cmbTheme.Items.Add("Light");
            cmbTheme.SelectedItem = string.Equals(_settings.Theme, "Light", StringComparison.OrdinalIgnoreCase) ? "Light" : "Dark";

            // Office suite
            cmbOffice.SelectedItem = _settings.OfficeSuite ?? "MSOffice";
            txtLoPath.Text = _settings.LibreOfficeProgramPath ?? string.Empty;

            // Parameter selections
            SelectComboItem(aiServiceSelector, _settings.AiService, "ChatGPT");
            SelectComboItem(messengerServiceSelector, _settings.MessengerService, "WhatsApp");
            SelectComboItem(chatServiceSelector, _settings.ChatService, "Discord");

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

        private void btnPickLocalPath_Click(object sender, RoutedEventArgs e)
        {
            var path = PickFolder();
            if (!string.IsNullOrEmpty(path)) txtLocalPath.Text = path;
        }

        private void btnPickScriptsPath_Click(object sender, RoutedEventArgs e)
        {
            var path = PickFolder();
            if (!string.IsNullOrEmpty(path)) txtScriptsPath.Text = path;
        }

        private void btnPickLoPath_Click(object sender, RoutedEventArgs e)
        {
            var path = PickFolder();
            if (!string.IsNullOrEmpty(path)) txtLoPath.Text = path;
        }

        private string PickFolder()
        {
            var dlg = new OpenFileDialog
            {
                CheckFileExists = false,
                CheckPathExists = true,
                ValidateNames = false,
                FileName = "Ordner auswaehlen"
            };
            var res = dlg.ShowDialog();
            if (res == true)
            {
                try { return Path.GetDirectoryName(dlg.FileName); } catch { }
            }
            return null;
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            _settings.Theme = cmbTheme.SelectedItem?.ToString() ?? "Dark";
            _settings.OfficeSuite = (cmbOffice.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? _settings.OfficeSuite ?? "MSOffice";
            _settings.LibreOfficeProgramPath = txtLoPath.Text ?? string.Empty;
            _settings.DefaultLocalPath = txtLocalPath.Text ?? string.Empty;
            _settings.DefaultScriptsPath = txtScriptsPath.Text ?? string.Empty;

            _settings.AiService = (aiServiceSelector.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? _settings.AiService;
            _settings.MessengerService = (messengerServiceSelector.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? _settings.MessengerService;
            _settings.ChatService = (chatServiceSelector.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? _settings.ChatService;

            var enabled = new List<string>();
            foreach (var child in modulesPanel.Children)
            {
                if (child is CheckBox cb && cb.IsChecked == true)
                    enabled.Add(cb.Content.ToString());
            }
            var disabled = AllTags.Where(t => !enabled.Contains(t, StringComparer.OrdinalIgnoreCase)).ToList();
            disabled.RemoveAll(t => string.Equals(t, "Settings", StringComparison.OrdinalIgnoreCase));
            _settings.DisabledModules = disabled;

            _settings.Save();
            MessageBox.Show("Einstellungen gespeichert. Neustart empfohlen.");
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            Init();
        }
    }
}
