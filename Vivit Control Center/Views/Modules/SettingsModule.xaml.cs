using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;
using System.Diagnostics;
using System.Security.Principal;
using Vivit_Control_Center.Settings;

namespace Vivit_Control_Center.Views.Modules
{
    public partial class SettingsModule : BaseSimpleModule
    {
        private AppSettings _settings;
        private static readonly string[] AllTags = new[]
        {
            "AI","News","Messenger","Chat","Explorer","Office","Notes","Media Player","Steam",
            "Webbrowser","Order Food","eBay","Temu","Terminal","Scripting","SSH","SFTP","Social","Settings"
        };

        private static readonly HashSet<string> DynamicSelectionTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { "AI", "Messenger", "Chat", "News" };

        private static readonly HashSet<string> WebModules = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "News","Media Player","Order Food","eBay","Temu","AI","Messenger","Chat","Webbrowser","Social"
        };

        private static readonly HashSet<string> NoCustomUrlTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { "Media Player", "eBay", "Temu" };

        private const double ModuleLabelColumnWidth = 180;
        private const string WinLogonKeyPath = @"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon";
        private string _originalShellCached;

        public SettingsModule()
        {
            InitializeComponent();
            Loaded += (_, __) => Init();
        }

        private void Init()
        {
            _settings = AppSettings.Load();
            if (string.IsNullOrWhiteSpace(_settings.OfficeSuite)) _settings.OfficeSuite = "MSOffice"; if (_settings.OfficeSuite != "MSOffice") _settings.OfficeSuite = "MSOffice"; if (!string.IsNullOrEmpty(_settings.LibreOfficeProgramPath)) _settings.LibreOfficeProgramPath = string.Empty;
            SelectComboItem(aiServiceSelector, _settings.AiService, "ChatGPT");
            SelectComboItem(messengerServiceSelector, _settings.MessengerService, "WhatsApp");
            SelectComboItem(chatServiceSelector, _settings.ChatService, "Discord");
            SelectComboItem(newsModeSelector, _settings.NewsMode, "Webnews");
            btnEditFeeds.Visibility = IsRssSelected() ? Visibility.Visible : Visibility.Collapsed;
            if (txtRssMax != null) txtRssMax.Text = (_settings.RssMaxArticles > 0 ? _settings.RssMaxArticles : 60).ToString();
            if (txtLocalPath != null) txtLocalPath.Text = _settings.DefaultLocalPath ?? string.Empty;
            if (txtScriptsPath != null) txtScriptsPath.Text = _settings.DefaultScriptsPath ?? string.Empty;
            if (txtFediverseUrl != null) txtFediverseUrl.Text = _settings.CustomFediverseUrl ?? "https://mastodon.social";

            modulesPanel.Children.Clear();
            foreach (var tag in AllTags)
            {
                if (string.Equals(tag, "Settings", StringComparison.OrdinalIgnoreCase)) continue;

                var grid = new Grid { Margin = new Thickness(0,2,0,2), VerticalAlignment = VerticalAlignment.Center };
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(ModuleLabelColumnWidth) });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                var cb = new CheckBox
                {
                    Content = tag,
                    IsChecked = !_settings.DisabledModules.Contains(tag, StringComparer.OrdinalIgnoreCase),
                    VerticalAlignment = VerticalAlignment.Center
                };
                Grid.SetColumn(cb, 0);
                grid.Children.Add(cb);

                if (WebModules.Contains(tag) && !DynamicSelectionTags.Contains(tag) && !NoCustomUrlTags.Contains(tag))
                {
                    var btn = new Button
                    {
                        Tag = tag,
                        Width = 18,
                        Height = 18,
                        Margin = new Thickness(4,0,0,0),
                        Padding = new Thickness(0),
                        ToolTip = $"URL für '{tag}' ändern",
                        HorizontalContentAlignment = HorizontalAlignment.Center,
                        VerticalContentAlignment = VerticalAlignment.Center
                    };
                    btn.Content = new TextBlock
                    {
                        Text = "\uE70F",
                        FontFamily = new FontFamily("Segoe MDL2 Assets"),
                        FontSize = 11,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center,
                        Margin = new Thickness(0,-1,0,0)
                    };
                    btn.Click += EditUrl_Click;
                    Grid.SetColumn(btn, 1);
                    grid.Children.Add(btn);
                }
                modulesPanel.Children.Add(grid);
            } 

            // Add programs manage button
            var progBtn = new Button
            {
                Content = "Programme verwalten",
                Margin = new Thickness(0,12,0,0),
                HorizontalAlignment = HorizontalAlignment.Left
            };
            progBtn.Click += (s, e) => { var dlg = new ProgramsDialog(_settings); dlg.Owner = Application.Current.MainWindow; dlg.ShowDialog(); };
            modulesPanel.Children.Add(progBtn);

            UpdateShellUi();
        }

        private void UpdateShellUi()
        {
            try
            {
                var btnSet = this.FindName("btnSetShell") as System.Windows.Controls.Button;
                var btnRestore = this.FindName("btnRestoreShell") as System.Windows.Controls.Button;
                var txtStatus = this.FindName("txtShellStatus") as System.Windows.Controls.TextBlock;
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(WinLogonKeyPath, false))
                {
                    var current = key?.GetValue("Shell") as string;
                    if (string.IsNullOrWhiteSpace(_originalShellCached))
                        _originalShellCached = current;
                    var exe = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;
                    bool isShell = string.Equals(current?.Trim('"'), exe, StringComparison.OrdinalIgnoreCase);
                    if (btnSet != null) btnSet.Visibility = isShell ? Visibility.Collapsed : Visibility.Visible;
                    if (btnRestore != null) btnRestore.Visibility = isShell ? Visibility.Visible : Visibility.Collapsed;
                    if (txtStatus != null) txtStatus.Text = isShell ? "Aktuell als Shell gesetzt" : "Aktuelle Shell: " + (current ?? "(Standard)");
                }
            }
            catch { }
        }
                           
        private void btnSetShell_Click(object sender, RoutedEventArgs e)
        {

            String text = File.ReadAllText(".\\shell.reg");
            File.WriteAllText(".\\shell_tmp.reg", text.Replace("explorer.exe", Process.GetCurrentProcess().MainModule.FileName.Replace("\\","\\\\")));            
            var psi = new ProcessStartInfo("regedit.exe","\""+Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName)+ "\\shell_tmp.reg\"")
            {
                UseShellExecute = true,
                Verb = "runas"
            };
            Process.Start(psi); ;
        }

        private void btnRestoreShell_Click(object sender, RoutedEventArgs e)
        {                    
            var psi = new ProcessStartInfo("regedit.exe", "\"" + Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName) + ".\\shell.reg\"")
            {
                UseShellExecute = true,
                Verb = "runas"
            };
            Process.Start(psi);
        }

        private void RunElevatedShellSetter(string newShell, bool saveOriginal, bool restore = false)
        {
            try
            {
                var exe = Process.GetCurrentProcess().MainModule.FileName;
                var args = restore ? "--restore-shell" : ($"--set-shell \"{newShell}\"");
                var psi = new ProcessStartInfo(exe, args)
                {
                    UseShellExecute = true,
                    Verb = "runas"
                };
                Process.Start(psi);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler bei Shell Änderung: " + ex.Message);
            }
        }

        private void EditUrl_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button b) || b.Tag == null) return;
            string tag = b.Tag.ToString();
            // Social module Fediverse override uses separate textbox; ignore here
            if (string.Equals(tag, "Social", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("Fediverse URL bitte im Feld unter News konfigurieren.");
                return;
            }
            var existing = _settings.CustomWebModuleUrls.FirstOrDefault(u => string.Equals(u.Tag, tag, StringComparison.OrdinalIgnoreCase));
            string currentUrl = existing?.Url ?? string.Empty;
            string input = PromptForUrl(tag, currentUrl);
            if (string.IsNullOrWhiteSpace(input)) return;
            if (!input.StartsWith("http", StringComparison.OrdinalIgnoreCase)) { MessageBox.Show("Bitte eine gültige URL (http/https) eingeben."); return; }
            if (existing == null) _settings.CustomWebModuleUrls.Add(new CustomWebModuleUrl { Tag = tag, Url = input.Trim() }); else existing.Url = input.Trim();
            _settings.Save();
        }

        private string PromptForUrl(string tag, string current)
        {
            var win = new Window { Title = $"URL für Modul '{tag}'", SizeToContent = SizeToContent.WidthAndHeight, WindowStartupLocation = WindowStartupLocation.CenterOwner, ResizeMode = ResizeMode.NoResize, Owner = Application.Current?.MainWindow };
            var stack = new StackPanel { Margin = new Thickness(12), MinWidth = 420 };
            stack.Children.Add(new TextBlock { Text = "Bitte URL eingeben (http/https):", Margin = new Thickness(0,0,0,6) });
            var tb = new TextBox { Text = current ?? string.Empty, Margin = new Thickness(0,0,0,12) };
            stack.Children.Add(tb);
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var ok = new Button { Content = "OK", Width = 80, Margin = new Thickness(0,0,8,0), IsDefault = true };
            var cancel = new Button { Content = "Abbrechen", Width = 80, IsCancel = true };
            string result = null;
            ok.Click += (_, __) => { result = tb.Text; win.DialogResult = true; };
            cancel.Click += (_, __) => { win.DialogResult = false; };
            buttons.Children.Add(ok); buttons.Children.Add(cancel); stack.Children.Add(buttons);
            win.Content = stack; win.ShowDialog();
            return result;
        }

        private bool IsRssSelected() => (newsModeSelector.SelectedItem as ComboBoxItem)?.Content?.ToString().Equals("RSS", StringComparison.OrdinalIgnoreCase) == true;

        private void SelectComboItem(ComboBox combo, string value, string fallback)
        {
            if (combo == null) return; string target = string.IsNullOrWhiteSpace(value) ? fallback : value; combo.SelectedIndex = -1;
            foreach (var item in combo.Items.OfType<ComboBoxItem>()) { if (string.Equals(item.Content?.ToString(), target, StringComparison.OrdinalIgnoreCase)) { combo.SelectedItem = item; return; } }
            if (combo.Items.Count > 0) combo.SelectedIndex = 0;
        }

        private void btnPickLocalPath_Click(object sender, RoutedEventArgs e) { var p = PickFolder(); if (!string.IsNullOrEmpty(p)) txtLocalPath.Text = p; }
        private void btnPickScriptsPath_Click(object sender, RoutedEventArgs e) { var p = PickFolder(); if (!string.IsNullOrEmpty(p)) txtScriptsPath.Text = p; }

        private string PickFolder()
        {
            var dlg = new OpenFileDialog { CheckFileExists = false, CheckPathExists = true, ValidateNames = false, FileName = "Ordner auswaehlen" }; var res = dlg.ShowDialog(); if (res == true) { try { return Path.GetDirectoryName(dlg.FileName); } catch { } } return null;
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            _settings.AiService = (aiServiceSelector.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? _settings.AiService;
            _settings.MessengerService = (messengerServiceSelector.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? _settings.MessengerService;
            _settings.ChatService = (chatServiceSelector.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? _settings.ChatService;
            _settings.NewsMode = (newsModeSelector.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? _settings.NewsMode;
            if (txtFediverseUrl != null && !string.IsNullOrWhiteSpace(txtFediverseUrl.Text) && txtFediverseUrl.Text.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                _settings.CustomFediverseUrl = txtFediverseUrl.Text.Trim();
            int max; if (txtRssMax != null && int.TryParse(txtRssMax.Text, out max) && max > 0) _settings.RssMaxArticles = max; else if (_settings.RssMaxArticles <= 0) _settings.RssMaxArticles = 60;
            _settings.DefaultLocalPath = txtLocalPath.Text ?? string.Empty; _settings.DefaultScriptsPath = txtScriptsPath.Text ?? string.Empty;
            var enabled = new List<string>(); foreach (var child in modulesPanel.Children) if (child is Grid g) { var chk = g.Children.OfType<CheckBox>().FirstOrDefault(); if (chk!=null && chk.IsChecked==true) enabled.Add(chk.Content.ToString()); }
            var disabled = AllTags.Where(t => !enabled.Contains(t, StringComparer.OrdinalIgnoreCase)).ToList(); disabled.RemoveAll(t => string.Equals(t, "Settings", StringComparison.OrdinalIgnoreCase)); _settings.DisabledModules = disabled; _settings.Save(); MessageBox.Show("Einstellungen gespeichert. Neustart empfohlen.");
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e) => Init();
        private void newsModeSelector_SelectionChanged(object sender, SelectionChangedEventArgs e) => btnEditFeeds.Visibility = IsRssSelected() ? Visibility.Visible : Visibility.Collapsed;
        private void btnEditFeeds_Click(object sender, RoutedEventArgs e)
        {
            var current = string.Join(Environment.NewLine, _settings.RssFeeds ?? new List<string>()); var dlg = new RssFeedsDialog(current);
            if (dlg.ShowDialog() == true)
            {
                var lines = (dlg.FeedsText ?? string.Empty).Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(l => l.Trim()).Where(l => l.StartsWith("http", StringComparison.OrdinalIgnoreCase)).Distinct(StringComparer.OrdinalIgnoreCase).ToList(); _settings.RssFeeds = lines; _settings.Save();
            }
        }
    }
}
