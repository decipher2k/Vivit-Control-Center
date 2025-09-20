using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls; // WPF Button
using System.Windows.Threading;
using Vivit_Control_Center.Views;
using Vivit_Control_Center.Views.Modules;
using Vivit_Control_Center.Settings;
using System.Windows.Input;
using System.Diagnostics;
using Microsoft.Win32; // autorun registry
using System.IO; // path handling
using System.ComponentModel;
using Vivit_Control_Center.Localization;
using System.Windows.Media; // VisualTreeHelper

namespace Vivit_Control_Center
{
    public partial class MainWindow : Window
    {
        private AppSettings _settings;

        private readonly Dictionary<string, IModule> _modules =
            new Dictionary<string, IModule>(StringComparer.OrdinalIgnoreCase);

        private static readonly string[] Tags = new[]
        {
            "AI","News","Messenger","Chat","Explorer","Office","Notes","Email","Media Player","Steam",
            "Webbrowser","Order Food","eBay","Temu","Terminal","Scripting","SSH","SFTP","Social","Launch","Programs","Settings"
        };

        private const string DefaultTag = "Webbrowser";
        private bool _splashHidden;
        private bool _preloadingStarted;
        private int _modulesLoaded = 0;

        private readonly HashSet<string> _counted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _preloadInvoked = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly TaskCompletionSource<bool> _allModulesLoadedTcs =
            new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

        private readonly HashSet<string> _pending = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private const double ShellTaskbarHeight = 40; // keep in sync with TaskbarWindow height (fallback)

        // Track disabled/enabled tags for this session
        private HashSet<string> _disabledTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private List<string> _enabledTags = new List<string>();
        private string _startupTag;

        public MainWindow()
        {
            InitializeComponent();

            _settings = AppSettings.Load();
            ApplyTheme(_settings.Theme);
            ComputeEnabledTags();
            FilterSidebarBySettings();

            SidebarRoot.IsEnabled = false;
            UpdateLoadProgress(0);
            PrecreateModules();
            Loaded += async (_, __) => await PreloadAllAndHideSplashAsync();
            Loaded += (_, __) =>
            {
                ApplyWorkAreaConstraints();
                ClampToWorkArea();
                FilterSidebarBySettings();
                if (App.IsShellMode)
                {
                    try
                    {
                        // Collapse the title-bar row completely in shell mode
                        var titleRow = this.FindName("TitleRow") as RowDefinition;
                        if (titleRow != null) titleRow.Height = new GridLength(0);
                        HideTitleBarButtonsSafe();

                        // NICHT maximieren – stattdessen selbst auf Bildschirmgröße setzen,
                        // damit unser Fenster auch die vom OS reservierte Sidebar-Fläche links abdecken kann.
                        WindowState = WindowState.Normal;
                        FitMainWindowToScreen();
                    }
                    catch { }
                }
            };
            StateChanged += (_, __) => { UpdateMaxRestoreIcon(); ApplyWorkAreaConstraints(); };
        }

        private void HideTitleBarButtonsSafe()
        {
            try
            {
                (FindName("MinButton") as Button)?.SetValue(VisibilityProperty, Visibility.Collapsed);
                (FindName("MaxButton") as Button)?.SetValue(VisibilityProperty, Visibility.Collapsed);
                (FindName("CloseButton") as Button)?.SetValue(VisibilityProperty, Visibility.Collapsed);
            }
            catch { }
        }

        private void FitMainWindowToScreen()
        {
            try
            {
                // DIPs: SystemParameters sind bereits in DIPs
                double screenW = SystemParameters.PrimaryScreenWidth;
                double screenH = SystemParameters.PrimaryScreenHeight;
                // Pixel->DIP Umrechnung für Taskbarhöhe
                var src = PresentationSource.FromVisual(this);
                double mFromDeviceY = src?.CompositionTarget?.TransformFromDevice.M22 ?? 1.0;
                double taskbarDip = App.ShellTaskbarHeightPx * mFromDeviceY;
                if (double.IsNaN(taskbarDip) || taskbarDip <= 0) taskbarDip = 40;

                Left = 0;
                Top = 0;
                Width = screenW;
                Height = Math.Max(200, screenH - taskbarDip);
            }
            catch { }
        }

        private void ComputeEnabledTags()
        {
            try
            {
                _disabledTags = new HashSet<string>(_settings.DisabledModules ?? new List<string>(), StringComparer.OrdinalIgnoreCase);
                // Settings must always be available
                _disabledTags.Remove("Settings");
                // Build enabled list keeping the order from Tags array
                _enabledTags = Tags.Where(t => !_disabledTags.Contains(t)).ToList();
                if (_enabledTags.Count == 0)
                {
                    // Fallback: ensure at least Settings exists
                    _enabledTags.Add("Settings");
                }
                // Choose startup tag: DefaultTag if enabled else first enabled
                _startupTag = _enabledTags.Contains(DefaultTag, StringComparer.OrdinalIgnoreCase)
                    ? DefaultTag
                    : _enabledTags.FirstOrDefault() ?? "Settings";
            }
            catch
            {
                _disabledTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                _enabledTags = Tags.ToList();
                _startupTag = DefaultTag;
            }
        }

        // Lädt AutoRun Anwendungen (Run-Keys + Startup Folder) nur im Shell Mode
        private async Task LoadShellAutoRunsAsync()
        {
            if (!App.IsShellMode) return;
            try
            {
                await Task.Delay(500); // kleinen Puffer für Stabilität
                var entries = new List<(string exe, string args)>();
                void AddFromRegistry(RegistryKey root, string subKey)
                {
                    try
                    {
                        using (var k = root.OpenSubKey(subKey, false))
                        {
                            if (k == null) return;
                            foreach (var name in k.GetValueNames())
                            {
                                var raw = k.GetValue(name) as string;
                                if (string.IsNullOrWhiteSpace(raw)) continue;
                                ParseCommand(raw, out var path, out var args);
                                if (IsExecutableEligible(path))
                                    entries.Add((path, args));
                            }
                        }
                    }
                    catch { }
                }

                AddFromRegistry(Registry.CurrentUser, "Software\\Microsoft\\Windows\\CurrentVersion\\Run");
                AddFromRegistry(Registry.LocalMachine, "Software\\Microsoft\\Windows\\CurrentVersion\\Run");

                // Startup Folder (user + common)
                void AddFromFolder(Environment.SpecialFolder sf)
                {
                    try
                    {
                        var dir = Environment.GetFolderPath(sf);
                        if (Directory.Exists(dir))
                        {
                            foreach (var file in Directory.EnumerateFiles(dir))
                            {
                                var ext = Path.GetExtension(file)?.ToLowerInvariant();
                                if (ext == ".lnk" || ext == ".exe" || ext == ".bat" || ext == ".cmd")
                                {
                                    // For .lnk we try resolve; quick fallback just treat as shell execute
                                    if (ext == ".lnk")
                                    {
                                        // simplest: let shell handle .lnk via Process.Start(file)
                                        entries.Add((file, "")); 
                                    }
                                    else if (IsExecutableEligible(file))
                                        entries.Add((file, "")); 
                                }
                            }
                        }
                    }
                    catch { }
                }
                AddFromFolder(Environment.SpecialFolder.Startup);
                AddFromFolder(Environment.SpecialFolder.CommonStartup);

                // Remove duplicates (by exe path + args)
                var distinct = entries
                    .GroupBy(e => (e.exe ?? "") + "|" + (e.args ?? ""), StringComparer.OrdinalIgnoreCase)
                    .Select(g => g.First())
                    .ToList();

                foreach (var entry in distinct)
                {
                    try
                    {
                        var psi = new ProcessStartInfo
                        {
                            FileName = entry.exe,
                            Arguments = entry.args,
                            UseShellExecute = true,
                            WorkingDirectory = SafeDir(entry.exe)
                        };
                        Process.Start(psi);
                        await Task.Delay(150); // throttle a bit
                    }
                    catch { }
                }
            }
            catch { }
        }

        private static string SafeDir(string path)
        {
            try { var d = Path.GetDirectoryName(path); return string.IsNullOrWhiteSpace(d) ? Environment.CurrentDirectory : d; } catch { return Environment.CurrentDirectory; }
        }

        private static bool IsExecutableEligible(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path)) return false;
                path = Environment.ExpandEnvironmentVariables(path).Trim().Trim('"');
                if (!File.Exists(path)) return false;
                var exeSelf = Process.GetCurrentProcess().MainModule.FileName;
                if (string.Equals(path, exeSelf, StringComparison.OrdinalIgnoreCase)) return false; // skip self
                if (string.Equals(Path.GetFileName(path), "explorer.exe", StringComparison.OrdinalIgnoreCase)) return false; // don't auto-start explorer in shell mode
                return true;
            }
            catch { return false; }
        }

        private static void ParseCommand(string raw, out string path, out string args)
        {
            path = null; args = "";
            try
            {
                raw = Environment.ExpandEnvironmentVariables(raw).Trim();
                if (raw.StartsWith("\""))
                {
                    int end = raw.IndexOf('"', 1);
                    if (end > 1)
                    {
                        path = raw.Substring(1, end - 1);
                        args = raw.Substring(end + 1).Trim();
                        return;
                    }
                }
                // no quotes: split on first space if file exists
                var firstSpace = raw.IndexOf(' ');
                if (firstSpace > 0)
                {
                    var potential = raw.Substring(0, firstSpace).Trim();
                    if (File.Exists(potential))
                    {
                        path = potential;
                        args = raw.Substring(firstSpace + 1).Trim();
                        return;
                    }
                }
                // fallback entire string as path
                path = raw;
            }
            catch { path = null; args = ""; }
        }

        // Public helper for external activation (e.g., taskbar button)
        public void ActivateModule(string tag)
        {
            try
            {
                ShowModule(tag);
                if (WindowState == WindowState.Minimized)
                    WindowState = WindowState.Normal;
                Activate();
                Focus();
            }
            catch { }
        }

        private void AdjustShellWorkspaceArea()
        {
            try
            {
                // Position full width, exclude taskbar height at bottom
                Left = 0;
                Top = 0;
                Width = SystemParameters.PrimaryScreenWidth;
                var fullHeight = SystemParameters.PrimaryScreenHeight;
                var taskbarDip = App.ShellTaskbarHeightPx * (PresentationSource.FromVisual(this)?.CompositionTarget?.TransformFromDevice.M22 ?? 1.0);
                var targetHeight = Math.Max(300, fullHeight - (double.IsNaN(taskbarDip) ? ShellTaskbarHeight : taskbarDip));
                Height = targetHeight;
                MaxHeight = targetHeight; // prevent user resize overlapping taskbar
            }
            catch { }
        }

        private void UpdateMaxRestoreIcon()
        {
            try
            {
                var maxIcon = this.FindName("MaxIcon") as TextBlock;
                if (maxIcon != null)
                {
                    if (WindowState == WindowState.Maximized)
                        maxIcon.Text = "❐"; // restore icon
                    else
                        maxIcon.Text = "□"; // maximize icon
                }
            }
            catch { }
        }

        private void ApplyTheme(string theme)
        {
            if (string.Equals(theme, "Light", StringComparison.OrdinalIgnoreCase))
                this.Background = System.Windows.Media.Brushes.WhiteSmoke;
            else
                this.Background = (System.Windows.Media.Brush) new System.Windows.Media.BrushConverter().ConvertFromString("#121212");
        }

        private IEnumerable<Button> EnumerateSidebarButtons()
        {
            foreach (var btn in FindVisualChildren<Button>(SidebarRoot))
            {
                if (btn.Tag is string) yield return btn;
            }
        }

        private IEnumerable<T> FindVisualChildren<T>(DependencyObject root) where T : DependencyObject
        {
            if (root == null) yield break;
            int count = VisualTreeHelper.GetChildrenCount(root);
            for (int i = 0; i < count; i++)
            {
                var child = VisualTreeHelper.GetChild(root, i);
                if (child is T t) yield return t;
                foreach (var g in FindVisualChildren<T>(child)) yield return g;
            }
        }

        private void FilterSidebarBySettings()
        {
            try
            {
                var disabled = new HashSet<string>(_settings.DisabledModules ?? new List<string>(), StringComparer.OrdinalIgnoreCase);
                disabled.Remove("Settings");

                foreach (var btn in EnumerateSidebarButtons())
                {
                    var tag = btn.Tag as string;
                    if (string.IsNullOrWhiteSpace(tag)) continue;
                    btn.Visibility = disabled.Contains(tag) ? Visibility.Collapsed : Visibility.Visible;
                }
            }
            catch { }
        }

        private void UpdateLoadProgress(int loaded)
        {
            var progressBar = FindName("LoadProgressBar") as System.Windows.Controls.ProgressBar;
            var progressText = FindName("LoadProgressText") as TextBlock;
            var pendingList = FindName("PendingList") as ItemsControl;

            int total = _modules.Count;

            if (progressBar != null)
            {
                progressBar.Maximum = total;
                progressBar.Value = Math.Min(loaded, total);
            }
            if (progressText != null)
            {
                var fmt = LocalizationManager.GetString("App.ModulesLoadedFormat", "{0} von {1} Modulen geladen");
                progressText.Text = string.Format(fmt, Math.Min(loaded, total), total);
            }
            if (pendingList != null)
            {
                pendingList.ItemsSource = _pending.OrderBy(x => x).ToList();
            }

            if (loaded >= total && !_allModulesLoadedTcs.Task.IsCompleted)
                _allModulesLoadedTcs.TrySetResult(true);
        }

        private void PrecreateModules()
        {
            // Only create enabled modules
            var uniqueTags = new HashSet<string>(_enabledTags, StringComparer.OrdinalIgnoreCase);
            foreach (var tag in uniqueTags)
            {
                _pending.Add(tag);
                var module = CreateModule(tag);
                module.View.Visibility = Visibility.Hidden;
                module.SetVisible(false);

                EventHandler handler = null;
                handler = (s, e) =>
                {
                    if (!_counted.Add(tag)) return;
                    if (s is IModule m) m.LoadCompleted -= handler;

                    Dispatcher.InvokeAsync(() =>
                    {
                        _pending.Remove(tag);
                        _modulesLoaded++;
                        UpdateLoadProgress(_modulesLoaded);
                    });
                };
                module.LoadCompleted += handler;

                _modules[tag] = module;
                ModuleHost.Children.Add(module.View);
            }

            UpdateLoadProgress(0);
        }

        private IModule CreateModule(string tag)
        {
            switch (tag)
            {
                case "Explorer":      return (IModule) new ExplorerModule();
                case "Media Player":  return (IModule) new MediaPlayerModule();
                case "Steam":         return (IModule) new SteamModule();
                case "Webbrowser":    return (IModule) new BrowserModule();
                case "Terminal":      return (IModule) new TerminalModule();
                case "Scripting":     return (IModule) new ScriptingModule();
                case "SSH":           return (IModule) new SshModule();
                case "SFTP":          return (IModule) new SftpModule();
                case "Settings":      return (IModule) new SettingsModule();
                case "Notes":         return (IModule) new NotesModule();
                case "Office":        return (IModule) new OfficeModule();
                case "Social":        return (IModule) new SocialModule();
                case "Launch":        return (IModule) new LaunchModule();
                case "Programs":      return (IModule) new LaunchModule(); // alias for Launch
                case "Email":         return (IModule) new EmailModule();
                case "News":
                    var settings = Settings.AppSettings.Load();
                    if (string.Equals(settings.NewsMode, "RSS", StringComparison.OrdinalIgnoreCase))
                        return new Views.Modules.RssNewsModule();
                    return (IModule) new WebViewModule(tag);
                default:               return (IModule) new WebViewModule(tag);
            }
        }

        private async Task PreloadAllAndHideSplashAsync()
        {
            if (_preloadingStarted) return;
            _preloadingStarted = true;

            await Dispatcher.Yield(DispatcherPriority.Background);

            var modules = _modules.Values.ToArray();
            foreach (var module in modules)
                _ = module.PreloadAsync();

            Task Wrap(Task t) => t.ContinueWith(_ => { }, TaskScheduler.Default);
            var allTasksSafe = Task.WhenAll(modules.Select(m => Wrap(m.LoadCompletedTask)));

            async Task WaitForCounterAsync(int total)
            {
                while (_modulesLoaded < total)
                    await Task.Delay(100).ConfigureAwait(false);
            }
            var waitCounter = WaitForCounterAsync(_modules.Count);

            var timeout = Task.Delay(TimeSpan.FromMinutes(1.5));
            var both = Task.WhenAll(allTasksSafe, waitCounter);
            var winner = await Task.WhenAny(both, timeout);

            SplashOverlay.Visibility = Visibility.Collapsed;
            _splashHidden = true;

            foreach (var m in modules)
                m.View.Visibility = Visibility.Collapsed;

            SidebarRoot.IsEnabled = true;
            ShowModule(_startupTag);
        }

        private void SidebarButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string tag && !string.IsNullOrWhiteSpace(tag))
            {
                if (_splashHidden && SplashOverlay != null)
                    SplashOverlay.Visibility = Visibility.Collapsed;

                ShowModule(tag);
            }
        }

        private void ShowModule(string tag)
        {
            if (string.IsNullOrWhiteSpace(tag) || !_modules.TryGetValue(tag, out var module))
            {
                // Fallback to startup tag if requested one is disabled/missing
                tag = _modules.ContainsKey(_startupTag) ? _startupTag : _modules.Keys.FirstOrDefault() ?? "Settings";
                module = _modules[tag];
            }

            foreach (var kv in _modules)
            {
                kv.Value.View.Visibility = Visibility.Collapsed;
                kv.Value.SetVisible(false);
            }

            module.View.Visibility = Visibility.Visible;
            module.SetVisible(true);
            this.Title = $"Vivit Control Center - {tag}";
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // Enable double-click maximize/restore and dragging in all modes
            if (e.ClickCount == 2)
            {
                ToggleMaximize();
                return;
            }
            try
            {
                DragMove();
            }
            catch { }
        }

        private void MinButton_Click(object sender, RoutedEventArgs e)
        {
            // Allow minimizing even in shell mode
            WindowState = WindowState.Minimized;
        }

        private void MaxButton_Click(object sender, RoutedEventArgs e)
        {
            ToggleMaximize();
        }

        private void ToggleMaximize()
        {
            if (WindowState == WindowState.Maximized)
                WindowState = WindowState.Normal;
            else
                WindowState = WindowState.Maximized;
            UpdateMaxRestoreIcon();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            if (App.IsShellMode)
            {
                var res = MessageBox.Show("Shell-Modus beenden und schließen? (Explorer wird nicht automatisch gestartet)", "Beenden", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (res != MessageBoxResult.Yes) return;
            }
            Close();
        }

        private void RebootButton_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("Möchten Sie den Computer wirklich neu starten?",
                "Neustart bestätigen", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                Process.Start("shutdown", "/r /t 0");
            }
        }

        private void ShutdownButton_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("Möchten Sie den Computer wirklich herunterfahren?",
                "Herunterfahren bestätigen", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                Process.Start("shutdown", "/s /t 0");
            }
        }

        private void RebootArea_Click(object sender, RoutedEventArgs e)
        {
            var result = System.Windows.MessageBox.Show("Aktion wählen:\nJa = Reboot\nNein = Shutdown\nAbbrechen = Cancel", "Power", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                try { Process.Start("shutdown", "/r /t 0"); } catch { }
            }
            else if (result == MessageBoxResult.No)
            {
                try { Process.Start("shutdown", "/s /t 0"); } catch { }
            }
        }

        protected override void OnStateChanged(EventArgs e)
        {
            base.OnStateChanged(e);
            // Do not hide window on minimize; let it minimize to the taskbar
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            // Wenn wirklich beendet werden soll: nicht abbrechen
            // Falls nur ins Tray: auskommentieren und stattdessen:
            // e.Cancel = true; Hide();
            base.OnClosing(e);
        }

        public void RestoreFromTray()
        {
            Show();
            if (WindowState == WindowState.Minimized) WindowState = WindowState.Normal;
            Activate();
        }

        // Ensure borderless window maximizes to the working area and not under the taskbar
        private void ApplyWorkAreaConstraints()
        {
            try
            {
                var wa = SystemParameters.WorkArea;
                if (WindowState == WindowState.Maximized)
                {
                    MaxWidth = wa.Width;
                    MaxHeight = wa.Height;
                    Left = wa.Left;
                    Top = wa.Top;
                }
                else
                {
                    // reset caps so the user can resize in normal state
                    MaxWidth = double.PositiveInfinity;
                    MaxHeight = double.PositiveInfinity;
                }
            }
            catch { }
        }

        // Prevent the window from starting off-screen
        private void ClampToWorkArea()
        {
            try
            {
                var wa = SystemParameters.WorkArea;
                if (double.IsNaN(Left) || double.IsNaN(Top)) return;
                if (Left < wa.Left) Left = wa.Left;
                if (Top < wa.Top) Top = wa.Top;
                if (Width > wa.Width) Width = wa.Width;
                if (Height > wa.Height) Height = wa.Height;
            }
            catch { }
        }

        // Obtain sidebar width (in DIPs) by measuring its ActualWidth and convert to pixels
        private void UpdateWorkAreaForSidebar()
        {
            try
            {
                if (!App.IsShellMode) return;
                if (SidebarRoot == null) return;

                // Sidebarbreite in DIPs -> in Pixel umrechnen
                double dipLeft = SidebarRoot.ActualWidth;
                var src = PresentationSource.FromVisual(this);
                double m11 = src?.CompositionTarget?.TransformToDevice.M11 ?? 1.0; // DIP->PX scale X
                int leftPx = (int)Math.Round(Math.Max(0, dipLeft) * m11);
                
                App.UpdateShellWorkAreaLeft(leftPx);
            }
            catch { }
        }

        // Apply sidebar tweaks when in shell mode:
        // - Remove Power button entirely
        // - Hide Launch (Start) button as well
        private void ApplyShellSidebarTweaks()
        {
            if (!App.IsShellMode) return;
            try
            {
                // Find the bottom StackPanel inside the sidebar (DockPanel.Dock == Bottom)
                var bottomPanel = FindVisualChildren<StackPanel>(SidebarRoot)
                    .FirstOrDefault(sp => DockPanel.GetDock(sp) == Dock.Bottom)
                    ?? FindVisualChildren<StackPanel>(SidebarRoot)
                        .FirstOrDefault(sp => sp.Children.OfType<Button>().Any(b => (b.Tag as string) == "__reboot" || (b.Tag as string) == "Launch"));
                if (bottomPanel == null) return;

                // Remove/collapse the power button completely
                var powerBtn = bottomPanel.Children.OfType<Button>().FirstOrDefault(b => (b.Tag as string) == "__reboot");
                if (powerBtn != null)
                {
                    bottomPanel.Children.Remove(powerBtn);
                }

                // Hide the Launch (Start) button too
                var launchBtn = bottomPanel.Children.OfType<Button>().FirstOrDefault(b => (b.Tag as string) == "Launch");
                if (launchBtn != null)
                {
                    launchBtn.Visibility = Visibility.Collapsed;
                    // If any context menu was previously attached, close it
                    try { launchBtn.ContextMenu = null; } catch { }
                }
            }
            catch { }
        }

        protected override void OnContentRendered(EventArgs e)
        {
            base.OnContentRendered(e);
            UpdateWorkAreaForSidebar();
            ApplyShellSidebarTweaks();
        }

        protected override void OnRenderSizeChanged(SizeChangedInfo sizeInfo)
        {
            base.OnRenderSizeChanged(sizeInfo);
            // If sidebar width changes (resize), update work area
            if (sizeInfo.WidthChanged)
            {
                UpdateWorkAreaForSidebar();
            }
        }

        // ...rest of class...
    }
}