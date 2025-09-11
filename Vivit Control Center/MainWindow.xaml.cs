using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using Vivit_Control_Center.Views;
using Vivit_Control_Center.Views.Modules;
using Vivit_Control_Center.Settings;

namespace Vivit_Control_Center
{
    public partial class MainWindow : Window
    {
        private AppSettings _settings;

        private readonly Dictionary<string, IModule> _modules =
            new Dictionary<string, IModule>(StringComparer.OrdinalIgnoreCase);

        private static readonly string[] Tags = new[]
        {
            "AI","News","Messenger","Chat","Explorer","Office","Notes","Media Player","Steam",
            "Webbrowser","Order Food","eBay","Temu","Terminal","Scripting","SSH","SFTP","Settings"
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

        public MainWindow()
        {
            InitializeComponent();
            _settings = AppSettings.Load();
            ApplyTheme(_settings.Theme);
            FilterSidebarBySettings();

            SidebarRoot.IsEnabled = false;
            UpdateLoadProgress(0);
            PrecreateModules();
            Loaded += async (_, __) => await PreloadAllAndHideSplashAsync();
        }

        private void ApplyTheme(string theme)
        {
            if (string.Equals(theme, "Light", StringComparison.OrdinalIgnoreCase))
                this.Background = System.Windows.Media.Brushes.WhiteSmoke;
            else
                this.Background = (System.Windows.Media.Brush) new System.Windows.Media.BrushConverter().ConvertFromString("#121212");
        }

        private void FilterSidebarBySettings()
        {
            try
            {
                var disabled = new HashSet<string>(_settings.DisabledModules ?? new List<string>(), StringComparer.OrdinalIgnoreCase);
                disabled.Remove("Settings");

                var sv = SidebarRoot.Child as ScrollViewer;
                var stack = sv?.Content as StackPanel;
                if (stack == null) return;
                foreach (var child in stack.Children.OfType<Button>())
                {
                    if (child.Tag is string tag && disabled.Contains(tag))
                        child.Visibility = Visibility.Collapsed;
                    else
                        child.Visibility = Visibility.Visible;
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
                progressText.Text = $"{Math.Min(loaded, total)} von {total} Modulen geladen";
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
            var uniqueTags = new HashSet<string>(Tags, StringComparer.OrdinalIgnoreCase);
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
            ShowModule(DefaultTag);
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
            if (!_modules.TryGetValue(tag, out var module))
            {
                tag = DefaultTag;
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
    }
}