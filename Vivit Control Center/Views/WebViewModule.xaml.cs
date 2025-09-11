using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Web.WebView2.Core;
using Vivit_Control_Center.Settings;

namespace Vivit_Control_Center.Views
{
    public partial class WebViewModule : UserControl, IModule
    {
        private readonly string _tag;
        private readonly string _url;

        private bool _initialized;
        private bool _navigated;
        private bool _eventsHooked;
        private bool _loadSignaled;
        private bool _preloadStarted; // Guard gegen mehrfaches Preload

        private readonly object _gate = new object();
        private CancellationTokenSource _stabilityCts;

        private readonly TaskCompletionSource<bool> _loadTcs =
            new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        public Task LoadCompletedTask => _loadTcs.Task;
        public event EventHandler LoadCompleted;

        public FrameworkElement View => this;

        private static readonly (string Tag, string Url)[] StaticRoutes = new[]
        {
            ("News", "https://news.google.com/?hl=de"),
            ("Media Player", "https://music.youtube.com/"),
            ("Order Food", "https://www.lieferando.de/"),
            ("eBay", "https://www.ebay.de/"),
            ("Temu", "https://www.temu.com/")
        };

        public WebViewModule(string tag)
        {
            InitializeComponent();
            _tag = tag ?? string.Empty;
            _url = ResolveUrl(_tag) ?? "about:blank";

            // WICHTIG: Während der Initialisierung sichtbar lassen, damit Handle/Controller erstellt wird
            Browser.Visibility = Visibility.Visible;
            Browser.IsHitTestVisible = false;
        }

        // Wird vom MainWindow beim Start für alle Module aufgerufen
        public async Task PreloadAsync()
        {
            if (_preloadStarted)
            {
                System.Diagnostics.Debug.WriteLine($"[{_tag}] PreloadAsync skipped (already started)");
                return;
            }
            _preloadStarted = true;

            System.Diagnostics.Debug.WriteLine($"[{_tag}] PreloadAsync started");

            await EnsureInitializedAsync();
            await NavigateOnceAsync();

            var moduleTimeout = Task.Delay(TimeSpan.FromSeconds(30));
            var completedTask = await Task.WhenAny(LoadCompletedTask, moduleTimeout);
            if (completedTask == moduleTimeout)
            {
                System.Diagnostics.Debug.WriteLine($"[{_tag}] Module load timeout after 30 seconds");
                SignalLoadCompleted();
            }
        }

        private async Task EnsureInitializedAsync()
        {
            if (_initialized) return;

            try
            {
                System.Diagnostics.Debug.WriteLine($"[{_tag}] Initializing WebView2...");
                await Browser.EnsureCoreWebView2Async(null);
                HookEventsOnce();
                _initialized = true;
                System.Diagnostics.Debug.WriteLine($"[{_tag}] WebView2 initialized successfully");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[{_tag}] WebView2 initialization failed: {ex.Message}");
                // Bei Initialisierungsfehler trotzdem als geladen markieren
                SignalLoadCompleted();
            }
        }

        private void HookEventsOnce()
        {
            if (_eventsHooked) return;
            _eventsHooked = true;

            Browser.NavigationStarting += Browser_NavigationStarting;
            Browser.NavigationCompleted += Browser_NavigationCompleted;
        }

        private void Browser_NavigationStarting(object sender, CoreWebView2NavigationStartingEventArgs e)
        {
            if (_loadSignaled) return;

            System.Diagnostics.Debug.WriteLine($"[{_tag}] Navigation starting to: {e.Uri}");
            // Jede neue Top-Level-Navigation bricht die Stabilitätswartezeit ab
            CancelStabilityTimer();
        }

        private void Browser_NavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            if (_loadSignaled) return;

            System.Diagnostics.Debug.WriteLine($"[{_tag}] Navigation completed. Success: {e.IsSuccess}");

            if (e.IsSuccess)
            {
                StartStabilityTimer();
            }
            else
            {
                SignalLoadCompleted();
            }
        }

        private Task NavigateOnceAsync()
        {
            if (_navigated) return Task.CompletedTask;

            try
            {
                System.Diagnostics.Debug.WriteLine($"[{_tag}] Navigating to: {_url}");
                Browser.Source = new Uri(_url);
                _navigated = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[{_tag}] Navigation failed: {ex.Message}");
                SignalLoadCompleted();
            }
            return Task.CompletedTask;
        }

        public void SetWebViewVisible(bool visible)
        {
            // WICHTIG: Vor Abschluss der Initialisierung nicht verstecken, sonst kann Handle fehlen
            if (!_initialized)
            {
                System.Diagnostics.Debug.WriteLine($"[{_tag}] SetWebViewVisible({visible}) ignored (not initialized)");
                return;
            }
            Browser.Visibility = visible ? Visibility.Visible : Visibility.Hidden;
            Browser.IsHitTestVisible = visible;
        }

        private static string ResolveUrl(string tag)
        {
            var settings = AppSettings.Load();

            if (string.Equals(tag, "AI", StringComparison.OrdinalIgnoreCase))
            {
                switch (settings.AiService)
                {
                    case "Perplexity AI": return "https://perplexity.ai";
                    case "Claude": return "https://claude.ai";
                    case "ChatGPT":
                    default: return "https://chatgpt.com";
                }
            }
            if (string.Equals(tag, "Messenger", StringComparison.OrdinalIgnoreCase))
            {
                switch (settings.MessengerService)
                {
                    case "Telegram": return "https://web.telegram.org";
                    case "WhatsApp":
                    default: return "https://web.whatsapp.com";
                }
            }
            if (string.Equals(tag, "Chat", StringComparison.OrdinalIgnoreCase))
            {
                switch (settings.ChatService)
                {
                    case "IRC IRCNet": return "https://webchat.ircnet.net";
                    case "IRC QuakeNet": return "https://webchat.quakenet.org";
                    case "Discord":
                    default: return "https://discord.com/channels/@me";
                }
            }

            var match = StaticRoutes.FirstOrDefault(r => string.Equals(r.Tag, tag, StringComparison.OrdinalIgnoreCase));
            return string.IsNullOrEmpty(match.Tag) ? null : match.Url;
        }

        private void StartStabilityTimer()
        {
            lock (_gate)
            {
                CancelStabilityTimer_NoLock();

                _stabilityCts = new CancellationTokenSource();
                var token = _stabilityCts.Token;

                Task.Run(async () =>
                {
                    try
                    {
                        await Task.Delay(1200, token);
                    }
                    catch (TaskCanceledException)
                    {
                        return;
                    }

                    SignalLoadCompleted();
                });
            }
        }

        private void SignalLoadCompleted()
        {
            lock (_gate)
            {
                if (_loadSignaled) return;

                _loadSignaled = true;
                _loadTcs.TrySetResult(true);

                System.Diagnostics.Debug.WriteLine($"[{_tag}] Module load completed (signaled)");

                Browser.NavigationStarting -= Browser_NavigationStarting;
                Browser.NavigationCompleted -= Browser_NavigationCompleted;

                CancelStabilityTimer_NoLock();
            }

            try { LoadCompleted?.Invoke(this, EventArgs.Empty); } catch { }
        }

        private void CancelStabilityTimer()
        {
            lock (_gate)
            {
                CancelStabilityTimer_NoLock();
            }
        }

        private void CancelStabilityTimer_NoLock()
        {
            if (_stabilityCts != null)
            {
                try { _stabilityCts.Cancel(); }
                catch { }
                _stabilityCts.Dispose();
                _stabilityCts = null;
            }
        }

        // Bestehender Code bleibt, nur diese Methode ergänzen:
        public void SetVisible(bool visible)
        {
            // Container und WebView synchron schalten
            Visibility = visible ? Visibility.Visible : Visibility.Hidden;
            IsHitTestVisible = visible;
            SetWebViewVisible(visible);
        }
    }
}