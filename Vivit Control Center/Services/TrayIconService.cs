using System;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;
using System.IO;
using System.Drawing; // added for SystemIcons
using Application = System.Windows.Application;

namespace Vivit_Control_Center.Services
{
    public sealed class TrayIconService : IDisposable
    {
        private static readonly Lazy<TrayIconService> _lazy = new Lazy<TrayIconService>(() => new TrayIconService());
        public static TrayIconService Current => _lazy.Value;

        private NotifyIcon _notifyIcon;
        private bool _initialized;
        private TaskbarWatcherWindow _watcher;
        private string _currentTheme = "Dark";
        private System.Drawing.Icon _iconDark;
        private System.Drawing.Icon _iconLight;
        private ContextMenuStrip _menu;
        private DateTime _lastBalloon = DateTime.MinValue;

        private TrayIconService() { }

        public void Initialize(string theme = null)
        {
            if (_initialized) return;

            LoadIcons();
            _currentTheme = string.IsNullOrWhiteSpace(theme) ? _currentTheme : theme;

            _notifyIcon = new NotifyIcon
            {
                Text = "Vivit Control Center",
                Visible = true,
                Icon = SelectIconForTheme(_currentTheme)
            };

            _menu = new ContextMenuStrip();
            _menu.Opening += Menu_Opening; // dynamischer Aufbau
            BuildMenu();
            _notifyIcon.ContextMenuStrip = _menu;

            _notifyIcon.DoubleClick += (_, __) => ShowMainWindow();

            _watcher = new TaskbarWatcherWindow();
            _watcher.TaskbarRestarted += (_, __) => RefreshIcon();

            Application.Current.Exit += (_, __) => Dispose();

            ShowBalloonSafe("Gestartet", "Vivit Control Center läuft im Hintergrund.");
            _initialized = true;
        }

        private void LoadIcons()
        {
            try
            {
                var baseDir = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                var darkPath = Path.Combine(baseDir, "tray_dark.ico");
                var lightPath = Path.Combine(baseDir, "tray_light.ico");
                if (File.Exists(darkPath)) _iconDark = new System.Drawing.Icon(darkPath);
                if (File.Exists(lightPath)) _iconLight = new System.Drawing.Icon(lightPath);
            }
            catch { }
        }

        private System.Drawing.Icon SelectIconForTheme(string theme)
        {
            bool light = string.Equals(theme, "Light", StringComparison.OrdinalIgnoreCase);
            if (light && _iconLight != null) return _iconLight;
            if (!light && _iconDark != null) return _iconDark;
            try
            {
                var assoc = System.Drawing.Icon.ExtractAssociatedIcon(Assembly.GetEntryAssembly().Location);
                if (assoc != null) return assoc;
            }
            catch { }
            return SystemIcons.Application; // fixed qualification
        }

        private void Menu_Opening(object sender, CancelEventArgs e)
        {
            try { BuildMenu(); } catch { }
        }

        private void BuildMenu()
        {
            if (_menu == null) return;
            _menu.Items.Clear();
            _menu.Items.Add("Öffnen", null, (_, __) => ShowMainWindow());
            _menu.Items.Add("Minimieren", null, (_, __) => MinimizeMainWindow());
            _menu.Items.Add("Neu laden", null, (_, __) => RefreshIcon());

            // Theme Toggle
            var nextTheme = string.Equals(_currentTheme, "Light", StringComparison.OrdinalIgnoreCase) ? "Dark" : "Light";
            _menu.Items.Add($"Theme: {nextTheme}", null, (_, __) => UpdateTheme(nextTheme));

            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add("Beenden", null, (_, __) => ExitApplication());
        }

        public void UpdateTheme(string theme)
        {
            try
            {
                _currentTheme = theme ?? _currentTheme;
                if (_notifyIcon != null)
                {
                    _notifyIcon.Icon = SelectIconForTheme(_currentTheme);
                    RefreshIcon();
                    ShowBalloonSafe("Theme", $"Theme gewechselt zu {_currentTheme}.");
                }
            }
            catch { }
        }

        private void MinimizeMainWindow()
        {
            try
            {
                if (Application.Current?.MainWindow == null) return;
                Application.Current.MainWindow.WindowState = WindowState.Minimized;
                Application.Current.MainWindow.Hide();
            }
            catch { }
        }

        private void ShowMainWindow()
        {
            if (Application.Current?.MainWindow == null) return;
            var w = Application.Current.MainWindow;
            if (w.WindowState == WindowState.Minimized) w.WindowState = WindowState.Normal;
            w.Show();
            w.Activate();
        }

        private void RefreshIcon()
        {
            if (_notifyIcon == null) return;
            try
            {
                _notifyIcon.Visible = false;
                _notifyIcon.Visible = true;
            }
            catch { }
        }

        private void ExitApplication()
        {
            try { Application.Current?.Shutdown(); } catch { }
        }

        public void ShowBalloonSafe(string title, string text, int timeoutMs = 3000, ToolTipIcon icon = ToolTipIcon.Info)
        {
            try
            {
                if (_notifyIcon == null) return;
                // Rate limiting (Explorer Neustart kann mehrfach auslösen)
                if ((DateTime.UtcNow - _lastBalloon).TotalSeconds < 2) return;
                _lastBalloon = DateTime.UtcNow;
                _notifyIcon.BalloonTipTitle = title;
                _notifyIcon.BalloonTipText = text;
                _notifyIcon.BalloonTipIcon = icon;
                _notifyIcon.ShowBalloonTip(timeoutMs);
            }
            catch { }
        }

        public void Dispose()
        {
            try
            {
                if (_notifyIcon != null)
                {
                    _notifyIcon.Visible = false;
                    _notifyIcon.Dispose();
                    _notifyIcon = null;
                }
                _watcher?.Dispose();
                _watcher = null;
                _iconDark?.Dispose();
                _iconLight?.Dispose();
            }
            catch { }
        }

        private class TaskbarWatcherWindow : NativeWindow, IDisposable
        {
            private static readonly int WM_TASKBARCREATED = RegisterWindowMessage("TaskbarCreated");
            private const int WS_OVERLAPPED = unchecked((int)0x00000000);

            public event EventHandler TaskbarRestarted;

            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            private static extern int RegisterWindowMessage(string lpString);

            public TaskbarWatcherWindow()
            {
                var cp = new CreateParams
                {
                    Caption = "TrayWatchHidden",
                    X = 0,
                    Y = 0,
                    Height = 0,
                    Width = 0,
                    Style = WS_OVERLAPPED
                };
                CreateHandle(cp);
            }

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WM_TASKBARCREATED)
                {
                    TaskbarRestarted?.Invoke(this, EventArgs.Empty);
                }
                base.WndProc(ref m);
            }

            public void Dispose()
            {
                try { DestroyHandle(); } catch { }
            }
        }
    }
}