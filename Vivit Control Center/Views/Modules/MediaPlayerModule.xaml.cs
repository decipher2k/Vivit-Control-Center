using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using System.Diagnostics;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Web.WebView2.Core;
using System.Collections.Generic;
using System.Xml.Serialization;
using Vivit_Control_Center.Settings;
using System.Threading.Tasks;

namespace Vivit_Control_Center.Views.Modules
{
    public partial class MediaPlayerModule : BaseSimpleModule
    {
        private class Track : INotifyPropertyChanged
        {
            public string FilePath { get; set; }
            public string Url { get; set; }
            public string DisplayName { get; set; }
            public bool IsYouTube { get; set; }
            private TimeSpan _duration;
            public TimeSpan Duration { get => _duration; set { _duration = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Duration))); } }
            public event PropertyChangedEventHandler PropertyChanged;
        }

        private readonly ObservableCollection<Track> _playlist = new ObservableCollection<Track>();
        private int _currentIndex = -1;
        private bool _isDraggingPosition;
        private DispatcherTimer _timer;
        private bool _playRequested;
        private bool _isYouTubeActive;
        private bool _isYouTubePlaying;

        public MediaPlayerModule()
        {
            InitializeComponent();
            lstPlaylist.ItemsSource = _playlist;
            volSlider.Value = 0.5;
            _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(250) };
            _timer.Tick += (s, e) => UpdateClockAndPosition();
            SetPlayVisual(false);

            // Load last playlist from settings on module creation
            try { LoadLastPlaylistFromSettings(); } catch { }
        }

        // UI Handlers
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Audio Files|*.mp3;*.wav;*.wma;*.aac;*.m4a;*.ogg;*.flac|All Files|*.*",
                Multiselect = true
            };
            if (dlg.ShowDialog() == true)
            {
                foreach (var path in dlg.FileNames)
                {
                    try
                    {
                        _playlist.Add(new Track
                        {
                            FilePath = path,
                            DisplayName = System.IO.Path.GetFileNameWithoutExtension(path)
                        });
                    }
                    catch { }
                }
                if (_currentIndex < 0 && _playlist.Count > 0)
                {
                    _currentIndex = 0;
                    lstPlaylist.SelectedIndex = 0;
                    LoadCurrent(autoPlay: false);
                }
                PersistPlaylistToSettings();
            }
        }

        private void btnAddYouTube_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var url = PromptForText("Enter YouTube playlist URL (list=...):", "Add YouTube Playlist");
                if (string.IsNullOrWhiteSpace(url)) return;
                url = url.Replace("music.youtube.com", "www.youtube.com");
                AddYouTubePlaylist(url.Trim());
                PersistPlaylistToSettings();
            }
            catch { }
        }

        private void btnRemove_Click(object sender, RoutedEventArgs e)
        {
            var idx = lstPlaylist.SelectedIndex >= 0 ? lstPlaylist.SelectedIndex : _currentIndex;
            if (idx < 0 || idx >= _playlist.Count) return;
            var wasCurrent = idx == _currentIndex;
            _playlist.RemoveAt(idx);
            if (_playlist.Count == 0)
            {
                StopPlayback();
                _currentIndex = -1;
                txtNowPlaying.Text = "Now Playing: -";
                txtTime.Text = "00:00 / 00:00";
                PersistPlaylistToSettings();
                return;
            }
            if (wasCurrent)
            {
                _currentIndex = Math.Min(idx, _playlist.Count - 1);
                lstPlaylist.SelectedIndex = _currentIndex;
                LoadCurrent(autoPlay: false);
            }
            PersistPlaylistToSettings();
        }

        private void btnClear_Click(object sender, RoutedEventArgs e)
        {
            _playlist.Clear();
            _currentIndex = -1;
            StopPlayback();
            txtNowPlaying.Text = "Now Playing: -";
            txtTime.Text = "00:00 / 00:00";
            PersistPlaylistToSettings();
        }

        private void lstPlaylist_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (lstPlaylist.SelectedIndex >= 0)
            {
                _currentIndex = lstPlaylist.SelectedIndex;
                LoadCurrent(autoPlay: true);
                PersistPlaylistToSettings();
            }
        }

        private void btnPrev_Click(object sender, RoutedEventArgs e)
        {
            if (_playlist.Count == 0) return;
            _currentIndex = (_currentIndex <= 0) ? _playlist.Count - 1 : _currentIndex - 1;
            lstPlaylist.SelectedIndex = _currentIndex;
            LoadCurrent(autoPlay: true);
            PersistPlaylistToSettings();
        }

        private void btnNext_Click(object sender, RoutedEventArgs e)
        {
            if (_playlist.Count == 0) return;
            _currentIndex = (_currentIndex + 1) % _playlist.Count;
            lstPlaylist.SelectedIndex = _currentIndex;
            LoadCurrent(autoPlay: true);
            PersistPlaylistToSettings();
        }

        private async void btnPlayPause_Click(object sender, RoutedEventArgs e)
        {
            if (_currentIndex < 0)
            {
                if (_playlist.Count == 0) return;
                _currentIndex = 0;
                lstPlaylist.SelectedIndex = 0;
                LoadCurrent(autoPlay: true);
                PersistPlaylistToSettings();
                return;
            }

            var track = (_currentIndex >= 0 && _currentIndex < _playlist.Count) ? _playlist[_currentIndex] : null;
            if (track != null && track.IsYouTube)
            {
                if (!_isYouTubeActive)
                {
                    await StartYouTubeAsync(track, autoPlay: true);
                    _isYouTubePlaying = true;
                    SetPlayVisual(true);
                }
                else
                {
                    if (_isYouTubePlaying)
                    {
                        await TrySendYouTubeCommandAsync("pause");
                        _isYouTubePlaying = false;
                        SetPlayVisual(false);
                    }
                    else
                    {
                        await TrySendYouTubeCommandAsync("play");
                        _isYouTubePlaying = true;
                        SetPlayVisual(true);
                    }
                }
                return;
            }

            if (_playRequested || media.CanPause)
            {
                try { media.Pause(); } catch { }
                _playRequested = false;
                SetPlayVisual(false);
            }
            else
            {
                try { media.Play(); _timer.Start(); } catch { }
                _playRequested = true;
                SetPlayVisual(true);
            }
        }

        private void StopPlayback()
        {
            try { media.Stop(); } catch { }
            _playRequested = false;
            if (_isYouTubeActive)
            {
                var _ = TrySendYouTubeCommandAsync("stop");
                _isYouTubePlaying = false;
            }
            SetPlayVisual(false);
            _timer.Stop();
            UpdateClockAndPosition();
        }

        private void posSlider_PpreviewMouseDown(object sender, MouseButtonEventArgs e) { _isDraggingPosition = true; }
        private void posSlider_PreviewMouseDown(object sender, MouseButtonEventArgs e) { _isDraggingPosition = true; }
        private void posSlider_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            _isDraggingPosition = false;
            try
            {
                if (_isYouTubeActive) return;
                var total = media.NaturalDuration.HasTimeSpan ? media.NaturalDuration.TimeSpan.TotalSeconds : 0;
                if (total > 0)
                {
                    media.Position = TimeSpan.FromSeconds(total * (posSlider.Value / posSlider.Maximum));
                }
            }
            catch { }
        }
        private void posSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_isDraggingPosition || _isYouTubeActive) return;
        }

        private void volSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_isYouTubeActive) return; // not supported currently for YouTube
            try { media.Volume = volSlider.Value; } catch { }
        }

        // MediaElement events
        private void media_MediaOpened(object sender, RoutedEventArgs e)
        {
            try
            {
                if (media.NaturalDuration.HasTimeSpan)
                {
                    posSlider.Minimum = 0;
                    posSlider.Maximum = media.NaturalDuration.TimeSpan.TotalSeconds;
                }
                UpdateNowPlaying();
                _timer.Start();
                if (_playRequested)
                {
                    media.Play();
                    SetPlayVisual(true);
                }
                else
                {
                    SetPlayVisual(false);
                }
            }
            catch { }
        }

        private void media_MediaEnded(object sender, RoutedEventArgs e)
        {
            // Auto next
            btnNext_Click(sender, e);
        }

        private void media_MediaFailed(object sender, ExceptionRoutedEventArgs e)
        {
            // Skip to next on failure
            btnNext_Click(sender, null);
        }

        private void LoadCurrent(bool autoPlay)
        {
            if (_currentIndex < 0 || _currentIndex >= _playlist.Count) return;
            var track = _playlist[_currentIndex];
            try
            {
                if (track.IsYouTube && !string.IsNullOrWhiteSpace(track.Url))
                {
                    Dispatcher.BeginInvoke(new Action(async () =>
                    {
                        await StartYouTubeAsync(track, autoPlay);
                    }));
                    return;
                }

                // Local file
                ShowYouTubeUi(false);
                if (!string.IsNullOrEmpty(track.FilePath))
                {
                    if (!File.Exists(track.FilePath)) { btnNext_Click(this, null); return; }
                    media.Stop();
                    media.Source = new Uri(track.FilePath);
                    _playRequested = autoPlay;
                    SetPlayVisual(autoPlay);
                    txtNowPlaying.Text = $"Now Playing: {track.DisplayName}";
                    txtTime.Text = "00:00 / 00:00";
                    if (autoPlay) { media.Play(); _timer.Start(); }
                    posSlider.IsEnabled = true;
                    volSlider.IsEnabled = true;
                }
            }
            catch { }
        }

        private void UpdateNowPlaying()
        {
            if (_currentIndex >= 0 && _currentIndex < _playlist.Count)
            {
                var t = _playlist[_currentIndex];
                txtNowPlaying.Text = $"Now Playing: {t.DisplayName}";
            }
            else
            {
                txtNowPlaying.Text = "Now Playing: -";
            }
        }

        private void UpdateClockAndPosition()
        {
            try
            {
                if (_isYouTubeActive)
                {
                    txtTime.Text = "--:-- / --:--";
                    return;
                }
                if (media.NaturalDuration.HasTimeSpan)
                {
                    var total = media.NaturalDuration.TimeSpan;
                    var pos = media.Position;
                    if (!_isDraggingPosition)
                    {
                        posSlider.Maximum = total.TotalSeconds;
                        posSlider.Value = Math.Min(pos.TotalSeconds, total.TotalSeconds);
                    }
                    txtTime.Text = $"{FormatTime(pos)} / {FormatTime(total)}";
                }
                else
                {
                    txtTime.Text = "00:00 / 00:00";
                }
            }
            catch { }
        }

        private static string FormatTime(TimeSpan ts)
        {
            if (ts.TotalHours >= 1)
                return string.Format("{0:00}:{1:00}:{2:00}", (int)ts.TotalHours, ts.Minutes, ts.Seconds);
            return string.Format("{0:00}:{1:00}", (int)ts.TotalMinutes, ts.Seconds);
        }

        private void SetPlayVisual(bool isPlaying)
        {
            if (iconPlay == null || iconPause == null) return;
            iconPlay.Visibility = isPlaying ? Visibility.Collapsed : Visibility.Visible;
            iconPause.Visibility = isPlaying ? Visibility.Visible : Visibility.Collapsed;
        }

        // Helpers for YouTube playlist handling
        private void AddYouTubePlaylist(string playlistUrl)
        {
            try
            {
                var listId = ExtractPlaylistId(playlistUrl);
                if (string.IsNullOrEmpty(listId))
                {
                    MessageBox.Show("Invalid YouTube playlist URL.", "YouTube", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var videoUrls = FetchYouTubePlaylistVideoUrls(listId, playlistUrl);
                if (videoUrls == null || videoUrls.Length == 0)
                {
                    MessageBox.Show("No videos found or failed to retrieve playlist.", "YouTube", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                foreach (var vUrl in videoUrls)
                {
                    var title = TryGetYouTubeTitle(vUrl);
                    _playlist.Add(new Track
                    {
                        Url = vUrl,
                        IsYouTube = true,
                        DisplayName = string.IsNullOrEmpty(title) ? vUrl : title
                    });
                }

                if (_currentIndex < 0 && _playlist.Count > 0)
                {
                    _currentIndex = 0;
                    lstPlaylist.SelectedIndex = 0;
                    LoadCurrent(autoPlay: false);
                }
            }
            catch { }
        }

        private static string ExtractPlaylistId(string url)
        {
            try
            {
                var uri = new Uri(url);
                var list = GetQueryParam(uri, "list");
                if (!string.IsNullOrEmpty(list)) return list;

                if (!string.IsNullOrEmpty(uri.Fragment))
                {
                    var fragList = GetQueryParamFromFragment(uri.Fragment, "list");
                    if (!string.IsNullOrEmpty(fragList)) return fragList;
                }
            }
            catch { }
            return null;
        }

        private static string GetQueryParam(Uri uri, string key)
        {
            try
            {
                var q = uri.Query;
                if (string.IsNullOrEmpty(q)) return null;
                if (q.StartsWith("?")) q = q.Substring(1);
                var parts = q.Split(new[] { '&' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var p in parts)
                {
                    var kv = p.Split(new[] { '=' }, 2);
                    if (kv.Length == 2 && string.Equals(kv[0], key, StringComparison.OrdinalIgnoreCase))
                        return Uri.UnescapeDataString(kv[1]);
                }
            }
            catch { }
            return null;
        }

        private static string GetQueryParamFromFragment(string fragment, string key)
        {
            try
            {
                var f = fragment ?? string.Empty;
                f = f.TrimStart('#');
                var parts = f.Split(new[] { '&' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var p in parts)
                {
                    var kv = p.Split(new[] { '=' }, 2);
                    if (kv.Length == 2 && string.Equals(kv[0], key, StringComparison.OrdinalIgnoreCase))
                        return Uri.UnescapeDataString(kv[1]);
                }
            }
            catch { }
            return null;
        }

        private static string[] FetchYouTubePlaylistVideoUrls(string listId, string originalUrl)
        {
            try
            {
                var viaApi = TryFetchViaYoutubei(listId);
                if (viaApi != null && viaApi.Length > 0) return viaApi;

                var url = $"https://www.youtube.com/playlist?list={Uri.EscapeDataString(listId)}&hl=en&persist_hl=1&gl=US&persist_gl=1";
                var html = HttpGetStringWithConsent(url);

                var ids = Regex.Matches(html, "\\\"videoId\\\"\\s*:\\s*\\\"([a-zA-Z0-9_-]{11})\\\"")
                              .Cast<Match>().Select(m => m.Groups[1].Value).ToList();

                if (ids.Count == 0)
                {
                    var m2 = Regex.Matches(html, @"href=""/watch\?v=([a-zA-Z0-9_-]{11})");
                    ids.AddRange(m2.Cast<Match>().Select(m => m.Groups[1].Value));
                }

                if (ids.Count == 0)
                {
                    var m3 = Regex.Matches(html, "data-video-id=\\\"([a-zA-Z0-9_-]{11})\\\"");
                    ids.AddRange(m3.Cast<Match>().Select(m => m.Groups[1].Value));
                }

                if (ids.Count == 0)
                {
                    string watchUrl = BuildWatchUrlForList(originalUrl, listId);
                    var watchHtml = HttpGetStringWithConsent(watchUrl);
                    var m4 = Regex.Matches(watchHtml, "\\\"videoId\\\"\\s*:\\s*\\\"([a-zA-Z0-9_-]{11})\\\"");
                    ids.AddRange(m4.Cast<Match>().Select(m => m.Groups[1].Value));
                    if (ids.Count == 0)
                    {
                        var m5 = Regex.Matches(watchHtml, @"href=""/watch\?v=([a-zA-Z0-9_-]{11})");
                        ids.AddRange(m5.Cast<Match>().Select(m => m.Groups[1].Value));
                    }
                }

                if (ids.Count == 0)
                {
                    var mobile = HttpGetStringMobile($"https://m.youtube.com/playlist?list={Uri.EscapeDataString(listId)}&hl=en");
                    var m6 = Regex.Matches(mobile, @"href=""/watch\?v=([a-zA-Z0-9_-]{11})");
                    ids.AddRange(m6.Cast<Match>().Select(m => m.Groups[1].Value));
                }

                var distinct = ids.Distinct().ToArray();
                return distinct.Select(id => $"https://www.youtube.com/watch?v={id}&list={listId}").ToArray();
            }
            catch { }
            return new string[0];
        }

        private static string BuildWatchUrlForList(string originalUrl, string listId)
        {
            try
            {
                var uri = new Uri(originalUrl);
                var v = GetQueryParam(uri, "v");
                if (string.IsNullOrEmpty(v)) v = "dQw4w9WgXcQ";
                return $"https://www.youtube.com/watch?v={v}&list={listId}&hl=en&persist_hl=1&gl=US&persist_gl=1";
            }
            catch { }
            return $"https://www.youtube.com/watch?v=dQw4w9WgXcQ&list={listId}&hl=en";
        }

        private void btnStop_Click(object sender, RoutedEventArgs e)
        {
            StopPlayback();
        }

        // NEW: Save/Load playlist to file
        private void btnSavePlaylist_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new SaveFileDialog { Filter = "Vivit Playlist (*.vpl)|*.vpl|XML (*.xml)|*.xml|All Files|*.*", DefaultExt = ".vpl" };
                if (dlg.ShowDialog() != true) return;
                var list = _playlist.Select(t => new MediaTrackInfo { FilePath = t.FilePath, Url = t.Url, DisplayName = t.DisplayName, IsYouTube = t.IsYouTube }).ToList();
                var ser = new XmlSerializer(typeof(List<MediaTrackInfo>));
                using (var fs = File.Create(dlg.FileName)) ser.Serialize(fs, list);
            }
            catch { }
        }

        private void btnLoadPlaylist_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new OpenFileDialog { Filter = "Vivit Playlist (*.vpl)|*.vpl|XML (*.xml)|*.xml|All Files|*.*" };
                if (dlg.ShowDialog() != true) return;
                var ser = new XmlSerializer(typeof(List<MediaTrackInfo>));
                List<MediaTrackInfo> list = null;
                using (var fs = File.OpenRead(dlg.FileName)) list = (List<MediaTrackInfo>)ser.Deserialize(fs);
                if (list == null) return;
                _playlist.Clear();
                foreach (var it in list)
                {
                    _playlist.Add(new Track { FilePath = it.FilePath, Url = it.Url, DisplayName = it.DisplayName, IsYouTube = it.IsYouTube });
                }
                _currentIndex = _playlist.Count > 0 ? 0 : -1;
                lstPlaylist.SelectedIndex = _currentIndex;
                if (_currentIndex >= 0) LoadCurrent(false);
                PersistPlaylistToSettings();
            }
            catch { }
        }

        // NEW: Persist playlist in AppSettings
        private void PersistPlaylistToSettings()
        {
            try
            {
                var settings = AppSettings.Load();
                settings.LastMediaPlaylist = _playlist.Select(t => new MediaTrackInfo { FilePath = t.FilePath, Url = t.Url, DisplayName = t.DisplayName, IsYouTube = t.IsYouTube }).ToList();
                settings.LastMediaCurrentIndex = _currentIndex;
                settings.Save();
            }
            catch { }
        }

        private void LoadLastPlaylistFromSettings()
        {
            try
            {
                var settings = AppSettings.Load();
                _playlist.Clear();
                foreach (var t in settings.LastMediaPlaylist ?? new List<MediaTrackInfo>())
                {
                    _playlist.Add(new Track { FilePath = t.FilePath, Url = t.Url, DisplayName = t.DisplayName, IsYouTube = t.IsYouTube });
                }
                _currentIndex = (settings.LastMediaCurrentIndex >= 0 && settings.LastMediaCurrentIndex < _playlist.Count) ? settings.LastMediaCurrentIndex : (_playlist.Count > 0 ? 0 : -1);
                if (_currentIndex >= 0) lstPlaylist.SelectedIndex = _currentIndex;
            }
            catch { }
        }

        // NEW: prompt helper
        private string PromptForText(string message, string title)
        {
            try
            {
                var win = new Window { Title = title ?? "Input", SizeToContent = SizeToContent.WidthAndHeight, WindowStartupLocation = WindowStartupLocation.CenterOwner, Owner = Application.Current?.MainWindow, ResizeMode = ResizeMode.NoResize };
                var stack = new StackPanel { Margin = new Thickness(12), MinWidth = 420 };
                stack.Children.Add(new TextBlock { Text = message ?? "", Margin = new Thickness(0, 0, 0, 8) });
                var tb = new TextBox { Margin = new Thickness(0, 0, 0, 12) };
                stack.Children.Add(tb);
                var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
                var ok = new Button { Content = "OK", Width = 80, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
                var cancel = new Button { Content = "Cancel", Width = 80, IsCancel = true };
                string result = null;
                ok.Click += (_, __) => { result = tb.Text; win.DialogResult = true; };
                cancel.Click += (_, __) => { win.DialogResult = false; };
                buttons.Children.Add(ok); buttons.Children.Add(cancel);
                stack.Children.Add(buttons);
                win.Content = stack; win.ShowDialog();
                return result;
            }
            catch { return null; }
        }

        // NEW: YouTube helpers (WebView2)
        private async Task StartYouTubeAsync(Track track, bool autoPlay)
        {
            try
            {
                if (track == null || string.IsNullOrWhiteSpace(track.Url)) return;
                ShowYouTubeUi(true);
                if (ytWeb.CoreWebView2 == null)
                {
                    await ytWeb.EnsureCoreWebView2Async();
                    if (ytWeb.CoreWebView2 != null)
                    {
                        ytWeb.CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = true;
                        ytWeb.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                        ytWeb.CoreWebView2.Settings.AreHostObjectsAllowed = false;
                        ytWeb.CoreWebView2.Settings.IsStatusBarEnabled = false;
                    }
                }

                _isYouTubeActive = true;
                _isYouTubePlaying = autoPlay;
                _playRequested = autoPlay;

                string navUrl = track.Url;
                if (!navUrl.Contains("hl=")) navUrl += (navUrl.Contains("?") ? "&" : "?") + "hl=en";
                if (autoPlay && !navUrl.Contains("autoplay=")) navUrl += (navUrl.Contains("?") ? "&" : "?") + "autoplay=1";

                ytWeb.Source = new Uri(navUrl);
                txtNowPlaying.Text = $"Now Playing: {track.DisplayName}";
                txtTime.Text = "--:-- / --:--";
                SetPlayVisual(autoPlay);
            }
            catch { }
        }

        private async Task<bool> TrySendYouTubeCommandAsync(string command)
        {
            try
            {
                if (ytWeb?.CoreWebView2 == null) return false;
                string js = null;
                switch ((command ?? "").ToLowerInvariant())
                {
                    case "play": js = "(function(){var v=document.querySelector('video'); if(v){v.muted=false; v.play(); return true;} return false;})()"; break;
                    case "pause": js = "(function(){var v=document.querySelector('video'); if(v){v.pause(); return true;} return false;})()"; break;
                    case "stop": js = "(function(){var v=document.querySelector('video'); if(v){v.pause(); v.currentTime=0; return true;} return false;})()"; break;
                }
                if (js == null) return false;
                var res = await ytWeb.CoreWebView2.ExecuteScriptAsync(js);
                return true;
            }
            catch { return false; }
        }

        private void ShowYouTubeUi(bool on)
        {
            try
            {
                _isYouTubeActive = on;
                if (ytWeb != null) ytWeb.Visibility = on ? Visibility.Visible : Visibility.Collapsed;
                if (artPlaceholder != null) artPlaceholder.Visibility = on ? Visibility.Collapsed : Visibility.Visible;
                if (artText != null) artText.Text = on ? "YouTube" : "No artwork";
                posSlider.IsEnabled = !on;
                volSlider.IsEnabled = !on;
            }
            catch { }
        }

        private static string TryGetYouTubeTitle(string url)
        {
            try
            {
                var html = HttpGetStringWithConsent(url);
                if (string.IsNullOrEmpty(html)) return null;
                var m = Regex.Match(html, "<meta\\s+property=\"og:title\"\\s+content=\"([^\"]+)\"", RegexOptions.IgnoreCase);
                string title = m.Success ? m.Groups[1].Value : null;
                if (string.IsNullOrEmpty(title))
                {
                    var t = Regex.Match(html, "<title>(.*?)</title>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                    if (t.Success) title = WebUtility.HtmlDecode(t.Groups[1].Value).Replace(" - YouTube", "").Trim();
                }
                return WebUtility.HtmlDecode(title ?? string.Empty);
            }
            catch { return null; }
        }

        private static string[] TryFetchViaYoutubei(string listId)
        {
            try
            {
                // For now, return null to fall back to HTML parsing.
                return null;
            }
            catch { return null; }
        }

        private static string HttpGetStringWithConsent(string url)
        {
            try
            {
                using (var wc = new WebClient())
                {
                    wc.Encoding = Encoding.UTF8;
                    wc.Headers[HttpRequestHeader.UserAgent] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36";
                    wc.Headers[HttpRequestHeader.AcceptLanguage] = "en-US,en;q=0.9";
                    wc.Headers[HttpRequestHeader.CacheControl] = "no-cache";
                    return wc.DownloadString(url);
                }
            }
            catch { return string.Empty; }
        }

        private static string HttpGetStringMobile(string url)
        {
            try
            {
                using (var wc = new WebClient())
                {
                    wc.Encoding = Encoding.UTF8;
                    wc.Headers[HttpRequestHeader.UserAgent] = "Mozilla/5.0 (Linux; Android 10; Mobile) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Mobile Safari/537.36";
                    wc.Headers[HttpRequestHeader.AcceptLanguage] = "en-US,en;q=0.9";
                    wc.Headers[HttpRequestHeader.CacheControl] = "no-cache";
                    return wc.DownloadString(url);
                }
            }
            catch { return string.Empty; }
        }
    }
}