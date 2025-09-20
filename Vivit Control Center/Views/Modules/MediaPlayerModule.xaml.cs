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
                return;
            }
            if (wasCurrent)
            {
                _currentIndex = Math.Min(idx, _playlist.Count - 1);
                lstPlaylist.SelectedIndex = _currentIndex;
                LoadCurrent(autoPlay: false);
            }
        }

        private void btnClear_Click(object sender, RoutedEventArgs e)
        {
            _playlist.Clear();
            _currentIndex = -1;
            StopPlayback();
            txtNowPlaying.Text = "Now Playing: -";
            txtTime.Text = "00:00 / 00:00";
        }

        private void lstPlaylist_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (lstPlaylist.SelectedIndex >= 0)
            {
                _currentIndex = lstPlaylist.SelectedIndex;
                LoadCurrent(autoPlay: true);
            }
        }

        private void btnPrev_Click(object sender, RoutedEventArgs e)
        {
            if (_playlist.Count == 0) return;
            _currentIndex = (_currentIndex <= 0) ? _playlist.Count - 1 : _currentIndex - 1;
            lstPlaylist.SelectedIndex = _currentIndex;
            LoadCurrent(autoPlay: true);
        }

        private void btnNext_Click(object sender, RoutedEventArgs e)
        {
            if (_playlist.Count == 0) return;
            _currentIndex = (_currentIndex + 1) % _playlist.Count;
            lstPlaylist.SelectedIndex = _currentIndex;
            LoadCurrent(autoPlay: true);
        }

        private async void btnPlayPause_Click(object sender, RoutedEventArgs e)
        {
            if (_currentIndex < 0)
            {
                if (_playlist.Count == 0) return;
                _currentIndex = 0;
                lstPlaylist.SelectedIndex = 0;
                LoadCurrent(autoPlay: true);
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
                // 1) Try youtubei browse API (both variants)
                var viaApi = TryFetchViaYoutubei(listId);
                if (viaApi != null && viaApi.Length > 0) return viaApi;

                // 2) Fallback: HTML scrape (desktop)
                var url = $"https://www.youtube.com/playlist?list={Uri.EscapeDataString(listId)}&hl=en&persist_hl=1&gl=US&persist_gl=1";
                var html = HttpGetStringWithConsent(url);

                var ids = Regex.Matches(html, "\\\"videoId\\\":\\\"([a-zA-Z0-9_-]{11})\\\"")
                              .Cast<Match>().Select(m => m.Groups[1].Value).ToList();

                if (ids.Count == 0)
                {
                    var m2 = Regex.Matches(html, "href=\\\"/watch\\?v=([a-zA-Z0-9_-]{11})");
                    ids.AddRange(m2.Cast<Match>().Select(m => m.Groups[1].Value));
                }

                if (ids.Count == 0)
                {
                    var m3 = Regex.Matches(html, "data-video-id=\\\"([a-zA-Z0-9_-]{11})\\\"");
                    ids.AddRange(m3.Cast<Match>().Select(m => m.Groups[1].Value));
                }

                // 3) Fallback: try a watch page with the playlist id (covers auto mixes)
                if (ids.Count == 0)
                {
                    string watchUrl = BuildWatchUrlForList(originalUrl, listId);
                    var watchHtml = HttpGetStringWithConsent(watchUrl);
                    var m4 = Regex.Matches(watchHtml, "\\\"videoId\\\":\\\"([a-zA-Z0-9_-]{11})\\\"");
                    ids.AddRange(m4.Cast<Match>().Select(m => m.Groups[1].Value));
                    if (ids.Count == 0)
                    {
                        var m5 = Regex.Matches(watchHtml, "href=\\\"/watch\\?v=([a-zA-Z0-9_-]{11})");
                        ids.AddRange(m5.Cast<Match>().Select(m => m.Groups[1].Value));
                    }
                }

                // 4) Mobile site fallback
                if (ids.Count == 0)
                {
                    var mobile = HttpGetStringMobile($"https://m.youtube.com/playlist?list={Uri.EscapeDataString(listId)}&hl=en");
                    var m6 = Regex.Matches(mobile, "href=\\\"/watch\\?v=([a-zA-Z0-9_-]{11})");
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
                if (string.IsNullOrEmpty(v)) v = "dQw4w9WgXcQ"; // seed when missing
                return $"https://www.youtube.com/watch?v={v}&list={listId}&hl=en&persist_hl=1&gl=US&persist_gl=1";
            }
            catch { }
            return $"https://www.youtube.com/watch?v=dQw4w9WgXcQ&list={listId}&hl=en";
        }

        private static string[] TryFetchViaYoutubei(string listId)
        {
            try
            {
                var playlistUrl = $"https://www.youtube.com/playlist?list={Uri.EscapeDataString(listId)}&hl=en&persist_hl=1&gl=US&persist_gl=1";
                var pageHtml = HttpGetStringWithConsent(playlistUrl);
                if (string.IsNullOrEmpty(pageHtml)) return new string[0];

                var apiKey = Regex.Match(pageHtml, @"INNERTUBE_API_KEY""\s*:\s*""([^""]+)""").Groups[1].Value;
                var clientVersion = Regex.Match(pageHtml, @"INNERTUBE_CONTEXT_CLIENT_VERSION""\s*:\s*""([^""]+)""").Groups[1].Value;
                if (string.IsNullOrEmpty(apiKey) || string.IsNullOrEmpty(clientVersion)) return new string[0];

                var browseUrl = $"https://www.youtube.com/youtubei/v1/browse?key={apiKey}";
                var contextJson = "\"context\":{\"client\":{\"clientName\":\"WEB\",\"clientVersion\":\"" + clientVersion + "\"}}";

                var ids = new System.Collections.Generic.List<string>();

                // Variant A: browseId = listId
                var payloadA = $"{{{contextJson},\"browseId\":\"{listId}\"}}";
                var resultA = HttpPostJsonWithConsent(browseUrl, payloadA, playlistUrl, clientVersion);
                ids.AddRange(Regex.Matches(resultA ?? string.Empty, "\\\"videoId\\\":\\\"([a-zA-Z0-9_-]{11})\\\"").Cast<Match>().Select(m => m.Groups[1].Value));

                // Variant B: browseId = VL+listId
                if (ids.Count == 0)
                {
                    var payloadB = $"{{{contextJson},\"browseId\":\"VL{listId}\"}}";
                    var resultB = HttpPostJsonWithConsent(browseUrl, payloadB, playlistUrl, clientVersion);
                    ids.AddRange(Regex.Matches(resultB ?? string.Empty, "\\\"videoId\\\":\\\"([a-zA-Z0-9_-]{11})\\\"").Cast<Match>().Select(m => m.Groups[1].Value));
                }

                return ids.Distinct().Select(id => $"https://www.youtube.com/watch?v={id}&list={listId}").ToArray();
            }
            catch { }
            return new string[0];
        }

        private static string ExtractContinuationToken(string json)
        {
            try
            {
                if (string.IsNullOrEmpty(json)) return null;
                var m = Regex.Match(json, "(\\\"continuationCommand\\\"\\s*:\\s*\\{\\s*\\\"token\\\"\\s*:\\s*\\\"([^\\\"]+)\\\")|(\\\"nextContinuationData\\\"\\s*:\\s*\\{[^}]*?\\\"continuation\\\"\\s*:\\s*\\\"([^\\\"]+)\\\")");
                if (m.Success)
                {
                    var g2 = m.Groups[2].Success ? m.Groups[2].Value : null;
                    var g4 = m.Groups[4].Success ? m.Groups[4].Value : null;
                    return !string.IsNullOrEmpty(g2) ? g2 : g4;
                }
            }
            catch { }
            return null;
        }

        private static string HttpGetStringWithConsent(string url)
        {
            try
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
                var req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = "GET";
                req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0 Safari/537.36";
                req.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
                req.Headers[HttpRequestHeader.AcceptLanguage] = "en-US,en;q=0.9";
                req.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                req.CookieContainer = new CookieContainer();
                req.CookieContainer.Add(new Cookie("CONSENT", "YES+cb", "/", ".youtube.com"));
                req.CookieContainer.Add(new Cookie("PREF", "hl=en&gl=US", "/", ".youtube.com"));
                req.CookieContainer.Add(new Cookie("SOCS", "CAISAiAB", "/", ".google.com"));
                req.AllowAutoRedirect = true;
                using (var resp = (HttpWebResponse)req.GetResponse())
                using (var stream = resp.GetResponseStream())
                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("YouTube GET error: " + ex.Message);
            }
            return string.Empty;
        }

        private static string HttpGetStringMobile(string url)
        {
            try
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
                var req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = "GET";
                req.UserAgent = "Mozilla/5.0 (Linux; Android 10; Pixel 3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0 Mobile Safari/537.36";
                req.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
                req.Headers[HttpRequestHeader.AcceptLanguage] = "en-US,en;q=0.9";
                req.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                req.CookieContainer = new CookieContainer();
                req.CookieContainer.Add(new Cookie("CONSENT", "YES+cb", "/", ".youtube.com"));
                req.CookieContainer.Add(new Cookie("PREF", "hl=en&gl=US", "/", ".youtube.com"));
                req.AllowAutoRedirect = true;
                using (var resp = (HttpWebResponse)req.GetResponse())
                using (var stream = resp.GetResponseStream())
                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("YouTube Mobile GET error: " + ex.Message);
            }
            return string.Empty;
        }

        private static string HttpPostJsonWithConsent(string url, string jsonPayload, string referer, string clientVersion)
        {
            try
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
                var req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = "POST";
                req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0 Safari/537.36";
                req.Accept = "application/json";
                req.Headers[HttpRequestHeader.AcceptLanguage] = "en-US,en;q=0.9";
                req.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                req.CookieContainer = new CookieContainer();
                req.CookieContainer.Add(new Cookie("CONSENT", "YES+cb", "/", ".youtube.com"));
                req.CookieContainer.Add(new Cookie("PREF", "hl=en&gl=US", "/", ".youtube.com"));
                req.CookieContainer.Add(new Cookie("SOCS", "CAISAiAB", "/", ".google.com"));
                req.Referer = referer;
                req.Headers["Origin"] = "https://www.youtube.com";
                req.ContentType = "application/json";
                req.Headers["X-YouTube-Client-Name"] = "1";
                req.Headers["X-YouTube-Client-Version"] = clientVersion;

                var bytes = Encoding.UTF8.GetBytes(jsonPayload ?? "{}");
                using (var stream = req.GetRequestStream())
                {
                    stream.Write(bytes, 0, bytes.Length);
                }
                using (var resp = (HttpWebResponse)req.GetResponse())
                using (var s = resp.GetResponseStream())
                using (var reader = new StreamReader(s))
                {
                    return reader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("YouTube POST error: " + ex.Message);
            }
            return string.Empty;
        }

        private static string TryGetYouTubeTitle(string videoUrl)
        {
            try
            {
                var oembed = $"https://www.youtube.com/oembed?url={Uri.EscapeDataString(videoUrl)}&format=json";
                using (var wc = new WebClient())
                {
                    wc.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)");
                    var json = wc.DownloadString(oembed);
                    var m = Regex.Match(json, "\\\"title\\\"\\s*:\\s*\\\"(.*?)\\\"");
                    if (m.Success)
                    {
                        return WebUtility.HtmlDecode(m.Groups[1].Value);
                    }
                }
            }
            catch { }
            return null;
        }

        private static string PromptForText(string message, string caption)
        {
            var win = new Window
            {
                Title = caption,
                Width = 480,
                Height = 160,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize,
                WindowStyle = WindowStyle.ToolWindow,
                Owner = Application.Current?.MainWindow
            };
            var grid = new Grid { Margin = new Thickness(12) };
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var txt = new TextBox { Margin = new Thickness(0, 8, 0, 12) };
            txt.MinWidth = 420;
            var lbl = new TextBlock { Text = message };
            var panel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var ok = new Button { Content = "OK", Width = 80, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
            var cancel = new Button { Content = "Cancel", Width = 80, IsCancel = true };
            ok.Click += (s, e) => { win.DialogResult = true; win.Close(); };
            cancel.Click += (s, e) => { win.DialogResult = false; win.Close(); };
            panel.Children.Add(ok);
            panel.Children.Add(cancel);
            Grid.SetRow(lbl, 0);
            Grid.SetRow(txt, 1);
            Grid.SetRow(panel, 2);
            grid.Children.Add(lbl);
            grid.Children.Add(txt);
            grid.Children.Add(panel);
            win.Content = grid;
            var res = win.ShowDialog();
            return res == true ? txt.Text : null;
        }

        // ===== YouTube WebView2 integration (watch page) =====
        private async System.Threading.Tasks.Task StartYouTubeAsync(Track track, bool autoPlay)
        {
            try
            {
                string videoId = ExtractVideoIdFromUrl(track.Url);
                if (string.IsNullOrEmpty(videoId)) return;

                await ytWeb.EnsureCoreWebView2Async(null);

                var listId = ExtractPlaylistId(track.Url);
                var watchUrl = BuildWatchUrlForList(track.Url, listId);
                if (!watchUrl.Contains("autoplay=1"))
                    watchUrl += (watchUrl.Contains("?") ? "&" : "?") + "autoplay=1";

                // After navigation complete, try to start playback or toggle state by clicking the play button
                ytWeb.CoreWebView2.NavigationCompleted += async (s, e) =>
                {
                    try
                    {
                        // Give the page a moment to initialize the player UI
                        var t = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(600) };
                        t.Tick += async (s2, e2) =>
                        {
                            t.Stop();
                            if (autoPlay)
                            {
                                await TrySendYouTubeCommandAsync("play");
                            }
                        };
                        t.Start();
                    }
                    catch { }
                };

                ytWeb.CoreWebView2.Navigate(watchUrl);

                txtNowPlaying.Text = $"Now Playing: {track.DisplayName}";
                txtTime.Text = "--:-- / --:--";
                ShowYouTubeUi(true);
                _isYouTubePlaying = autoPlay;
                SetPlayVisual(autoPlay);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("YouTube start failed: " + ex.Message);
            }
        }

        private void ShowYouTubeUi(bool show)
        {
            _isYouTubeActive = show;
            if (ytWeb == null || artPlaceholder == null || artText == null) return;
            ytWeb.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
            artPlaceholder.Visibility = show ? Visibility.Collapsed : Visibility.Visible;
            artText.Visibility = show ? Visibility.Collapsed : Visibility.Visible;
            posSlider.IsEnabled = !show;
            volSlider.IsEnabled = !show;
        }

        private static string ExtractVideoIdFromUrl(string url)
        {
            try
            {
                var m = Regex.Match(url ?? string.Empty, @"[?&]v=([a-zA-Z0-9_-]{11})");
                if (m.Success) return m.Groups[1].Value;
                var m2 = Regex.Match(url ?? string.Empty, @"youtu\.be/([a-zA-Z0-9_-]{11})");
                if (m2.Success) return m2.Groups[1].Value;
            }
            catch { }
            return null;
        }

        private async System.Threading.Tasks.Task TrySendYouTubeCommandAsync(string cmd)
        {
            try
            {
                if (ytWeb?.CoreWebView2 == null) return;
                if (cmd == "play")
                {
                    await ytWeb.CoreWebView2.ExecuteScriptAsync(
                        @"(function(){var b=document.querySelector('.ytp-play-button[title*=""Play"" i], .ytp-play-button[aria-label*=""Play"" i]'); if(b){b.click(); return 'clicked';} document.dispatchEvent(new KeyboardEvent('keydown',{key:'k'})); return 'key';})();");
                }
                else if (cmd == "pause")
                {
                    await ytWeb.CoreWebView2.ExecuteScriptAsync(
                        @"(function(){var b=document.querySelector('.ytp-play-button[title*=""Pause"" i], .ytp-play-button[aria-label*=""Pause"" i]'); if(b){b.click(); return 'clicked';} document.dispatchEvent(new KeyboardEvent('keydown',{key:'k'})); return 'key';})();");
                }
                else if (cmd == "stop")
                {
                    await ytWeb.CoreWebView2.ExecuteScriptAsync(
                        @"(function(){var b=document.querySelector('.ytp-play-button[aria-label*=""Pause"" i], .ytp-play-button[title*=""Pause"" i]'); if(b){b.click();} })();");
                }
            }
            catch { }
        }

        private void btnStop_Click(object sender, RoutedEventArgs e)
        {
            StopPlayback();
        }
    }
}