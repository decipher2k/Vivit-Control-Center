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

namespace Vivit_Control_Center.Views.Modules
{
    public partial class MediaPlayerModule : BaseSimpleModule
    {
        private class Track : INotifyPropertyChanged
        {
            public string FilePath { get; set; }
            public string DisplayName { get; set; }
            private TimeSpan _duration;
            public TimeSpan Duration { get => _duration; set { _duration = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Duration))); } }
            public event PropertyChangedEventHandler PropertyChanged;
        }

        private readonly ObservableCollection<Track> _playlist = new ObservableCollection<Track>();
        private int _currentIndex = -1;
        private bool _isDraggingPosition;
        private DispatcherTimer _timer;
        private bool _playRequested;

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

        private void btnPlayPause_Click(object sender, RoutedEventArgs e)
        {
            if (_currentIndex < 0)
            {
                if (_playlist.Count == 0) return;
                _currentIndex = 0;
                lstPlaylist.SelectedIndex = 0;
                LoadCurrent(autoPlay: true);
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

        private void btnStop_Click(object sender, RoutedEventArgs e) => StopPlayback();

        private void posSlider_PreviewMouseDown(object sender, MouseButtonEventArgs e) { _isDraggingPosition = true; }
        private void posSlider_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            _isDraggingPosition = false;
            try
            {
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
            if (_isDraggingPosition) return;
        }

        private void volSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
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
                if (!File.Exists(track.FilePath)) { btnNext_Click(this, null); return; }
                media.Stop();
                media.Source = new Uri(track.FilePath);
                _playRequested = autoPlay;
                SetPlayVisual(autoPlay);
                txtNowPlaying.Text = $"Now Playing: {track.DisplayName}";
                txtTime.Text = "00:00 / 00:00";
                if (autoPlay) { media.Play(); _timer.Start(); }
            }
            catch { }
        }

        private void StopPlayback()
        {
            try { media.Stop(); } catch { }
            _playRequested = false;
            SetPlayVisual(false);
            _timer.Stop();
            UpdateClockAndPosition();
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
    }
}