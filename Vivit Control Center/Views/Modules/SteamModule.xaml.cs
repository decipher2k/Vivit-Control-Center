using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;

namespace Vivit_Control_Center.Views.Modules
{
    public partial class SteamModule : BaseSimpleModule
    {
        private readonly ObservableCollection<SteamGame> _games = new ObservableCollection<SteamGame>();
        private string _steamPath;
        private readonly HttpClient _httpClient = new HttpClient();

        public ObservableCollection<SteamGame> Games => _games;

        public SteamModule()
        {
            InitializeComponent();
            DataContext = this;
            
            Task.Run(() => LoadGamesAsync())
                .ContinueWith(_ => SignalLoadedOnce(), TaskScheduler.FromCurrentSynchronizationContext());
        }

        private async Task LoadGamesAsync()
        {
            try
            {
                // Steam-Installationspfad aus der Registry ermitteln
                _steamPath = GetSteamInstallPath();
                if (string.IsNullOrEmpty(_steamPath) || !Directory.Exists(_steamPath))
                {
                    MessageBox.Show("Steam-Installation konnte nicht gefunden werden.", "Fehler", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                Debug.WriteLine($"Steam-Pfad gefunden: {_steamPath}");

                // Bibliotheksordner aus libraryfolders.vdf auslesen
                var libraryFolders = GetSteamLibraryFolders();
                if (libraryFolders.Count == 0)
                    libraryFolders.Add(_steamPath); // Fallback auf Haupt-Steam-Verzeichnis

                Debug.WriteLine($"Gefundene Bibliotheksordner: {string.Join(", ", libraryFolders)}");

                // Alle installierten Spiele aus den Bibliotheksordnern auslesen
                foreach (var libraryFolder in libraryFolders)
                {
                    var appsPath = Path.Combine(libraryFolder, "steamapps");
                    if (!Directory.Exists(appsPath))
                    {
                        Debug.WriteLine($"Verzeichnis nicht gefunden: {appsPath}");
                        continue;
                    }

                    // Alle appmanifest_*.acf Dateien durchgehen
                    foreach (var manifestFile in Directory.GetFiles(appsPath, "appmanifest_*.acf"))
                    {
                        var game = ParseAppManifest(manifestFile);
                        if (game != null)
                        {
                            Debug.WriteLine($"Spiel gefunden: {game.Name} (AppID: {game.AppId})");
                            
                            // Thumbnail von Steam CDN laden
                            await LoadGameThumbnailFromSteamCDNAsync(game);
                            
                            // Spiel zur Sammlung hinzufügen
                            await Application.Current.Dispatcher.InvokeAsync(() => _games.Add(game));
                        }
                    }
                }

                // Nach Namen sortieren
                var sorted = _games.OrderBy(g => g.Name).ToList();
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    _games.Clear();
                    foreach (var game in sorted)
                        _games.Add(game);
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Fehler beim Laden der Steam-Spiele: {ex}");
                MessageBox.Show($"Fehler beim Laden der Steam-Spiele: {ex.Message}", "Fehler", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private string GetSteamInstallPath()
        {
            try
            {
                // 64-bit Registry-Pfad für 32-bit Steam
                using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\WOW6432Node\Valve\Steam"))
                {
                    if (key != null)
                    {
                        var path = key.GetValue("InstallPath") as string;
                        if (!string.IsNullOrEmpty(path))
                            return path;
                    }
                }

                // Direkter Registry-Pfad
                using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Valve\Steam"))
                {
                    if (key != null)
                    {
                        var path = key.GetValue("InstallPath") as string;
                        if (!string.IsNullOrEmpty(path))
                            return path;
                    }
                }

                // Standard-Installationspfade
                string[] defaultPaths = {
                    @"C:\Program Files (x86)\Steam",
                    @"C:\Program Files\Steam",
                    @"D:\Steam",
                    @"E:\Steam"
                };

                return defaultPaths.FirstOrDefault(Directory.Exists);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Fehler beim Ermitteln des Steam-Pfads: {ex.Message}");
                // Fallback auf Standard-Pfad
                return @"C:\Program Files (x86)\Steam";
            }
        }

        private System.Collections.Generic.List<string> GetSteamLibraryFolders()
        {
            var libraryFolders = new System.Collections.Generic.List<string>();
            
            try
            {
                // libraryfolders.vdf Pfad
                var vdfPath = Path.Combine(_steamPath, "steamapps", "libraryfolders.vdf");
                if (File.Exists(vdfPath))
                {
                    // Einfaches Parsing der VDF-Datei, um Bibliothekspfade zu extrahieren
                    var content = File.ReadAllText(vdfPath);
                    var pathRegex = new Regex(@"""path""[\s]*""([^""]+)""");
                    var matches = pathRegex.Matches(content);

                    foreach (Match match in matches)
                    {
                        if (match.Groups.Count > 1)
                        {
                            var path = match.Groups[1].Value.Replace("\\\\", "\\");
                            if (Directory.Exists(path))
                                libraryFolders.Add(path);
                        }
                    }
                }
            }
            catch (Exception ex) 
            { 
                Debug.WriteLine($"Fehler beim Lesen der Steam-Bibliotheksordner: {ex.Message}");
            }

            return libraryFolders;
        }

        private SteamGame ParseAppManifest(string manifestFile)
        {
            try
            {
                var content = File.ReadAllText(manifestFile);
                
                // AppID extrahieren
                var appIdMatch = Regex.Match(content, @"""appid""[\s]*""(\d+)""");
                if (!appIdMatch.Success) return null;
                var appId = appIdMatch.Groups[1].Value;
                
                // Name extrahieren
                var nameMatch = Regex.Match(content, @"""name""[\s]*""([^""]+)""");
                if (!nameMatch.Success) return null;
                var name = nameMatch.Groups[1].Value;
                
                return new SteamGame
                {
                    AppId = appId,
                    Name = name
                };
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Fehler beim Parsen der Manifest-Datei {manifestFile}: {ex.Message}");
                return null;
            }
        }

        private async Task LoadGameThumbnailFromSteamCDNAsync(SteamGame game)
        {
            try
            {
                // Steam CDN URLs für Spielbilder
                string[] imageUrls = {
                    $"https://cdn.cloudflare.steamstatic.com/steam/apps/{game.AppId}/header.jpg",
                    $"https://cdn.cloudflare.steamstatic.com/steam/apps/{game.AppId}/capsule_616x353.jpg",
                    $"https://cdn.cloudflare.steamstatic.com/steam/apps/{game.AppId}/library_600x900.jpg",
                    $"https://cdn.cloudflare.steamstatic.com/steam/apps/{game.AppId}/portrait.png"
                };

                foreach (var url in imageUrls)
                {
                    try
                    {
                        Debug.WriteLine($"Versuche, Bild zu laden von: {url}");
                        
                        // Bild vom CDN laden
                        var imageBytes = await _httpClient.GetByteArrayAsync(url);
                        
                        // Wenn wir hier sind, war der Download erfolgreich
                        await Application.Current.Dispatcher.InvokeAsync(() =>
                        {
                            try
                            {
                                var bitmap = new BitmapImage();
                                bitmap.BeginInit();
                                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                                bitmap.StreamSource = new MemoryStream(imageBytes);
                                bitmap.EndInit();
                                bitmap.Freeze(); // Für Thread-Sicherheit
                                game.ThumbnailImage = bitmap;
                                Debug.WriteLine($"Thumbnail erfolgreich geladen von: {url}");
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"Fehler beim Konvertieren des Bildes: {ex.Message}");
                            }
                        });
                        
                        // Wenn wir erfolgreich ein Bild geladen haben, brechen wir die Schleife ab
                        return;
                    }
                    catch (HttpRequestException ex)
                    {
                        Debug.WriteLine($"HTTP-Fehler beim Laden des Bildes von {url}: {ex.Message}");
                        // Wir versuchen die nächste URL
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Allgemeiner Fehler beim Laden des Bildes von {url}: {ex.Message}");
                        // Wir versuchen die nächste URL
                    }
                }
                
                Debug.WriteLine($"Konnte kein Bild für AppID {game.AppId} vom Steam CDN laden");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Fehler beim Laden des Thumbnails vom CDN: {ex.Message}");
            }
        }

        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2 && sender is FrameworkElement element && element.DataContext is SteamGame game)
            {
                try
                {
                    // Spiel über Steam-Protokoll starten
                    Process.Start($"steam://run/{game.AppId}");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Fehler beim Starten des Spiels: {ex.Message}", "Fehler", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }

    public class SteamGame : INotifyPropertyChanged
    {
        private string _appId;
        private string _name;
        private BitmapSource _thumbnailImage;

        public string AppId
        {
            get => _appId;
            set
            {
                if (_appId != value)
                {
                    _appId = value;
                    OnPropertyChanged();
                }
            }
        }

        public string Name
        {
            get => _name;
            set
            {
                if (_name != value)
                {
                    _name = value;
                    OnPropertyChanged();
                }
            }
        }

        public BitmapSource ThumbnailImage
        {
            get => _thumbnailImage;
            set
            {
                if (_thumbnailImage != value)
                {
                    _thumbnailImage = value;
                    OnPropertyChanged();
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}