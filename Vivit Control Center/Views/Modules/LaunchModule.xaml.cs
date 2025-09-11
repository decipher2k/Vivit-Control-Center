using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Media.Imaging;
using Vivit_Control_Center.Settings;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Controls;

namespace Vivit_Control_Center.Views.Modules
{
    public partial class LaunchModule : BaseSimpleModule
    {
        private AppSettings _settings;
        private List<ProgramTile> _tiles = new List<ProgramTile>();
        private ItemsControl _items => this.FindName("itemsPrograms") as ItemsControl;

        public LaunchModule()
        {
            InitializeComponent();
            Loaded += (_, __) => LoadPrograms();
        }

        private void LoadPrograms()
        {
            _settings = AppSettings.Load();
            _tiles.Clear();
            var list = _settings.ExternalProgramsDetailed ?? new List<ExternalProgram>();
            foreach (var p in list)
            {
                if (string.IsNullOrWhiteSpace(p?.Path) || !File.Exists(p.Path)) continue;
                _tiles.Add(new ProgramTile
                {
                    Path = p.Path,
                    Caption = string.IsNullOrWhiteSpace(p.Caption) ? System.IO.Path.GetFileNameWithoutExtension(p.Path) : p.Caption,
                    Icon = ExtractIconImage(p.Path)
                });
            }
            var ic = _items; if (ic != null) { ic.ItemsSource = null; ic.ItemsSource = _tiles.OrderBy(t => t.Caption).ToList(); }
        }

        private BitmapImage ExtractIconImage(string file)
        {
            try
            {
                Icon ico = Icon.ExtractAssociatedIcon(file);
                if (ico == null) return null;
                using (var bmp = ico.ToBitmap())
                using (var ms = new MemoryStream())
                {
                    bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    ms.Position = 0;
                    var bi = new BitmapImage();
                    bi.BeginInit();
                    bi.CacheOption = BitmapCacheOption.OnLoad;
                    bi.StreamSource = ms;
                    bi.EndInit();
                    bi.Freeze();
                    return bi;
                }
            }
            catch { return null; }
        }

        private void Refresh_Click(object sender, RoutedEventArgs e) => LoadPrograms();

        private void ProgramTile_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            var border = sender as Border;
            if (border?.DataContext is ProgramTile tile)
            {
                try { Process.Start(tile.Path); }
                catch (Exception ex) { MessageBox.Show($"Start fehlgeschlagen: {ex.Message}"); }
            }
        }

        public override System.Threading.Tasks.Task PreloadAsync() => base.PreloadAsync();

        private class ProgramTile
        {
            public string Path { get; set; }
            public string Caption { get; set; }
            public BitmapImage Icon { get; set; }
        }
    }
}
