using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Vivit_Control_Center.Settings;

namespace Vivit_Control_Center.Views.Modules
{
    public partial class ProgramsDialog : Window
    {
        private readonly AppSettings _settings;
        private List<ExternalProgram> _working;
        // ensure field for code-behind access (named element normally auto-generated; fallback assign after InitializeComponent)
        private DataGrid _dg => this.FindName("dgPrograms") as DataGrid;

        public ProgramsDialog(AppSettings settings)
        {
            InitializeComponent();
            _settings = settings;
            _working = settings.ExternalProgramsDetailed?.Select(p => new ExternalProgram { Path = p.Path, Caption = p.Caption }).ToList() ?? new List<ExternalProgram>();
            if (_working.Count == 0 && (settings.ExternalPrograms?.Count > 0))
            {
                foreach (var p in settings.ExternalPrograms)
                {
                    if (string.IsNullOrWhiteSpace(p)) continue;
                    _working.Add(new ExternalProgram { Path = p, Caption = System.IO.Path.GetFileNameWithoutExtension(p) });
                }
            }
            RefreshGrid();
        }

        private void RefreshGrid()
        {
            var dg = _dg; if (dg == null) return; dg.ItemsSource = null; dg.ItemsSource = _working.OrderBy(x => System.IO.Path.GetFileName(x.Path)).ToList();
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "Programme (*.exe)|*.exe|Alle Dateien (*.*)|*.*" };
            if (dlg.ShowDialog() == true)
            {
                if (!_working.Any(p => string.Equals(p.Path, dlg.FileName, StringComparison.OrdinalIgnoreCase)))
                {
                    _working.Add(new ExternalProgram { Path = dlg.FileName, Caption = System.IO.Path.GetFileNameWithoutExtension(dlg.FileName) });
                    RefreshGrid();
                }
            }
        }

        private void Remove_Click(object sender, RoutedEventArgs e)
        {
            var dg = _dg; if (dg == null) return; var selected = dg.SelectedItems.Cast<ExternalProgram>().ToList(); if (selected.Count == 0) return; _working.RemoveAll(p => selected.Any(s => string.Equals(s.Path, p.Path, StringComparison.OrdinalIgnoreCase))); RefreshGrid();
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            _settings.ExternalProgramsDetailed = _working.Where(p => !string.IsNullOrWhiteSpace(p.Path)).Select(p => new ExternalProgram { Path = p.Path, Caption = p.Caption?.Trim() }).ToList(); _settings.Save(); MessageBox.Show("Gespeichert.");
        }

        private void Close_Click(object sender, RoutedEventArgs e) => Close();
    }
}
