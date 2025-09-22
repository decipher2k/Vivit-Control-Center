using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using Vivit_Control_Center.Localization;

namespace Vivit_Control_Center.Views.Modules
{
    public partial class NotesModule : BaseSimpleModule
    {
        private string _rootDir;
        private bool _loading;
        private DispatcherTimer _autosaveTimer;
        private bool _dirty;

        private Brush DarkForeground => (Brush)Application.Current.FindResource("DarkForegroundBrush");
        private Brush DarkBackground => (Brush)Application.Current.FindResource("DarkControlBrush");

        public NotesModule()
        {
            InitializeComponent();
            Loaded += (_, __) => Init();
        }

        private void Init()
        {
            _rootDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VivitControlCenter", "notes");
            Directory.CreateDirectory(_rootDir);
            _autosaveTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2) };
            _autosaveTimer.Tick += (s, e) => { _autosaveTimer.Stop(); SaveCurrent(); };
            LoadTree();
            cmbFontSize.SelectedIndex = 1; // 12
        }

        private void LoadTree()
        {
            tvTopics.Items.Clear();
            BuildChildren(_rootDir, tvTopics.Items);
            // Ensure a default topic exists
            if (tvTopics.Items.Count == 0)
            {
                var dir = Path.Combine(_rootDir, LocalizationManager.GetString("Notes.DefaultRoot", "Allgemein"));
                Directory.CreateDirectory(dir);
                BuildChildren(_rootDir, tvTopics.Items);
            }
        }

        private void BuildChildren(string baseDir, ItemCollection items)
        {
            foreach (var dir in Directory.GetDirectories(baseDir).OrderBy(d => d, StringComparer.OrdinalIgnoreCase))
            {
                var name = Path.GetFileName(dir);
                var tvi = CreateTopicItem(name, dir);
                items.Add(tvi);
                BuildChildren(dir, tvi.Items);
            }
        }

        private TreeViewItem CreateTopicItem(string name, string fullPath)
        {
            var tvi = new TreeViewItem { Header = name, Tag = fullPath };
            tvi.ContextMenu = tvTopics.ContextMenu;
            tvi.PreviewMouseRightButtonDown += (s, e) => { tvi.IsSelected = true; e.Handled = false; };
            return tvi;
        }

        private void tvTopics_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (_dirty) SaveCurrent();
            if (tvTopics.SelectedItem is TreeViewItem tvi)
            {
                LoadTopic(tvi);
            }
        }

        private void LoadTopic(TreeViewItem tvi)
        {
            var dir = tvi.Tag as string;
            var file = Path.Combine(dir, "note.rtf");
            _loading = true;
            try
            {
                var doc = new FlowDocument
                {
                    Background = DarkBackground,
                    Foreground = DarkForeground
                };
                rtbEditor.Document = doc;
                if (File.Exists(file))
                {
                    using (var fs = File.OpenRead(file))
                    {
                        var range = new TextRange(rtbEditor.Document.ContentStart, rtbEditor.Document.ContentEnd);
                        range.Load(fs, DataFormats.Rtf);
                    }
                    // Ensure dark theme colors after loading existing content
                    ApplyDarkThemeToDocument(rtbEditor.Document);
                }
                _dirty = false;
            }
            finally
            {
                _loading = false;
            }
        }

        private void ApplyDarkThemeToDocument(FlowDocument document)
        {
            if (document == null) return;
            document.Background = DarkBackground;
            document.Foreground = DarkForeground;

            foreach (var run in EnumerateRuns(document))
            {
                var scb = run.Foreground as SolidColorBrush;
                if (run.Foreground == null || (scb != null && scb.Color == Colors.Black))
                {
                    run.Foreground = DarkForeground;
                }
            }
        }

        private static IEnumerable<Run> EnumerateRuns(FlowDocument document)
        {
            foreach (var block in document.Blocks)
            {
                foreach (var run in EnumerateRuns(block))
                    yield return run;
            }
        }

        private static IEnumerable<Run> EnumerateRuns(Block block)
        {
            if (block is Paragraph p)
            {
                foreach (Inline inline in p.Inlines)
                {
                    foreach (var run in EnumerateRuns(inline))
                        yield return run;
                }
            }
            else if (block is Section s)
            {
                foreach (var child in s.Blocks)
                {
                    foreach (var run in EnumerateRuns(child))
                        yield return run;
                }
            }
            else if (block is List list)
            {
                foreach (ListItem item in list.ListItems)
                {
                    foreach (var child in item.Blocks)
                    {
                        foreach (var run in EnumerateRuns(child))
                            yield return run;
                    }
                }
            }
            else if (block is Table table)
            {
                foreach (TableRowGroup group in table.RowGroups)
                {
                    foreach (TableRow row in group.Rows)
                    {
                        foreach (TableCell cell in row.Cells)
                        {
                            foreach (var child in cell.Blocks)
                            {
                                foreach (var run in EnumerateRuns(child))
                                    yield return run;
                            }
                        }
                    }
                }
            }
            yield break;
        }

        private static IEnumerable<Run> EnumerateRuns(Inline inline)
        {
            if (inline is Run r)
            {
                yield return r;
            }
            else if (inline is Span span)
            {
                foreach (Inline child in span.Inlines)
                {
                    foreach (var run in EnumerateRuns(child))
                        yield return run;
                }
            }
            // Ignore InlineUIContainer etc.
        }

        private void rtbEditor_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_loading) return;
            _dirty = true;
            _autosaveTimer.Stop();
            _autosaveTimer.Start();
        }

        private void rtbEditor_LostFocus(object sender, RoutedEventArgs e)
        {
            if (_dirty) SaveCurrent();
        }

        private void SaveCurrent()
        {
            if (!(tvTopics.SelectedItem is TreeViewItem tvi)) return;
            var dir = tvi.Tag as string;
            if (string.IsNullOrEmpty(dir)) return;
            Directory.CreateDirectory(dir);
            var file = Path.Combine(dir, "note.rtf");
            try
            {
                using (var fs = File.Create(file))
                {
                    var range = new TextRange(rtbEditor.Document.ContentStart, rtbEditor.Document.ContentEnd);
                    range.Save(fs, DataFormats.Rtf);
                }
                _dirty = false;
            }
            catch { }
        }

        private void AddTopic_Click(object sender, RoutedEventArgs e)
        {
            var name = Prompt(LocalizationManager.GetString("Notes.Prompt.TopicName", "Topic-Name:"));
            if (string.IsNullOrWhiteSpace(name)) return;
            var safe = MakeSafeName(name);
            string parentDir = _rootDir;
            if (tvTopics.SelectedItem is TreeViewItem sel && sel.Tag is string p) parentDir = p;
            var dir = Path.Combine(parentDir, safe);
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            // Insert into tree at correct parent
            ItemsControl targetParent = tvTopics;
            if (tvTopics.SelectedItem is TreeViewItem parentItem)
                targetParent = parentItem;

            var item = CreateTopicItem(safe, dir);
            if (targetParent is TreeViewItem pitem)
                pitem.Items.Add(item);
            else
                tvTopics.Items.Add(item);

            if (targetParent is TreeViewItem expandItem)
                expandItem.IsExpanded = true;

            item.IsSelected = true;
        }

        private void RenameTopic_Click(object sender, RoutedEventArgs e)
        {
            if (!(tvTopics.SelectedItem is TreeViewItem tvi)) return;
            var oldName = tvi.Header.ToString();
            var newName = Prompt(LocalizationManager.GetString("Notes.Prompt.NewName", "Neuer Name:"), oldName);
            if (string.IsNullOrWhiteSpace(newName) || string.Equals(newName, oldName, StringComparison.Ordinal)) return;
            var oldDir = tvi.Tag as string;
            var newDir = Path.Combine(Path.GetDirectoryName(oldDir) ?? _rootDir, MakeSafeName(newName));
            if (Directory.Exists(newDir)) { MessageBox.Show(LocalizationManager.GetString("Notes.NameExists", "Name existiert schon.")); return; }
            try
            {
                Directory.Move(oldDir, newDir);
                tvi.Header = MakeSafeName(newName);
                tvi.Tag = newDir;
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(LocalizationManager.GetString("Notes.RenameFailed", "Umbenennen fehlgeschlagen: {0}"), ex.Message));
            }
        }

        private void DeleteTopic_Click(object sender, RoutedEventArgs e)
        {
            if (!(tvTopics.SelectedItem is TreeViewItem tvi)) return;
            var dir = tvi.Tag as string;
            if (MessageBox.Show(string.Format(LocalizationManager.GetString("Notes.DeleteConfirm", "'{0}' loeschen?"), tvi.Header), LocalizationManager.GetString("Notes.DeleteTitle", "Bestaetigen"), MessageBoxButton.YesNo) != MessageBoxResult.Yes) return;
            try
            {
                Directory.Delete(dir, true);
                // Remove from parent collection
                var parent = tvi.Parent as ItemsControl;
                if (parent is TreeViewItem pti)
                    pti.Items.Remove(tvi);
                else
                    tvTopics.Items.Remove(tvi);
                rtbEditor.Document = new FlowDocument();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(LocalizationManager.GetString("Notes.DeleteFailed", "Loeschen fehlgeschlagen: {0}"), ex.Message));
            }
        }

        private void MoveUp_Click(object sender, RoutedEventArgs e)
        {
            if (!(tvTopics.SelectedItem is TreeViewItem tvi)) return;
            var parent = tvi.Parent as ItemsControl;
            var items = parent == null ? tvTopics.Items : (parent as TreeViewItem).Items;
            var idx = items.IndexOf(tvi);
            if (idx > 0)
            {
                items.RemoveAt(idx);
                items.Insert(idx - 1, tvi);
            }
        }

        private void MoveDown_Click(object sender, RoutedEventArgs e)
        {
            if (!(tvTopics.SelectedItem is TreeViewItem tvi)) return;
            var parent = tvi.Parent as ItemsControl;
            var items = parent == null ? tvTopics.Items : (parent as TreeViewItem).Items;
            var idx = items.IndexOf(tvi);
            if (idx >= 0 && idx < items.Count - 1)
            {
                items.RemoveAt(idx);
                items.Insert(idx + 1, tvi);
            }
        }

        private static void ToggleFormatting(DependencyProperty prop, object on, object off)
        {
            // no-op here; kept in other file earlier, but we keep individual handlers below
        }

        private void Bold_Click(object sender, RoutedEventArgs e) => rtbEditor.Selection.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold);
        private void Italic_Click(object sender, RoutedEventArgs e) => rtbEditor.Selection.ApplyPropertyValue(TextElement.FontStyleProperty, FontStyles.Italic);
        private void Underline_Click(object sender, RoutedEventArgs e)
        {
            var sel = rtbEditor.Selection;
            var has = sel.GetPropertyValue(Inline.TextDecorationsProperty);
            sel.ApplyPropertyValue(Inline.TextDecorationsProperty, has == DependencyProperty.UnsetValue || has == null ? TextDecorations.Underline : null);
        }
        private void Bullets_Click(object sender, RoutedEventArgs e) => EditingCommands.ToggleBullets.Execute(null, rtbEditor);
        private void Numbering_Click(object sender, RoutedEventArgs e) => EditingCommands.ToggleNumbering.Execute(null, rtbEditor);
        private void AlignLeft_Click(object sender, RoutedEventArgs e) => EditingCommands.AlignLeft.Execute(null, rtbEditor);
        private void AlignCenter_Click(object sender, RoutedEventArgs e) => EditingCommands.AlignCenter.Execute(null, rtbEditor);
        private void AlignRight_Click(object sender, RoutedEventArgs e) => EditingCommands.AlignRight.Execute(null, rtbEditor);
        private void cmbFontSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbFontSize.SelectedItem is ComboBoxItem cbi && double.TryParse(cbi.Content.ToString(), out var size))
                rtbEditor.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, size);
        }

        private static string MakeSafeName(string input)
        {
            foreach (var c in Path.GetInvalidFileNameChars()) input = input.Replace(c, '_');
            input = input.Trim();
            return string.IsNullOrEmpty(input) ? Localization.LocalizationManager.GetString("Notes.Unnamed", "Unbenannt") : input;
        }

        private string Prompt(string message, string initial = "")
        {
            var win = new Window { Title = LocalizationManager.GetString("Notes.Title", "Notiz"), Width = 360, Height = 140, WindowStartupLocation = WindowStartupLocation.CenterOwner, Owner = Application.Current.MainWindow, ResizeMode = ResizeMode.NoResize };
            var grid = new Grid { Margin = new Thickness(12) };
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var lbl = new TextBlock { Text = message, Margin = new Thickness(0, 0, 0, 8) };
            Grid.SetRow(lbl, 0);
            var tb = new TextBox { Text = initial, Margin = new Thickness(0, 0, 0, 8) };
            Grid.SetRow(tb, 1);
            var panel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var ok = new Button { Content = LocalizationManager.GetString("Dialog.OK", "OK"), Width = 80, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
            var cancel = new Button { Content = LocalizationManager.GetString("Settings.Cancel", "Abbrechen"), Width = 80, IsCancel = true };
            ok.Click += (_, __) => { win.DialogResult = true; win.Close(); };
            cancel.Click += (_, __) => { win.DialogResult = false; win.Close(); };
            panel.Children.Add(ok); panel.Children.Add(cancel);
            Grid.SetRow(panel, 2);
            grid.Children.Add(lbl); grid.Children.Add(tb); grid.Children.Add(panel);
            win.Content = grid;
            var res = win.ShowDialog();
            return res == true ? tb.Text : null;
        }

        public override void SetVisible(bool visible)
        {
            base.SetVisible(visible);
            if (visible && tvTopics.Items.Count == 0)
            {
                var dir = Path.Combine(_rootDir, LocalizationManager.GetString("Notes.DefaultRoot", "Allgemein"));
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                LoadTree();
                if (tvTopics.Items.Count > 0 && tvTopics.Items[0] is TreeViewItem t) t.IsSelected = true;
            }
        }
    }
}
