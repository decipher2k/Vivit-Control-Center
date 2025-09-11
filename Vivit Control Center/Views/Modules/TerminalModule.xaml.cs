using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Text.RegularExpressions;

namespace Vivit_Control_Center.Views.Modules
{
    public partial class TerminalModule : UserControl, IModule
    {
        private Process powershellProcess;
        private CancellationTokenSource cancellationTokenSource;
        private List<string> commandHistory = new List<string>();
        private bool isRunning = false;
        private string currentDir = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        
        private bool _signaled;
        private readonly TaskCompletionSource<bool> _tcs =
            new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

        public event EventHandler LoadCompleted;
        public Task LoadCompletedTask => _tcs.Task;
        public FrameworkElement View => this;

        public TerminalModule()
        {
            InitializeComponent();
            Loaded += OnModuleLoaded;
        }

        private void OnModuleLoaded(object sender, RoutedEventArgs e)
        {
            SignalLoadedOnce();
            InitializeTerminal();
        }

        public virtual Task PreloadAsync()
        {
            SignalLoadedOnce();
            return _tcs.Task;
        }

        public virtual void SetVisible(bool visible)
        {
            Visibility = visible ? Visibility.Visible : Visibility.Hidden;
            IsHitTestVisible = visible;
        }

        protected void SignalLoadedOnce()
        {
            if (_signaled) return;
            _signaled = true;
            _tcs.TrySetResult(true);
            try { LoadCompleted?.Invoke(this, EventArgs.Empty); } catch { }
        }

        private void InitializeTerminal()
        {
            try
            {
                AppendText("PowerShell Terminal wird initialisiert...\n", Colors.Green);
                AppendText("Bereit für Eingaben.\n", Colors.Green);
                
                // Aktuelles Verzeichnis anzeigen
                AppendText($"PS {currentDir}> ", Colors.Cyan);
            }
            catch (Exception ex)
            {
                AppendText($"Fehler beim Initialisieren des Terminals: {ex.Message}\n", Colors.Red);
            }
        }

        private async void ExecuteCommand(string command)
        {
            if (string.IsNullOrWhiteSpace(command))
                return;

            // Command zu History hinzufügen
            if (!commandHistory.Contains(command))
            {
                commandHistory.Add(command);
                cmbHistory.Items.Add(command);
            }
            
            txtCommand.Clear();
            AppendText(command + "\n", Colors.White);
            
            isRunning = true;
            btnStop.IsEnabled = true;
            btnExecute.IsEnabled = false;

            try
            {
                cancellationTokenSource = new CancellationTokenSource();
                
                await Task.Run(() =>
                {
                    try
                    {
                        RunPowerShell(command);
                    }
                    catch (Exception ex)
                    {
                        Application.Current.Dispatcher.Invoke(() => 
                            AppendText($"Ausführungsfehler: {ex.Message}\n", Colors.Red));
                    }
                    
                    Application.Current.Dispatcher.Invoke(() => 
                        AppendText($"PS {currentDir}> ", Colors.Cyan));
                    
                }, cancellationTokenSource.Token);
            }
            catch (OperationCanceledException)
            {
                AppendText("Befehl wurde abgebrochen.\n", Colors.Orange);
                AppendText($"PS {currentDir}> ", Colors.Cyan);
            }
            catch (Exception ex)
            {
                AppendText($"Fehler: {ex.Message}\n", Colors.Red);
                AppendText($"PS {currentDir}> ", Colors.Cyan);
            }
            finally
            {
                isRunning = false;
                btnStop.IsEnabled = false;
                btnExecute.IsEnabled = true;
                cancellationTokenSource = null;
            }
        }

        private string NormalizeCommand(string command)
        {
            if (string.IsNullOrWhiteSpace(command)) return command;

            // Reine Einzel- oder Doppel-Quote um den gesamten Befehl: &-Aufruf erzwingen
            var mSingle = Regex.Match(command, @"^\s*'([^']+)'\s*(.*)$");
            if (mSingle.Success)
            {
                var path = mSingle.Groups[1].Value;
                var rest = mSingle.Groups[2].Value;
                return ($"& '{path}' {rest}").TrimEnd();
            }

            var mDouble = Regex.Match(command, @"^\s*""([^""]+)""\s*(.*)$");
            if (mDouble.Success)
            {
                var path = mDouble.Groups[1].Value;
                var rest = mDouble.Groups[2].Value;
                return ($"& \"{path}\" {rest}").TrimEnd();
            }

            return command;
        }

        private void RunPowerShell(string command)
        {
            string tempDirectoryFile = Path.Combine(Path.GetTempPath(), $"ps_dir_{Guid.NewGuid()}.txt");

            // Befehl normalisieren: falls nur ein gequoteter Pfad, dann über & ausführen
            var normalized = NormalizeCommand(command);
            // Doppelte Anführungszeichen für -Command Kontext escapen
            var normalizedEscaped = normalized.Replace("\"", "`\"");

            string commandWithPath = $"cd \"{currentDir}\"; {normalizedEscaped}; [System.IO.File]::WriteAllText('{tempDirectoryFile.Replace("\\", "\\\\")}', (Get-Location).Path)";

            using (powershellProcess = new Process())
            {
                powershellProcess.StartInfo.FileName = "powershell.exe";
                powershellProcess.StartInfo.Arguments = $"-NoLogo -ExecutionPolicy Bypass -Command \"{commandWithPath}\"";
                powershellProcess.StartInfo.UseShellExecute = false;
                powershellProcess.StartInfo.RedirectStandardOutput = true;
                powershellProcess.StartInfo.RedirectStandardError = true;
                powershellProcess.StartInfo.CreateNoWindow = true;
                powershellProcess.StartInfo.WorkingDirectory = currentDir;

                powershellProcess.OutputDataReceived += (sender, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                    {
                        Application.Current.Dispatcher.Invoke(() => AppendText(e.Data + "\n", Colors.White));
                    }
                };

                powershellProcess.ErrorDataReceived += (sender, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                    {
                        Application.Current.Dispatcher.Invoke(() => AppendText(e.Data + "\n", Colors.Red));
                    }
                };

                powershellProcess.Start();
                powershellProcess.BeginOutputReadLine();
                powershellProcess.BeginErrorReadLine();
                powershellProcess.WaitForExit();

                if (File.Exists(tempDirectoryFile))
                {
                    try
                    {
                        string newDir = File.ReadAllText(tempDirectoryFile).Trim();
                        if (!string.IsNullOrEmpty(newDir) && Directory.Exists(newDir))
                        {
                            currentDir = newDir;
                        }
                        File.Delete(tempDirectoryFile);
                    }
                    catch { }
                }
            }
        }

        private void AppendText(string text, Color color)
        {
            TextRange tr = new TextRange(txtOutput.Document.ContentEnd, txtOutput.Document.ContentEnd);
            tr.Text = text;
            tr.ApplyPropertyValue(TextElement.ForegroundProperty, new SolidColorBrush(color));
            txtOutput.ScrollToEnd();
        }

        private void btnExecute_Click(object sender, RoutedEventArgs e)
        {
            ExecuteCommand(txtCommand.Text);
        }

        private void btnClear_Click(object sender, RoutedEventArgs e)
        {
            txtOutput.Document.Blocks.Clear();
            AppendText("Terminal-Ausgabe geloescht.\n", Colors.Green);
            AppendText($"PS {currentDir}> ", Colors.Cyan);
        }

        private void btnStop_Click(object sender, RoutedEventArgs e)
        {
            cancellationTokenSource?.Cancel();
            if (powershellProcess != null && !powershellProcess.HasExited)
            {
                try
                {
                    powershellProcess.Kill();
                    AppendText("Befehl abgebrochen.\n", Colors.Orange);
                    AppendText($"PS {currentDir}> ", Colors.Cyan);
                }
                catch (Exception ex)
                {
                    AppendText($"Fehler beim Abbrechen: {ex.Message}\n", Colors.Red);
                }
            }
        }

        private void txtCommand_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                ExecuteCommand(txtCommand.Text);
                e.Handled = true;
            }
        }

        private void cmbHistory_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbHistory.SelectedItem != null)
            {
                txtCommand.Text = cmbHistory.SelectedItem.ToString();
                cmbHistory.SelectedIndex = -1; // Auswahl zurücksetzen
                txtCommand.Focus();
                txtCommand.SelectionStart = txtCommand.Text.Length;
            }
        }
    }
}