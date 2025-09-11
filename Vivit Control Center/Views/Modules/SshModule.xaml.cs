using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;

namespace Vivit_Control_Center.Views.Modules
{
    public partial class SshModule : BaseSimpleModule
    {
        private enum SshMode { None, SshNet, SshExe }

        // SSH.NET (dynamic)
        private object _client; // Renci.SshNet.SshClient
        private object _shell;  // Renci.SshNet.ShellStream
        private CancellationTokenSource _readCts;

        // Fallback ssh.exe
        private Process _sshProcess;
        private StreamWriter _sshInput;

        private bool _connected;
        private SshMode _mode = SshMode.None;

        public SshModule()
        {
            InitializeComponent();
            SetStatus("Bereit.");
        }

        private void SetStatus(string text)
        {
            if (txtStatus != null) txtStatus.Text = text;
        }

        private void AppendOutput(string text, Color color)
        {
            if (txtTerminal == null) return;
            var tr = new TextRange(txtTerminal.Document.ContentEnd, txtTerminal.Document.ContentEnd) { Text = text };
            tr.ApplyPropertyValue(TextElement.ForegroundProperty, new SolidColorBrush(color));
            txtTerminal.ScrollToEnd();
        }

        private async void btnConnect_Click(object sender, RoutedEventArgs e)
        {
            await ConnectAsync();
        }

        private async void txtAddress_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                await ConnectAsync();
            }
        }

        private (string user, string host, int port) ParseAddress(string addr)
        {
            string user = null; string hostPart = addr; int port = 22;
            var at = addr.IndexOf('@');
            if (at >= 0)
            {
                user = addr.Substring(0, at);
                hostPart = addr.Substring(at + 1);
            }
            var colon = hostPart.LastIndexOf(':');
            if (colon > 0 && colon < hostPart.Length - 1 && int.TryParse(hostPart.Substring(colon + 1), out var p))
            {
                port = p;
                hostPart = hostPart.Substring(0, colon);
            }
            return (user, hostPart, port);
        }

        private async Task ConnectAsync()
        {
            if (_connected)
            {
                AppendOutput("Bereits verbunden.\n", Colors.Orange);
                return;
            }

            var addr = txtAddress?.Text?.Trim();
            if (string.IsNullOrEmpty(addr))
            {
                AppendOutput("Bitte Adresse eingeben (user@host oder host[:port]).\n", Colors.Orange);
                return;
            }

            var (user, host, port) = ParseAddress(addr);
            if (string.IsNullOrWhiteSpace(host))
            {
                AppendOutput("Ungueltige Adresse.\n", Colors.Red);
                return;
            }

            try
            {
                SetStatus("Verbinden...");
                AppendOutput($"Verbinde zu {addr}...\n", Colors.Cyan);

                // Prefer SSH.NET if available
                if (TryLoadSshNet(out var sshTypes))
                {
                    // If user missing, ask for it
                    if (string.IsNullOrWhiteSpace(user))
                        user = Prompt("Benutzername:", initial: Environment.UserName, isPassword: false);
                    if (user == null) { AppendOutput("Abgebrochen.\n", Colors.Orange); SetStatus("Bereit."); return; }
                    var password = Prompt($"Passwort fuer {user}@{host}:", initial: string.Empty, isPassword: true);
                    if (password == null) { AppendOutput("Abgebrochen.\n", Colors.Orange); SetStatus("Bereit."); return; }

                    await Task.Run(() =>
                    {
                        _client = Activator.CreateInstance(sshTypes.SshClient, host, port, user, password);
                        // Ersetze diese Zeile:
                        // sshTypes.SshClient_Get_Connect(_client).Invoke(_client, null); // Connect()
                        // durch:
                        sshTypes.SshClient_Get_Connect.Invoke(_client, null); // Connect()                        
                                                                              // Ersetze diese Zeile:
                                                                              // _shell = sshTypes.SshClient_Get_CreateShellStream(_client)
                                                                              // durch:
                        _shell = sshTypes.SshClient_Get_CreateShellStream.Invoke(_client, new object[] { "xterm", 80u, 24u, 800u, 600u, 1024 });
                    });

                    _mode = SshMode.SshNet;
                    _readCts = new CancellationTokenSource();
                    _ = Task.Run(() => ReadShellLoop(sshTypes, _readCts.Token));
                }
                else
                {
                    // Fallback to ssh.exe
                    var psi = new ProcessStartInfo
                    {
                        FileName = "ssh",
                        Arguments = addr,
                        UseShellExecute = false,
                        RedirectStandardInput = true,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        CreateNoWindow = true
                    };

                    _sshProcess = new Process { StartInfo = psi, EnableRaisingEvents = true };
                    _sshProcess.OutputDataReceived += (s, e) => { if (!string.IsNullOrEmpty(e.Data)) Dispatcher.Invoke(() => AppendOutput(e.Data + "\n", Colors.White)); };
                    _sshProcess.ErrorDataReceived += (s, e) => { if (!string.IsNullOrEmpty(e.Data)) Dispatcher.Invoke(() => AppendOutput(e.Data + "\n", Colors.Red)); };
                    _sshProcess.Exited += (s, e) => Dispatcher.Invoke(() => OnDisconnected());

                    if (!_sshProcess.Start())
                    {
                        AppendOutput("Konnte ssh nicht starten.\n", Colors.Red);
                        SetStatus("Fehler");
                        return;
                    }

                    _sshProcess.BeginOutputReadLine();
                    _sshProcess.BeginErrorReadLine();
                    _sshInput = _sshProcess.StandardInput;
                    _mode = SshMode.SshExe;
                }

                _connected = true;
                btnConnect.IsEnabled = false;
                btnDisconnect.IsEnabled = true;
                txtAddress.IsEnabled = false;
                SetStatus("Verbunden");
            }
            catch (Exception ex)
            {
                AppendOutput("Verbindungsfehler: " + ex.Message + "\n", Colors.Red);
                SetStatus("Fehler");
                Cleanup();
            }
        }

        private void Cleanup()
        {
            try { _readCts?.Cancel(); } catch { }
            try { (_shell as IDisposable)?.Dispose(); } catch { }
            try { (_client as IDisposable)?.Dispose(); } catch { }
            try { _sshInput?.Dispose(); } catch { }
            try { _sshProcess?.Dispose(); } catch { }
            _readCts = null; _shell = null; _client = null; _sshInput = null; _sshProcess = null;
            _mode = SshMode.None; _connected = false;
        }

        private void ReadShellLoop((Type SshClient, Type ShellStream, MethodInfo SshClient_Get_Connect, MethodInfo SshClient_Get_CreateShellStream, MethodInfo ShellStream_Read, MethodInfo ShellStream_WriteLine) types, CancellationToken token)
        {
            try
            {
                var buffer = new byte[4096];
                while (!token.IsCancellationRequested && _shell != null)
                {
                    int read = (int)types.ShellStream_Read.Invoke(_shell, new object[] { buffer, 0, buffer.Length });
                    if (read > 0)
                    {
                        var text = Encoding.UTF8.GetString(buffer, 0, read);
                        Dispatcher.Invoke(() => AppendOutput(text, Colors.White));
                    }
                    else
                    {
                        Thread.Sleep(10);
                    }
                }
            }
            catch { }
        }

        private bool TryLoadSshNet(out (Type SshClient, Type ShellStream, MethodInfo SshClient_Get_Connect, MethodInfo SshClient_Get_CreateShellStream, MethodInfo ShellStream_Read, MethodInfo ShellStream_WriteLine) types)
        {
            types = (null, null, null, null, null, null);
            try
            {
                // Try to load by name; succeeds if Renci.SshNet.dll is next to exe or referenced
                var sshClientType = Type.GetType("Renci.SshNet.SshClient, Renci.SshNet", throwOnError: false);
                var shellStreamType = Type.GetType("Renci.SshNet.ShellStream, Renci.SshNet", throwOnError: false);
                if (sshClientType == null || shellStreamType == null) return false;

                var miConnect = sshClientType.GetMethod("Connect", BindingFlags.Public | BindingFlags.Instance);
                var miCreateShell = sshClientType.GetMethod("CreateShellStream", BindingFlags.Public | BindingFlags.Instance, null,
                    new Type[] { typeof(string), typeof(uint), typeof(uint), typeof(uint), typeof(uint), typeof(int) }, null);
                var miRead = shellStreamType.GetMethod("Read", BindingFlags.Public | BindingFlags.Instance, null,
                    new Type[] { typeof(byte[]), typeof(int), typeof(int) }, null);
                var miWriteLine = shellStreamType.GetMethod("WriteLine", BindingFlags.Public | BindingFlags.Instance, null, new[] { typeof(string) }, null);
                if (miConnect == null || miCreateShell == null || miRead == null || miWriteLine == null) return false;

                types = (sshClientType, shellStreamType, miConnect, miCreateShell, miRead, miWriteLine);
                return true;
            }
            catch { return false; }
        }

        private void btnDisconnect_Click(object sender, RoutedEventArgs e)
        {
            Disconnect();
        }

        private void Disconnect()
        {
            if (!_connected) return;
            try
            {
                if (_mode == SshMode.SshNet)
                {
                    _readCts?.Cancel();
                    try { (_shell as IDisposable)?.Dispose(); } catch { }
                    try { (_client as IDisposable)?.Dispose(); } catch { }
                }
                else if (_mode == SshMode.SshExe)
                {
                    if (_sshProcess != null && !_sshProcess.HasExited)
                    {
                        try { _sshInput?.WriteLine("exit"); _sshInput?.Flush(); } catch { }
                        if (!_sshProcess.WaitForExit(1000))
                        {
                            _sshProcess.Kill();
                        }
                    }
                }
            }
            catch { }
            finally
            {
                Cleanup();
                btnConnect.IsEnabled = true;
                btnDisconnect.IsEnabled = false;
                txtAddress.IsEnabled = true;
                SetStatus("Getrennt");
            }
        }

        private void txtInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                SendCurrentInput();
            }
        }

        private void btnSend_Click(object sender, RoutedEventArgs e)
        {
            SendCurrentInput();
        }

        private void SendCurrentInput()
        {
            if (!_connected) return;
            var text = txtInput.Text;
            txtInput.Clear();
            try
            {
                if (_mode == SshMode.SshNet)
                {
                    // Use ShellStream.WriteLine via reflection
                    var shellType = _shell?.GetType();
                    var write = shellType?.GetMethod("WriteLine", BindingFlags.Public | BindingFlags.Instance, null, new[] { typeof(string) }, null);
                    write?.Invoke(_shell, new object[] { text });
                }
                else if (_mode == SshMode.SshExe)
                {
                    _sshInput?.WriteLine(text);
                    _sshInput?.Flush();
                }
            }
            catch (Exception ex)
            {
                AppendOutput("Sende-Fehler: " + ex.Message + "\n", Colors.Red);
            }
        }

        // Simple prompt helpers (no external dependencies)
        private string Prompt(string message, string initial, bool isPassword)
        {
            var win = new Window
            {
                Title = "SSH Anmeldung",
                Width = 420,
                Height = 160,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize,
                Owner = Application.Current?.MainWindow
            };
            var grid = new Grid { Margin = new Thickness(12) };
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var lbl = new TextBlock { Text = message, Margin = new Thickness(0, 0, 0, 8) };
            Grid.SetRow(lbl, 0);
            Control input;
            if (isPassword)
            {
                var pb = new PasswordBox { Margin = new Thickness(0, 0, 0, 8) };
                pb.Password = initial ?? string.Empty;
                input = pb;
            }
            else
            {
                var tb = new TextBox { Margin = new Thickness(0, 0, 0, 8), Text = initial ?? string.Empty };
                input = tb;
            }
            Grid.SetRow(input, 1);
            var panel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var ok = new Button { Content = "OK", Width = 80, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
            var cancel = new Button { Content = "Abbrechen", Width = 80, IsCancel = true };
            ok.Click += (_, __) => { win.DialogResult = true; win.Close(); };
            cancel.Click += (_, __) => { win.DialogResult = false; win.Close(); };
            panel.Children.Add(ok); panel.Children.Add(cancel);
            Grid.SetRow(panel, 2);
            grid.Children.Add(lbl); grid.Children.Add(input); grid.Children.Add(panel);
            win.Content = grid;
            win.ShowInTaskbar = false;
            var result = win.ShowDialog();
            if (result != true) return null;
            if (isPassword) return ((PasswordBox)input).Password;
            return ((TextBox)input).Text;
        }

        protected override void OnVisualParentChanged(DependencyObject oldParent)
        {
            base.OnVisualParentChanged(oldParent);
            if (oldParent != null && VisualParent == null)
            {
                Disconnect();
            }
        }
        // Füge diese Methode in die Klasse SshModule ein (z.B. am Ende der Klasse):
        private void OnDisconnected()
        {
            Disconnect();
        }
    }
}