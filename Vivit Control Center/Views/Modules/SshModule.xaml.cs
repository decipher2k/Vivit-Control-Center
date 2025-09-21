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
using System.Text.RegularExpressions;
using Vivit_Control_Center.Localization;
using Vivit_Control_Center.Settings;

namespace Vivit_Control_Center.Views.Modules
{
    public partial class SshModule : BaseSimpleModule
    {
        private enum SshMode { None, SshNet, SshExe }
        private object _client;
        private object _shell;
        private CancellationTokenSource _readCts;
        private Process _sshProcess;
        private StreamWriter _sshInput;
        private bool _connected;
        private SshMode _mode = SshMode.None;

        // Capture mode for one-shot log load
        private volatile bool _captureLog;
        private readonly StringBuilder _logBuffer = new StringBuilder();
        private string _captureEndMarker;

        // Follow mode state
        private volatile bool _isFollowing;
        private CancellationTokenSource _followCts;
        private object _logShell;            // for SSH.NET
        private Process _logSshProcess;      // for ssh.exe

        // ANSI stripping regex: CSI, OSC, ESC with single final, and ESC with one intermediate + final (covers ESC(B, ESC)0, etc.)
        private static readonly Regex AnsiRegex = new Regex(
            @"\x1B\[[0-?]*[ -/]*[@-~]|\x1B\][^\x1B\x07]*(\x07|\x1B\\)|\x1B[ -/][@-~]|\x1B[@-Z\\-_]",
            RegexOptions.Compiled);

        private static string StripAnsi(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            try { return AnsiRegex.Replace(s, string.Empty); } catch { return s; }
        }

        // Helpers to find sidebar controls without generated fields
        private ComboBox GetLogCombo() => FindName("cboLogFiles") as ComboBox;
        private Panel GetMacrosPanel() => FindName("panelMacros") as Panel;
        private TextBox GetLogViewer() => FindName("txtLogViewer") as TextBox;
        private CheckBox GetFollowCheckbox() => FindName("chkFollow") as CheckBox;
        private Button GetLoadButton() => FindName("btnLoadLog") as Button;

        public SshModule()
        {
            InitializeComponent();
            SetStatus(LocalizationManager.GetString("SSH.Status.Ready", "Ready."));
            Loaded += (_, __) => InitSidebarFromSettings();

            var btnClear = FindName("btnClearLog") as Button; if (btnClear != null) btnClear.Click += btnClearLog_Click;
            var chk = GetFollowCheckbox(); if (chk != null) { chk.Checked += chkFollow_Checked; chk.Unchecked += chkFollow_Unchecked; }
        }

        private void InitSidebarFromSettings()
        {
            try
            {
                var s = AppSettings.Load();
                var combo = GetLogCombo();
                if (combo != null)
                {
                    combo.ItemsSource = (s.SshLogFiles != null && s.SshLogFiles.Count > 0)
                        ? s.SshLogFiles
                        : new System.Collections.Generic.List<string> { "/var/log/auth.log", "/var/log/syslog" };
                    if (combo.Items.Count > 0) combo.SelectedIndex = 0;
                }

                var panel = GetMacrosPanel();
                if (panel != null)
                {
                    panel.Children.Clear();
                    if (s.SshMacros != null)
                    {
                        foreach (var m in s.SshMacros)
                        {
                            if (string.IsNullOrWhiteSpace(m?.Name) || string.IsNullOrWhiteSpace(m.Command)) continue;
                            var btn = new Button { Content = m.Name, Margin = new Thickness(0,0,6,6), MinWidth = 120, Padding = new Thickness(8,4,8,4), Tag = m.Command };
                            btn.Click += MacroButton_Click;
                            panel.Children.Add(btn);
                        }
                    }
                }
            }
            catch { }
        }

        private void MacroButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button b && b.Tag is string cmd)
            {
                _ = ExecuteCommandAsync(cmd);
            }
        }

        private async Task ExecuteCommandAsync(string cmd)
        {
            if (!_connected) { AppendOutput(LocalizationManager.GetString("SSH.NotConnected","Not connected.") + "\n", Colors.Orange); return; }
            try
            {
                if (_mode == SshMode.SshNet)
                {
                    var shellType = _shell?.GetType();
                    var write = shellType?.GetMethod("WriteLine", BindingFlags.Public | BindingFlags.Instance, null, new[] { typeof(string) }, null);
                    write?.Invoke(_shell, new object[] { cmd });
                }
                else if (_mode == SshMode.SshExe)
                {
                    _sshInput?.WriteLine(cmd);
                    _sshInput?.Flush();
                }
            }
            catch (Exception ex)
            {
                AppendOutput(string.Format(LocalizationManager.GetString("SSH.SendError", "Send error: {0}"), ex.Message) + "\n", Colors.Red);
            }
        }

        private void SetStatus(string text) { if (txtStatus != null) txtStatus.Text = text; }
        private void AppendOutput(string text, Color color)
        {
            if (txtTerminal == null) return;
            var tr = new TextRange(txtTerminal.Document.ContentEnd, txtTerminal.Document.ContentEnd) { Text = text };
            SolidColorBrush brush;
            if (color == Colors.White)
            {
                var bg = txtTerminal.Background as SolidColorBrush;
                var c = bg?.Color ?? Colors.White;
                double luminance = (0.299 * c.R + 0.587 * c.G + 0.114 * c.B) / 255.0;
                brush = luminance > 0.5 ? Brushes.Black : Brushes.White;
            }
            else
            {
                brush = new SolidColorBrush(color);
            }
            tr.ApplyPropertyValue(TextElement.ForegroundProperty, brush);
            txtTerminal.ScrollToEnd();
        }

        private async void btnConnect_Click(object sender, RoutedEventArgs e) => await ConnectAsync();
        private async void txtAddress_KeyDown(object sender, KeyEventArgs e) { if (e.Key == Key.Enter) { e.Handled = true; await ConnectAsync(); } }
        private (string user, string host, int port) ParseAddress(string addr)
        {
            string user = null; string hostPart = addr; int port = 22;
            var at = addr.IndexOf('@');
            if (at >= 0) { user = addr.Substring(0, at); hostPart = addr.Substring(at + 1); }
            var colon = hostPart.LastIndexOf(':');
            if (colon > 0 && colon < hostPart.Length - 1 && int.TryParse(hostPart.Substring(colon + 1), out var p)) { port = p; hostPart = hostPart.Substring(0, colon); }
            return (user, hostPart, port);
        }

        private async Task ConnectAsync()
        {
            if (_connected)
            {
                AppendOutput(LocalizationManager.GetString("SSH.AlreadyConnected", "Already connected.") + "\n", Colors.Orange);
                return;
            }
            var addr = txtAddress?.Text?.Trim();
            if (string.IsNullOrEmpty(addr))
            {
                AppendOutput(LocalizationManager.GetString("SSH.EnterAddress", "Enter address (user@host or host[:port]).") + "\n", Colors.Orange);
                return;
            }
            var (user, host, port) = ParseAddress(addr);
            if (string.IsNullOrWhiteSpace(host))
            {
                AppendOutput(LocalizationManager.GetString("SSH.InvalidAddress", "Invalid address.") + "\n", Colors.Red);
                return;
            }
            try
            {
                SetStatus(LocalizationManager.GetString("SSH.Status.Connecting", "Connecting..."));
                AppendOutput(string.Format(LocalizationManager.GetString("SSH.ConnectingTo", "Connecting to {0}..."), addr) + "\n", Colors.Cyan);
                if (TryLoadSshNet(out var sshTypes))
                {
                    if (string.IsNullOrWhiteSpace(user))
                        user = Prompt(LocalizationManager.GetString("SSH.UsernamePrompt", "Username:"), initial: Environment.UserName, isPassword: false);
                    if (user == null)
                    {
                        AppendOutput(LocalizationManager.GetString("SSH.Cancelled", "Cancelled.") + "\n", Colors.Orange);
                        SetStatus(LocalizationManager.GetString("SSH.Status.Ready", "Ready."));
                        return;
                    }
                    var password = Prompt(string.Format(LocalizationManager.GetString("SSH.PasswordPrompt", "Password for {0}:"), $"{user}@{host}"), initial: string.Empty, isPassword: true);
                    if (password == null)
                    {
                        AppendOutput(LocalizationManager.GetString("SSH.Cancelled", "Cancelled.") + "\n", Colors.Orange);
                        SetStatus(LocalizationManager.GetString("SSH.Status.Ready", "Ready."));
                        return;
                    }
                    await Task.Run(() =>
                    {
                        _client = Activator.CreateInstance(sshTypes.SshClient, host, port, user, password);
                        sshTypes.SshClient_Get_Connect.Invoke(_client, null);
                        _shell = sshTypes.SshClient_Get_CreateShellStream.Invoke(_client, new object[] { "xterm", 80u, 24u, 800u, 600u, 1024 });
                    });
                    if (_client == null || _shell == null)
                        throw new InvalidOperationException("SSH.NET shell creation failed.");

                    _mode = SshMode.SshNet;
                    _readCts = new CancellationTokenSource();
                    _ = Task.Run(() => ReadShellLoop(sshTypes, _readCts.Token));
                }
                else
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = "ssh",
                        Arguments = "-tt " + addr,
                        UseShellExecute = false,
                        RedirectStandardInput = true,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        CreateNoWindow = true
                    };
                    _sshProcess = new Process { StartInfo = psi, EnableRaisingEvents = true };
                    _sshProcess.Exited += (s, e) => Dispatcher.Invoke(() => OnDisconnected());
                    if (!_sshProcess.Start())
                    {
                        AppendOutput("ssh start failed.\n", Colors.Red);
                        SetStatus(LocalizationManager.GetString("SSH.Status.Error", "Error"));
                        return;
                    }
                    _readCts = new CancellationTokenSource();
                    _ = Task.Run(() => ReadSshExeAsync(_readCts.Token));
                    _sshInput = _sshProcess.StandardInput;
                    _mode = SshMode.SshExe;
                }
                _connected = true;
                btnConnect.IsEnabled = false;
                btnDisconnect.IsEnabled = true;
                txtAddress.IsEnabled = false;
                SetStatus(LocalizationManager.GetString("SSH.Status.Connected", "Connected"));
            }
            catch (Exception ex)
            {
                AppendOutput(string.Format(LocalizationManager.GetString("SSH.ConnectionError", "Connection error: {0}"), ex.Message) + "\n", Colors.Red);
                SetStatus(LocalizationManager.GetString("SSH.Status.Error", "Error"));
                Cleanup();
            }
        }

        private void OnShellOutputLine(string line)
        {
            var clean = StripAnsi(line);
            if (_captureLog)
            {
                if (!string.IsNullOrEmpty(_captureEndMarker) && clean.Contains(_captureEndMarker))
                {
                    _captureLog = false;
                    try
                    {
                        var viewer = GetLogViewer();
                        if (viewer != null) viewer.Text = _logBuffer.ToString();
                        _logBuffer.Clear();
                    }
                    catch { }
                    return;
                }
                _logBuffer.AppendLine(clean);
                return;
            }
            AppendOutput(clean + "\n", Colors.White);
        }

        private void Cleanup()
        {
            try { _readCts?.Cancel(); } catch { }
            try { (_shell as IDisposable)?.Dispose(); } catch { }
            try { (_client as IDisposable)?.Dispose(); } catch { }
            try { _sshInput?.Dispose(); } catch { }
            try { _sshProcess?.Dispose(); } catch { }
            _readCts = null; _shell = null; _client = null; _sshInput = null; _sshProcess = null; _mode = SshMode.None; _connected = false;
            _captureLog = false; _captureEndMarker = null; _logBuffer.Clear();
            StopFollow();
        }

        private void ReadShellLoop((Type SshClient, Type ShellStream, MethodInfo SshClient_Get_Connect, MethodInfo SshClient_Get_CreateShellStream, MethodInfo ShellStream_Read, MethodInfo ShellStream_WriteLine) types, CancellationToken token)
        {
            try
            {
                var buffer = new byte[4096]; var enc = new UTF8Encoding(false);
                while (!token.IsCancellationRequested && _shell != null)
                {
                    int read = (int)types.ShellStream_Read.Invoke(_shell, new object[] { buffer, 0, buffer.Length });
                    if (read > 0)
                    {
                        var text = enc.GetString(buffer, 0, read);
                        Dispatcher.Invoke(() => OnShellChunk(text));
                    }
                    else Thread.Sleep(10);
                }
            }
            catch { }
        }

        private async Task ReadSshExeAsync(CancellationToken token)
        {
            try
            {
                var outReader = _sshProcess.StandardOutput;
                var errReader = _sshProcess.StandardError;
                var outTask = Task.Run(async () =>
                {
                    var buf = new char[2048];
                    while (!token.IsCancellationRequested && !_sshProcess.HasExited)
                    {
                        int n = await outReader.ReadAsync(buf, 0, buf.Length).ConfigureAwait(false);
                        if (n > 0)
                        {
                            var text = new string(buf, 0, n);
                            Dispatcher.Invoke(() => OnShellChunk(text));
                        }
                        else await Task.Delay(10).ConfigureAwait(false);
                    }
                }, token);
                var errTask = Task.Run(async () =>
                {
                    var buf = new char[2048];
                    while (!token.IsCancellationRequested && !_sshProcess.HasExited)
                    {
                        int n = await errReader.ReadAsync(buf, 0, buf.Length).ConfigureAwait(false);
                        if (n > 0)
                        {
                            var text = new string(buf, 0, n);
                            Dispatcher.Invoke(() => OnShellChunk(text));
                        }
                        else await Task.Delay(10).ConfigureAwait(false);
                    }
                }, token);
                await Task.WhenAny(Task.WhenAll(outTask, errTask), Task.Run(() => { while (!_sshProcess.HasExited && !token.IsCancellationRequested) Thread.Sleep(50); }));
            }
            catch { }
        }

        private void OnShellChunk(string chunk)
        {
            var clean = StripAnsi(chunk);
            if (_captureLog)
            {
                if (!string.IsNullOrEmpty(_captureEndMarker))
                {
                    var idx = clean.IndexOf(_captureEndMarker, StringComparison.Ordinal);
                    if (idx >= 0)
                    {
                        _logBuffer.Append(clean.Substring(0, idx));
                        _captureLog = false;
                        try { var viewer = GetLogViewer(); if (viewer != null) viewer.Text = _logBuffer.ToString(); } catch { }
                        _logBuffer.Clear();
                        return;
                    }
                }
                _logBuffer.Append(clean);
                return;
            }
            AppendOutput(clean, Colors.White);
        }

        private bool TryLoadSshNet(out (Type SshClient, Type ShellStream, MethodInfo SshClient_Get_Connect, MethodInfo SshClient_Get_CreateShellStream, MethodInfo ShellStream_Read, MethodInfo ShellStream_WriteLine) types)
        {
            types = (null, null, null, null, null, null);
            try
            {
                var sshClientType = Type.GetType("Renci.SshNet.SshClient, Renci.SshNet", false);
                var shellStreamType = Type.GetType("Renci.SshNet.ShellStream, Renci.SshNet", false);
                if (sshClientType == null || shellStreamType == null) return false;
                var miConnect = sshClientType.GetMethod("Connect", BindingFlags.Public | BindingFlags.Instance);
                var miCreateShell = sshClientType.GetMethod("CreateShellStream", BindingFlags.Public | BindingFlags.Instance, null, new Type[] { typeof(string), typeof(uint), typeof(uint), typeof(uint), typeof(uint), typeof(int) }, null);
                var miRead = shellStreamType.GetMethod("Read", BindingFlags.Public | BindingFlags.Instance, null, new Type[] { typeof(byte[]), typeof(int), typeof(int) }, null);
                var miWriteLine = shellStreamType.GetMethod("WriteLine", BindingFlags.Public | BindingFlags.Instance, null, new[] { typeof(string) }, null);
                if (miConnect == null || miCreateShell == null || miRead == null || miWriteLine == null) return false;
                types = (sshClientType, shellStreamType, miConnect, miCreateShell, miRead, miWriteLine);
                return true;
            }
            catch { return false; }
        }

        private void btnDisconnect_Click(object sender, RoutedEventArgs e) => Disconnect();
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
                    try { _readCts?.Cancel(); } catch { }
                    if (_sshProcess != null && !_sshProcess.HasExited)
                    {
                        try { _sshInput?.WriteLine("exit"); _sshInput?.Flush(); } catch { }
                        if (!_sshProcess.WaitForExit(1000)) _sshProcess.Kill();
                    }
                }
            }
            catch { }
            finally
            {
                Cleanup();
                btnConnect.IsEnabled = true; btnDisconnect.IsEnabled = false; txtAddress.IsEnabled = true;
                SetStatus(LocalizationManager.GetString("SSH.Status.Disconnected", "Disconnected"));
            }
        }

        private void txtInput_KeyDown(object sender, KeyEventArgs e) { if (e.Key == Key.Enter) { e.Handled = true; SendCurrentInput(); } }
        private void btnSend_Click(object sender, RoutedEventArgs e) => SendCurrentInput();
        private void SendCurrentInput()
        {
            if (!_connected) return;
            var text = txtInput.Text; txtInput.Clear();
            _ = ExecuteCommandAsync(text);
        }

        private string Prompt(string message, string initial, bool isPassword)
        {
            var win = new Window { Title = LocalizationManager.GetString("SSH.LoginTitle", "SSH Login"), Width = 420, Height = 160, WindowStartupLocation = WindowStartupLocation.CenterOwner, ResizeMode = ResizeMode.NoResize, Owner = Application.Current?.MainWindow };
            var grid = new Grid { Margin = new Thickness(12) };
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var lbl = new TextBlock { Text = message, Margin = new Thickness(0, 0, 0, 8) }; Grid.SetRow(lbl, 0);
            Control input;
            if (isPassword) { var pb = new PasswordBox { Margin = new Thickness(0, 0, 0, 8) }; pb.Password = initial ?? string.Empty; input = pb; }
            else { var tb = new TextBox { Margin = new Thickness(0, 0, 0, 8), Text = initial ?? string.Empty }; input = tb; }
            Grid.SetRow(input, 1);
            var panel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var ok = new Button { Content = LocalizationManager.GetString("Dialog.OK", "OK"), Width = 80, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
            var cancel = new Button { Content = LocalizationManager.GetString("Settings.Cancel", "Cancel"), Width = 80, IsCancel = true };
            ok.Click += (_, __) => { win.DialogResult = true; win.Close(); };
            cancel.Click += (_, __) => { win.DialogResult = false; win.Close(); };
            panel.Children.Add(ok); panel.Children.Add(cancel);
            Grid.SetRow(panel, 2);
            grid.Children.Add(lbl); grid.Children.Add(input); grid.Children.Add(panel);
            win.Content = grid; win.ShowInTaskbar = false;
            var result = win.ShowDialog(); if (result != true) return null;
            return isPassword ? ((PasswordBox)input).Password : ((TextBox)input).Text;
        }

        protected override void OnVisualParentChanged(DependencyObject oldParent)
        { base.OnVisualParentChanged(oldParent); if (oldParent != null && VisualParent == null) Disconnect(); }
        private void OnDisconnected() { Disconnect(); }

        private void btnClearLog_Click(object sender, RoutedEventArgs e)
        { var viewer = GetLogViewer(); if (viewer != null) viewer.Text = string.Empty; }

        private async void btnLoadLog_Click(object sender, RoutedEventArgs e)
        {
            if (!_connected) { MessageBox.Show(LocalizationManager.GetString("SSH.NotConnected", "Not connected.")); return; }
            var combo = GetLogCombo(); var path = combo?.SelectedItem as string; if (string.IsNullOrWhiteSpace(path)) return;
            var chk = GetFollowCheckbox(); if (chk != null && chk.IsChecked == true) { await StartFollowAsync(path); }
            else
            {
                try
                {
                    _logBuffer.Clear();
                    _captureEndMarker = $"__VIVIT_LOG_END_{Guid.NewGuid().ToString("N").ToUpperInvariant()}__";
                    _captureLog = true;
                    var viewer = GetLogViewer(); if (viewer != null) viewer.Text = string.Empty;
                    var cmd = $"tail -n 1000 {EscapePathForShell(path)} 2>&1; echo {_captureEndMarker}";
                    await ExecuteCommandAsync(cmd);
                }
                catch (Exception ex)
                {
                    _captureLog = false; AppendOutput("Log load error: " + ex.Message + "\n", Colors.Red);
                }
            }
        }

        private async void chkFollow_Checked(object sender, RoutedEventArgs e)
        { var combo = GetLogCombo(); var path = combo?.SelectedItem as string; if (!_connected || string.IsNullOrWhiteSpace(path)) return; await StartFollowAsync(path); }
        private void chkFollow_Unchecked(object sender, RoutedEventArgs e) { StopFollow(); }

        private async Task StartFollowAsync(string path)
        {
            StopFollow();
            var viewer = GetLogViewer(); if (viewer != null) viewer.Text = string.Empty;
            var btn = GetLoadButton(); if (btn != null) btn.IsEnabled = false;
            _isFollowing = true; _followCts = new CancellationTokenSource();
            try
            {
                if (_mode == SshMode.SshNet && _client != null && TryLoadSshNet(out var types))
                {
                    _logShell = types.SshClient_Get_CreateShellStream.Invoke(_client, new object[] { "xterm", 80u, 24u, 800u, 600u, 1024 });
                    types.ShellStream_WriteLine.Invoke(_logShell, new object[] { $"tail -n 200 -F {EscapePathForShell(path)} 2>&1" });
                    _ = Task.Run(() => ReadFollowLoop(types, _followCts.Token));
                }
                else if (_mode == SshMode.SshExe)
                {
                    var addr = txtAddress?.Text?.Trim(); if (string.IsNullOrWhiteSpace(addr)) throw new InvalidOperationException("No address");
                    var psi = new ProcessStartInfo
                    {
                        FileName = "ssh",
                        Arguments = "-tt " + addr + " " + $"tail -n 200 -F {EscapePathForShell(path)} 2>&1",
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        CreateNoWindow = true
                    };
                    _logSshProcess = new Process { StartInfo = psi, EnableRaisingEvents = true };
                    _logSshProcess.OutputDataReceived += (s, e2) => { if (!string.IsNullOrEmpty(e2.Data)) Dispatcher.Invoke(() => AppendToLogViewer(StripAnsi(e2.Data + "\n"))); };
                    _logSshProcess.ErrorDataReceived += (s, e2) => { if (!string.IsNullOrEmpty(e2.Data)) Dispatcher.Invoke(() => AppendToLogViewer(StripAnsi(e2.Data + "\n"))); };
                    _logSshProcess.Exited += (s, e2) => Dispatcher.Invoke(() => StopFollow());
                    _logSshProcess.Start();
                    _logSshProcess.BeginOutputReadLine();
                    _logSshProcess.BeginErrorReadLine();
                }
                else
                {
                    await ExecuteCommandAsync($"tail -n 200 {EscapePathForShell(path)} 2>&1");
                }
            }
            catch
            {
                StopFollow();
            }
        }

        private void ReadFollowLoop((Type SshClient, Type ShellStream, MethodInfo SshClient_Get_Connect, MethodInfo SshClient_Get_CreateShellStream, MethodInfo ShellStream_Read, MethodInfo ShellStream_WriteLine) types, CancellationToken token)
        {
            try
            {
                var buffer = new byte[4096]; var enc = new UTF8Encoding(false);
                while (!token.IsCancellationRequested && _logShell != null)
                {
                    int read = (int)types.ShellStream_Read.Invoke(_logShell, new object[] { buffer, 0, buffer.Length });
                    if (read > 0)
                    {
                        var text = enc.GetString(buffer, 0, read);
                        Dispatcher.Invoke(() => AppendToLogViewer(StripAnsi(text)));
                    }
                    else Thread.Sleep(10);
                }
            }
            catch { }
        }

        private void AppendToLogViewer(string text)
        { var viewer = GetLogViewer(); if (viewer == null) return; viewer.AppendText(text); viewer.ScrollToEnd(); }

        private void StopFollow()
        {
            _isFollowing = false;
            try { _followCts?.Cancel(); } catch { }
            _followCts = null;
            try { (_logShell as IDisposable)?.Dispose(); } catch { }
            _logShell = null;
            try { if (_logSshProcess != null && !_logSshProcess.HasExited) _logSshProcess.Kill(); } catch { }
            _logSshProcess = null;
            var btn = GetLoadButton(); if (btn != null) btn.IsEnabled = true;
            var chk = GetFollowCheckbox(); if (chk != null) chk.IsChecked = false;
        }

        private static string EscapePathForShell(string path)
        {
            if (string.IsNullOrEmpty(path)) return path;
            return "'" + path.Replace("'", "'\\''") + "'";
        }
    }
}