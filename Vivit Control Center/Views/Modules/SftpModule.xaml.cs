using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Documents;
using System.Threading;
using System.ComponentModel;
using System.Globalization;
using System.Diagnostics;
using System.Reflection;
using Vivit_Control_Center.Localization;

namespace Vivit_Control_Center.Views.Modules
{
    public partial class SftpModule : BaseSimpleModule
    {
        private string _localPath;
        private object _sftpClient; // Renci.SshNet.SftpClient (loaded via reflection)
        private bool _connected;

        public SftpModule()
        {
            InitializeComponent();
            _localPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            LoadLocal(_localPath);
            SetStatus(LocalizationManager.GetString("SFTP.Status.Ready", "Bereit."));
        }

        private void SetStatus(string text)
        {
            if (txtStatus != null) txtStatus.Text = text;
        }

        // Local file model
        private class FsItem
        {
            public string Name { get; set; }
            public string FullPath { get; set; }
            public bool IsDir { get; set; }
            public long Size { get; set; }
            public DateTime Modified { get; set; }
            public string SizeDisplay => IsDir ? "<DIR>" : Size.ToString("N0", CultureInfo.CurrentCulture);
            public string ModifiedDisplay => Modified.ToString("G");
        }

        private void LoadLocal(string path)
        {
            try
            {
                var dir = new DirectoryInfo(path);
                if (!dir.Exists) return;
                txtLocalPath.Text = dir.FullName;
                var items = new List<FsItem>();
                items.AddRange(dir.GetDirectories().OrderBy(d => d.Name).Select(d => new FsItem
                {
                    Name = d.Name,
                    FullPath = d.FullName,
                    IsDir = true,
                    Size = 0,
                    Modified = d.LastWriteTime
                }));
                items.AddRange(dir.GetFiles().OrderBy(f => f.Name).Select(f => new FsItem
                {
                    Name = f.Name,
                    FullPath = f.FullName,
                    IsDir = false,
                    Size = f.Length,
                    Modified = f.LastWriteTime
                }));
                lvLocal.ItemsSource = items;
                _localPath = dir.FullName;
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(LocalizationManager.GetString("SFTP.LocalLoadError", "Lokales Verzeichnis kann nicht geladen werden: {0}"), ex.Message));
            }
        }

        private void btnLocalUp_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var parent = Directory.GetParent(_localPath);
                if (parent != null) LoadLocal(parent.FullName);
            }
            catch { }
        }

        private void lvLocal_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var item = lvLocal.SelectedItem as FsItem;
            if (item == null) return;
            if (item.IsDir) LoadLocal(item.FullPath);
            else // download to remote current path
                _ = UploadAsync(item.FullPath);
        }

        // Remote (SFTP) via reflection to avoid hard dependency if SSH.NET missing
        private class RemoteItem
        {
            public string Name { get; set; }
            public string FullName { get; set; }
            public bool IsDir { get; set; }
            public long Size { get; set; }
            public DateTime Modified { get; set; }
            public string SizeDisplay => IsDir ? "<DIR>" : Size.ToString("N0", CultureInfo.CurrentCulture);
            public string ModifiedDisplay => Modified.ToString("G");
        }

        private string _remotePath = "/";
        private (Type SftpClient, Type SftpFile, MethodInfo Connect, MethodInfo Disconnect, MethodInfo ListDirectory, MethodInfo ChangeDirectory, MethodInfo UploadFile, MethodInfo DownloadFile, MethodInfo DeleteFile, MethodInfo CreateDirectory, MethodInfo DeleteDirectory) _sftpRefs;

        private bool TryLoadSshNetSftp()
        {
            try
            {
                var tClient = Type.GetType("Renci.SshNet.SftpClient, Renci.SshNet", throwOnError: false);
                var tFile = Type.GetType("Renci.SshNet.Sftp.SftpFile, Renci.SshNet", throwOnError: false) ?? Type.GetType("Renci.SshNet.SftpFile, Renci.SshNet", false);
                if (tClient == null) return false;
                var miConnect = tClient.GetMethod("Connect");
                var miDisconnect = tClient.GetMethod("Disconnect");
                var miList = tClient.GetMethod("ListDirectory", new[] { typeof(string) });
                var miCwd = tClient.GetMethod("ChangeDirectory", new[] { typeof(string) });
                var miUpload = tClient.GetMethod("UploadFile", new[] { typeof(Stream), typeof(string), typeof(bool) });
                var miDownload = tClient.GetMethod("DownloadFile", new[] { typeof(string), typeof(Stream) });
                var miDelFile = tClient.GetMethod("DeleteFile", new[] { typeof(string) });
                var miMkDir = tClient.GetMethod("CreateDirectory", new[] { typeof(string) });
                var miRmDir = tClient.GetMethod("DeleteDirectory", new[] { typeof(string) });
                _sftpRefs = (tClient, tFile, miConnect, miDisconnect, miList, miCwd, miUpload, miDownload, miDelFile, miMkDir, miRmDir);
                return true;
            }
            catch { return false; }
        }

        private async void btnLogin_Click(object sender, RoutedEventArgs e)
        {
            await LoginAsync();
        }

        private async Task LoginAsync()
        {
            if (_connected) return;
            if (!TryLoadSshNetSftp())
            {
                MessageBox.Show(LocalizationManager.GetString("SFTP.SshNetMissing", "SSH.NET (Renci.SshNet) not found. Please provide the library."));
                return;
            }

            var host = txtHost.Text?.Trim();
            var user = txtUser.Text?.Trim();
            var pass = txtPassword.Password;
            int port = 22; int.TryParse(txtPort.Text, out port);
            if (string.IsNullOrEmpty(host) || string.IsNullOrEmpty(user))
            {
                MessageBox.Show(LocalizationManager.GetString("SFTP.CredentialsRequired", "Host und Benutzer sind erforderlich."));
                return;
            }

            try
            {
                SetStatus(LocalizationManager.GetString("SFTP.Status.Connecting", "Verbinden..."));
                await Task.Run(() =>
                {
                    _sftpClient = Activator.CreateInstance(_sftpRefs.SftpClient, host, port, user, pass);
                    _sftpRefs.Connect.Invoke(_sftpClient, null);
                });
                _connected = true;
                btnLogin.IsEnabled = false; btnLogout.IsEnabled = true;
                txtHost.IsEnabled = txtUser.IsEnabled = txtPassword.IsEnabled = txtPort.IsEnabled = false;
                SetStatus(LocalizationManager.GetString("SFTP.Status.Connected", "Verbunden"));
                await LoadRemoteAsync(".");
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(LocalizationManager.GetString("SFTP.LoginFailed", "Login fehlgeschlagen: {0}"), ex.Message));
                SetStatus(LocalizationManager.GetString("SFTP.Status.Error", "Fehler"));
            }
        }

        private void btnLogout_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _sftpRefs.Disconnect?.Invoke(_sftpClient, null);
            }
            catch { }
            _connected = false;
            btnLogin.IsEnabled = true; btnLogout.IsEnabled = false;
            txtHost.IsEnabled = txtUser.IsEnabled = txtPassword.IsEnabled = txtPort.IsEnabled = true;
            lvRemote.ItemsSource = null;
            txtRemotePath.Text = string.Empty;
            SetStatus(LocalizationManager.GetString("SFTP.Status.Disconnected", "Getrennt"));
        }

        private async Task LoadRemoteAsync(string path)
        {
            if (!_connected || _sftpClient == null) return;
            await Task.Run(() =>
            {
                try
                {
                    _sftpRefs.ChangeDirectory.Invoke(_sftpClient, new object[] { path });
                    var list = _sftpRefs.ListDirectory.Invoke(_sftpClient, new object[] { "." }) as System.Collections.IEnumerable;
                    var items = new List<RemoteItem>();
                    foreach (var f in list)
                    {
                        var t = f.GetType();
                        var name = (string)t.GetProperty("Name").GetValue(f);
                        if (name == "." || name == "..") continue;
                        var isDir = (bool)t.GetProperty("IsDirectory").GetValue(f);
                        var len = (long)t.GetProperty("Length").GetValue(f);
                        var mod = (DateTime)t.GetProperty("LastWriteTime").GetValue(f);
                        items.Add(new RemoteItem
                        {
                            Name = name,
                            FullName = (string)t.GetProperty("FullName").GetValue(f),
                            IsDir = isDir,
                            Size = len,
                            Modified = mod
                        });
                    }
                    items = items.OrderBy(i => i.IsDir ? 0 : 1).ThenBy(i => i.Name).ToList();
                    Dispatcher.Invoke(() =>
                    {
                        lvRemote.ItemsSource = items;
                        txtRemotePath.Text = GetRemotePwd();
                    });
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() => MessageBox.Show(string.Format(LocalizationManager.GetString("SFTP.RemoteLoadFailed", "Remote laden fehlgeschlagen: {0}"), ex.Message)));
                }
            });
        }

        private string GetRemotePwd()
        {
            try
            {
                var prop = _sftpClient.GetType().GetProperty("WorkingDirectory");
                if (prop != null) return prop.GetValue(_sftpClient) as string;
            }
            catch { }
            return _remotePath;
        }

        private async void lvRemote_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var item = lvRemote.SelectedItem as RemoteItem;
            if (item == null) return;
            if (item.IsDir) await LoadRemoteAsync(item.FullName);
            else await DownloadAsync(item.FullName);
        }

        private async void btnRemoteUp_Click(object sender, RoutedEventArgs e)
        {
            if (!_connected) return;
            var cwd = GetRemotePwd();
            var parent = System.IO.Path.GetDirectoryName(cwd.TrimEnd('/').Replace('/', System.IO.Path.DirectorySeparatorChar))?.Replace(System.IO.Path.DirectorySeparatorChar, '/');
            if (string.IsNullOrEmpty(parent)) parent = "/";
            await LoadRemoteAsync(parent);
        }

        private async Task UploadAsync(string localFile)
        {
            if (!_connected) return;
            try
            {
                await Task.Run(() =>
                {
                    using (var fs = File.OpenRead(localFile))
                    {
                        var fileName = System.IO.Path.GetFileName(localFile);
                        var remote = GetRemotePwd().TrimEnd('/') + "/" + fileName;
                        _sftpRefs.UploadFile.Invoke(_sftpClient, new object[] { fs, remote, true });
                    }
                });
                await LoadRemoteAsync(".");
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(LocalizationManager.GetString("SFTP.UploadFailed", "Upload fehlgeschlagen: {0}"), ex.Message));
            }
        }

        private async Task DownloadAsync(string remoteFile)
        {
            if (!_connected) return;
            try
            {
                var target = System.IO.Path.Combine(_localPath, System.IO.Path.GetFileName(remoteFile));
                await Task.Run(() =>
                {
                    using (var fs = File.Create(target))
                    {
                        _sftpRefs.DownloadFile.Invoke(_sftpClient, new object[] { remoteFile, fs });
                    }
                });
                LoadLocal(_localPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(LocalizationManager.GetString("SFTP.DownloadFailed", "Download fehlgeschlagen: {0}"), ex.Message));
            }
        }
    }
}