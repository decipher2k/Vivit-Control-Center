using System;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using System.Diagnostics;

namespace Vivit_Control_Center
{
    public partial class App : Application
    {
        private const string WinLogonKeyPath = @"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon";
        public static bool IsShellMode { get; private set; }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            try
            {
                if (e.Args != null && e.Args.Length > 0)
                {
                    if (e.Args.Contains("--restore-shell", StringComparer.OrdinalIgnoreCase))
                    {
                        SetShellValue("explorer.exe");
                        Current.Shutdown();
                        return;
                    }
                    int idx = Array.FindIndex(e.Args, a => string.Equals(a, "--set-shell", StringComparison.OrdinalIgnoreCase));
                    if (idx >= 0 && idx < e.Args.Length - 1)
                    {
                        var newShell = e.Args[idx + 1];
                        if (!string.IsNullOrWhiteSpace(newShell))
                        {
                            SetShellValue(newShell);
                            Current.Shutdown();
                            return;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                try { MessageBox.Show("Shell Änderung fehlgeschlagen: " + ex.Message); } catch { }
            }

            // Prüfen ob aktuelle EXE als Shell eingetragen ist
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(WinLogonKeyPath, false))
                {
                    var regVal = key?.GetValue("Shell") as string;
                    if (!string.IsNullOrWhiteSpace(regVal))
                    {
                        regVal = regVal.Trim().Trim('"');
                        var exe = Process.GetCurrentProcess().MainModule.FileName.Trim().Trim('"');
                        if (string.Equals(regVal, exe, StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(regVal, System.IO.Path.GetFileName(exe), StringComparison.OrdinalIgnoreCase))
                        {
                            IsShellMode = true;
                        }
                    }
                }
            }
            catch { }
        }

        private void SetShellValue(string value)
        {
            using (var key = Registry.LocalMachine.OpenSubKey(WinLogonKeyPath, true))
            {
                if (key == null) throw new InvalidOperationException("Winlogon Key nicht gefunden");
                key.SetValue("Shell", value, RegistryValueKind.String);
            }
        }
    }
}
