using System;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Interop;
using Vivit_Control_Center.Settings;
using Vivit_Control_Center.Localization;

namespace Vivit_Control_Center
{
    public partial class App : Application
    {
        private const string WinLogonKeyPath = @"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon";
        public static bool IsShellMode { get; private set; }
        private TaskbarWindow _taskbarWindow;

        private const int ShellTaskbarHeight = 40; // keep consistent with TaskbarWindow & MainWindow

        #region Win32 WorkArea
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SystemParametersInfo(int uiAction, int uiParam, ref RECT pvParam, int fWinIni);

        private const int SPI_SETWORKAREA = 0x002F; // 47
        private const int SPIF_SENDCHANGE = 0x02;
        #endregion

        #region Win + D hook
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int VK_LWIN = 0x5B;
        private const int VK_RWIN = 0x5C;
        private const int VK_D = 0x44;
        private const int SW_SHOW = 5;
        private const int SW_RESTORE = 9;

        private static IntPtr _keyboardHookHandle = IntPtr.Zero;
        private static LowLevelKeyboardProc _keyboardProcDelegate; // keep reference

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct KBDLLHOOKSTRUCT
        {
            public int vkCode;
            public int scanCode;
            public int flags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private void InstallKeyboardHookIfNeeded()
        {
            if (!IsShellMode || _keyboardHookHandle != IntPtr.Zero) return;
            _keyboardProcDelegate = KeyboardHookProc;
            _keyboardHookHandle = SetWindowsHookEx(WH_KEYBOARD_LL, _keyboardProcDelegate, IntPtr.Zero, 0);
        }

        private IntPtr KeyboardHookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (nCode >= 0)
                {
                    int msg = wParam.ToInt32();
                    if (msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN)
                    {
                        var data = (KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(KBDLLHOOKSTRUCT));
                        if (data.vkCode == VK_D)
                        {
                            bool winDown = (GetAsyncKeyState(VK_LWIN) & 0x8000) != 0 || (GetAsyncKeyState(VK_RWIN) & 0x8000) != 0;
                            if (winDown)
                            {
                                BringMainWindowToFront();
                                // Swallow to prevent default Win+D minimize-all when in custom shell
                                return (IntPtr)1;
                            }
                        }
                    }
                }
            }
            catch { }
            return CallNextHookEx(_keyboardHookHandle, nCode, wParam, lParam);
        }

        private void BringMainWindowToFront()
        {
            try
            {
                var main = Current?.MainWindow;
                if (main == null) return;
                var hwnd = new WindowInteropHelper(main).Handle;
                if (hwnd == IntPtr.Zero) return;
                if (IsIconic(hwnd))
                    ShowWindow(hwnd, SW_RESTORE);
                else
                    ShowWindow(hwnd, SW_SHOW);
                SetForegroundWindow(hwnd);
            }
            catch { }
        }
        #endregion

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            // Load settings early to apply language
            AppSettings loadedSettings = null;
            try { loadedSettings = AppSettings.Load(); } catch { }
            if (loadedSettings != null)
            {
                try { LocalizationManager.ApplyLanguage(loadedSettings.Language ?? "en"); } catch { }
            }

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

            // Create custom taskbar if running as shell and adjust desktop work area
            if (IsShellMode)
            {
                SetShellWorkArea();
                this.Dispatcher.InvokeAsync(() =>
                {
                    try
                    {
                        _taskbarWindow = new TaskbarWindow();
                        _taskbarWindow.Show();
                        InstallKeyboardHookIfNeeded();
                    }
                    catch { }
                });
            }
        }

        protected override void OnExit(ExitEventArgs e)
        {
            try { _taskbarWindow?.Close(); } catch { }
            if (IsShellMode)
            {
                // restore full screen work area when exiting custom shell
                RestoreWorkArea();
            }
            if (_keyboardHookHandle != IntPtr.Zero)
            {
                try { UnhookWindowsHookEx(_keyboardHookHandle); } catch { }
                _keyboardHookHandle = IntPtr.Zero;
            }
            base.OnExit(e);
        }

        private void SetShellValue(string value)
        {
            using (var key = Registry.LocalMachine.OpenSubKey(WinLogonKeyPath, true))
            {
                if (key == null) throw new InvalidOperationException("Winlogon Key nicht gefunden");
                key.SetValue("Shell", value, RegistryValueKind.String);
            }
        }

        private void SetShellWorkArea()
        {
            try
            {
                var width = (int)SystemParameters.PrimaryScreenWidth;
                var height = (int)SystemParameters.PrimaryScreenHeight;
                var rect = new RECT
                {
                    left = 0,
                    top = 0,
                    right = width,
                    bottom = height - ShellTaskbarHeight
                };
                SystemParametersInfo(SPI_SETWORKAREA, 0, ref rect, SPIF_SENDCHANGE);
            }
            catch { }
        }

        private void RestoreWorkArea()
        {
            try
            {
                var width = (int)SystemParameters.PrimaryScreenWidth;
                var height = (int)SystemParameters.PrimaryScreenHeight;
                var rect = new RECT { left = 0, top = 0, right = width, bottom = height };
                SystemParametersInfo(SPI_SETWORKAREA, 0, ref rect, SPIF_SENDCHANGE);
            }
            catch { }
        }
    }
}
