// PATCH: Entfernt das erzwungene Setzen von IsShellMode = true (falsch) und belässt nur die echte Erkennung.
//        Keine weitere Änderung hier außer Entfernen der Zeile.
using System;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Interop;
using Vivit_Control_Center.Settings;
using Vivit_Control_Center.Localization;
using Vivit_Control_Center.Services;
using System.Threading.Tasks;

namespace Vivit_Control_Center
{
    public partial class App : Application
    {
        private const string WinLogonKeyPath = @"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon";
        public static bool IsShellMode { get; private set; }
        private TaskbarWindow _taskbarWindow;

        // Dynamic taskbar height determined at runtime to match Windows taskbar height
        public static int ShellTaskbarHeightPx { get; private set; } = 40;

        #region Win32
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int left; public int top; public int right; public int bottom; }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SystemParametersInfo(int uiAction, int uiParam, ref RECT pvParam, int fWinIni);
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        [DllImport("user32.dll")] private static extern int GetSystemMetrics(int nIndex);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        private const int SM_CXSCREEN = 0;
        private const int SM_CYSCREEN = 1;
        private const int SPI_SETWORKAREA = 0x002F;
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
        private static LowLevelKeyboardProc _keyboardProcDelegate;

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
        [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int vKey);
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll", SetLastError = true)] private static extern bool SetWindowPos(IntPtr hWnd, IntPtr insertAfter, int X, int Y, int cx, int cy, uint flags);
        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOACTIVATE = 0x0010;

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

        protected override async void OnStartup(StartupEventArgs e)
         {
            base.OnStartup(e);

            AppSettings loadedSettings = null;
            try { loadedSettings = AppSettings.Load(); } catch { }
            if (loadedSettings != null)
            {
                try { LocalizationManager.ApplyLanguage(loadedSettings.Language ?? "en"); } catch { }
            }

            try
            {
                // Correctly detect shell mode from HKLM where Winlogon Shell is stored
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

            // Detect native taskbar height (before changing work area)
            ShellTaskbarHeightPx = DetectNativeTaskbarHeight();

            if (IsShellMode)
            {
                SetShellWorkArea();
                Dispatcher.InvokeAsync(() =>
                {
                    try
                    {
                        _taskbarWindow = new TaskbarWindow();
                        _taskbarWindow.Show();
                        InstallKeyboardHookIfNeeded();
                        //_ = LaunchExplorerForTrayAsync();
                    }
                    catch { }
                });
            }

            // Start email background sync service early and perform initial refresh
            try
            {
                EmailSyncService.Current.Initialize(loadedSettings ?? AppSettings.Load());
                await EmailSyncService.Current.SafeRefreshAllAsync();
            }
            catch { }

            try
            {
                if (e.Args != null && e.Args.Length > 0)
                {
                    if (Array.Exists(e.Args, a => string.Equals(a, "--restore-shell", StringComparison.OrdinalIgnoreCase)))
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

       

            TrayIconService.Current.Initialize(loadedSettings?.Theme);
        }

        private static int DetectNativeTaskbarHeight()
        {
            try
            {
                int screenH = GetSystemMetrics(SM_CYSCREEN);
                var hwnd = FindWindow("Shell_TrayWnd", null);
                if (hwnd != IntPtr.Zero)
                {
                    if (GetWindowRect(hwnd, out var rc))
                    {
                        int height = Math.Abs(rc.bottom - rc.top);
                        if (height >= 24 && height <= screenH / 2)
                            return height;
                    }
                }
                // fallback: difference between screen and workarea in pixels (approximate, rarely used)
                int waApprox = (int)Math.Round(SystemParameters.PrimaryScreenHeight - SystemParameters.WorkArea.Height);
                if (waApprox >= 24 && waApprox <= screenH / 2) return waApprox;
            }
            catch { }
            return 40;
        }

        private async Task LaunchExplorerForTrayAsync()
        {
            try
            {
                if (!Process.GetProcessesByName("explorer").Any())
                {
                    try { Process.Start("explorer.exe"); } catch { }
                }

                for (int i = 0; i < 20; i++)
                {
                    await Task.Delay(300);
                    var hTrayTest = FindWindow("Shell_TrayWnd", null);
                    if (hTrayTest != IntPtr.Zero) break;
                }

                var hTray = FindWindow("Shell_TrayWnd", null);
                if (hTray != IntPtr.Zero)
                {
                    int screenW = GetSystemMetrics(SM_CXSCREEN);
                    int screenH = GetSystemMetrics(SM_CYSCREEN);
                    try { SetWindowPos(hTray, IntPtr.Zero, -5000, screenH - 2, screenW, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_NOACTIVATE); } catch { }
                }
            }
            catch { }
        }

        protected override void OnExit(ExitEventArgs e)
        {
            try { _taskbarWindow?.Close(); } catch { }
            if (IsShellMode) RestoreWorkArea();
            if (_keyboardHookHandle != IntPtr.Zero)
            {
                try { UnhookWindowsHookEx(_keyboardHookHandle); } catch { }
                _keyboardHookHandle = IntPtr.Zero;
            }
            TrayIconService.Current.Dispose();
            try { EmailSyncService.Current.Dispose(); } catch { }
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
                int widthPx = GetSystemMetrics(SM_CXSCREEN);
                int heightPx = GetSystemMetrics(SM_CYSCREEN);
                var rect = new RECT { left = 0, top = 0, right = widthPx, bottom = heightPx - ShellTaskbarHeightPx };
                SystemParametersInfo(SPI_SETWORKAREA, 0, ref rect, SPIF_SENDCHANGE);
            }
            catch { }
        }

        public static void UpdateShellWorkAreaLeft(int leftPx)
        {
            try
            {
                if (!IsShellMode) return;
                int widthPx = GetSystemMetrics(SM_CXSCREEN);
                int heightPx = GetSystemMetrics(SM_CYSCREEN);
                if (leftPx < 0) leftPx = 0;
                if (leftPx > widthPx - 50) leftPx = Math.Max(0, widthPx - 50);
                var rect = new RECT { left = leftPx, top = 0, right = widthPx, bottom = heightPx - ShellTaskbarHeightPx };
                SystemParametersInfo(SPI_SETWORKAREA, 0, ref rect, SPIF_SENDCHANGE);
            }
            catch { }
        }

        private void RestoreWorkArea()
        {
            try
            {
                int widthPx = GetSystemMetrics(SM_CXSCREEN);
                int heightPx = GetSystemMetrics(SM_CYSCREEN);
                var rect = new RECT { left = 0, top = 0, right = widthPx, bottom = heightPx };
                SystemParametersInfo(SPI_SETWORKAREA, 0, ref rect, SPIF_SENDCHANGE);
            }
            catch { }
        }
    }
}