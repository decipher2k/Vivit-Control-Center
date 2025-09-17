using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.IO;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Threading.Tasks;
using System.Windows.Interop;

namespace Vivit_Control_Center
{
    public partial class TaskbarWindow : Window
    {
        private DispatcherTimer _clockTimer;
        private DispatcherTimer _processRefreshTimer;
        private DispatcherTimer _trayRefreshTimer;

        private StackPanel _processIconsPanel;
        private ListBox _windowsListBox;
        private TextBlock _popupProcessTitle;
        private Popup _windowsPopup;
        private StackPanel _trayAreaPanel;

        private static BitmapSource _defaultProcessIcon24;
        private static BitmapSource _defaultWindowIcon16;

        private HwndSource _shellTraySource;             // Shell_TrayWnd
        private HwndSource _shellSecondaryTraySource;    // Shell_SecondaryTrayWnd (multi monitor scenarios / some apps probe it)
        private HwndSource _notifyOverflowSource;        // NotifyIconOverflowWindow (apps sometimes look it up)
        private uint _msgTaskbarCreated;

        #region Default Icons
        private void EnsureDefaultIcons()
        {
            if (_defaultProcessIcon24 == null)
                _defaultProcessIcon24 = CreateDefaultIcon(24, 24, Colors.DodgerBlue);
            if (_defaultWindowIcon16 == null)
                _defaultWindowIcon16 = CreateDefaultIcon(16, 16, Colors.Gray);
        }
        private static BitmapSource CreateDefaultIcon(int w, int h, Color accent)
        {
            var dv = new DrawingVisual();
            using (var dc = dv.RenderOpen())
            {
                var bg = new SolidColorBrush(Color.FromRgb(32, 32, 32));
                dc.DrawRectangle(bg, null, new Rect(0, 0, w, h));
                var pen = new Pen(new SolidColorBrush(accent), Math.Max(1, w / 12.0));
                dc.DrawRectangle(null, pen, new Rect(w * 0.15, h * 0.2, w * 0.7, h * 0.65));
                dc.DrawLine(pen, new Point(w * 0.15, h * 0.35), new Point(w * 0.85, h * 0.35));
                dc.DrawRectangle(new SolidColorBrush(accent), null, new Rect(w * 0.3, h * 0.45, w * 0.2, h * 0.2));
            }
            var rtb = new RenderTargetBitmap(w, h, 96, 96, PixelFormats.Pbgra32);
            rtb.Render(dv); rtb.Freeze(); return rtb;
        }
        #endregion

        public TaskbarWindow()
        {
            InitializeComponent();
            Loaded += TaskbarWindow_Loaded;
            Deactivated += (s, e) => { Topmost = true; };
            Loaded += (s, e) =>
            {
                _processIconsPanel = (StackPanel)FindName("ProcessIconsPanel");
                _windowsListBox = (ListBox)FindName("WindowsListBox");
                _popupProcessTitle = (TextBlock)FindName("PopupProcessTitle");
                _windowsPopup = (Popup)FindName("WindowsPopup");
                _trayAreaPanel = (StackPanel)FindName("TrayAreaPanel");
            };
        }

        private void AppButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                (Application.Current?.MainWindow as MainWindow)?.ActivateModule("Programs");
            }
            catch { }
        }

        private void TaskbarWindow_Loaded(object sender, RoutedEventArgs e)
        {
            EnsureDefaultIcons();
            PositionAtBottom();

            if (App.IsShellMode)
                InitShellTrayHost();

            _clockTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(10) };
            _clockTimer.Tick += (_, __) => UpdateClock(); _clockTimer.Start(); UpdateClock();

            _processRefreshTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(3) };
            _processRefreshTimer.Tick += (_, __) => RefreshProcessIcons(); _processRefreshTimer.Start(); RefreshProcessIcons();

            _trayRefreshTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(10) };
            _trayRefreshTimer.Tick += (_, __) => RefreshTrayIcons(); _trayRefreshTimer.Start(); RefreshTrayIcons();

            try { var vol = (int)(GetMasterVolume() * 100); VolumeSlider.Value = vol; UpdateVolumeIcon(vol); } catch { }
        }

        private void PositionAtBottom()
        { Left = 0; Width = SystemParameters.PrimaryScreenWidth; Height = 40; Top = SystemParameters.PrimaryScreenHeight - Height; }
        private void UpdateClock() { var now = DateTime.Now; TimeText.Text = now.ToString("HH:mm"); DateText.Text = now.ToString("dd.MM.yyyy"); }

        #region Process + Window enumeration (unchanged core)
        private class WindowInfo { public IntPtr Hwnd; public string Title; public BitmapSource Icon; }
        private class ProcessInfo { public Process Process; public List<WindowInfo> Windows = new List<WindowInfo>(); public string DisplayName => string.IsNullOrWhiteSpace(Process?.MainWindowTitle) ? Process?.ProcessName : Process?.MainWindowTitle; public IntPtr FirstWindow => Windows.FirstOrDefault()?.Hwnd ?? IntPtr.Zero; }
        private readonly Dictionary<int, ProcessInfo> _processCache = new Dictionary<int, ProcessInfo>();
        private void RefreshProcessIcons()
        {
            if (_processIconsPanel == null) return; EnsureDefaultIcons();
            try
            {
                var windows = EnumerateTopLevelWindows();
                var procGroups = windows.GroupBy(w => w.ProcessId);
                var currentPids = new HashSet<int>(procGroups.Select(g => g.Key));
                foreach (var stale in _processCache.Keys.Where(k => !currentPids.Contains(k)).ToList()) _processCache.Remove(stale);
                foreach (var g in procGroups)
                {
                    Process proc = null; try { proc = Process.GetProcessById(g.Key); } catch { }
                    if (proc == null) continue;
                    if (!_processCache.TryGetValue(g.Key, out var pinfo)) { pinfo = new ProcessInfo { Process = proc }; _processCache[g.Key] = pinfo; }
                    pinfo.Windows = g.Select(w => new WindowInfo { Hwnd = w.Hwnd, Title = w.Title, Icon = GetWindowIcon(w.Hwnd) ?? _defaultWindowIcon16 }).ToList();
                }
                _processIconsPanel.Children.Clear();
                foreach (var p in _processCache.Values.OrderBy(p => p.DisplayName)) _processIconsPanel.Children.Add(CreateProcessButton(p));
            }
            catch { }
        }
        private Button CreateProcessButton(ProcessInfo pinfo)
        {
            var btn = new Button { Width = 34, Height = 34, Padding = new Thickness(0), Margin = new Thickness(2,3,2,3), Background = Brushes.Transparent, BorderBrush = Brushes.Transparent, ToolTip = pinfo.DisplayName, Tag = pinfo };
            btn.MouseEnter += ProcessButton_MouseEnter; btn.Click += ProcessButton_Click;
            var iconSource = TryGetProcessIcon(pinfo.Process) ?? GetWindowIcon(pinfo.FirstWindow) ?? _defaultProcessIcon24;
            btn.Content = new Image { Source = iconSource, Width = 24, Height = 24, Stretch = Stretch.Uniform, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
            return btn;
        }
        private void ProcessButton_Click(object s, RoutedEventArgs e) { if (s is Button b && b.Tag is ProcessInfo p) { var hwnd = p.FirstWindow; if (hwnd != IntPtr.Zero) RestoreAndActivate(hwnd); } }
        private void ProcessButton_MouseEnter(object s, MouseEventArgs e)
        {
            if (_windowsPopup == null || _windowsListBox == null || _popupProcessTitle == null) return;
            if (!(s is Button b) || !(b.Tag is ProcessInfo pinfo) || pinfo.Windows.Count == 0) return;
            _windowsListBox.ItemsSource = pinfo.Windows; _popupProcessTitle.Text = pinfo.DisplayName; _windowsPopup.PlacementTarget = b; _windowsPopup.IsOpen = true;
        }
        private void WindowsListBox_MouseLeave(object s, MouseEventArgs e) { if (_windowsPopup != null) _windowsPopup.IsOpen = false; }
        private void WindowsListBox_MouseLeftButtonUp(object s, MouseButtonEventArgs e)
        { if (_windowsListBox?.SelectedItem is WindowInfo w) { RestoreAndActivate(w.Hwnd); if (_windowsPopup != null) _windowsPopup.IsOpen = false; } }
        #endregion

        #region Native window enumeration
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll")] private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
        [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll")] private static extern IntPtr GetShellWindow();
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);
        private const uint GW_OWNER = 4;
        private class TopWindow { public IntPtr Hwnd; public int ProcessId; public string Title; }
        private static List<TopWindow> EnumerateTopLevelWindows()
        { var list = new List<TopWindow>(); var shell = GetShellWindow(); EnumWindows((h, l) => { if (h == shell) return true; if (!IsWindowVisible(h)) return true; if (GetWindow(h, GW_OWNER) != IntPtr.Zero) return true; int pid; GetWindowThreadProcessId(h, out pid); var sb = new System.Text.StringBuilder(512); GetWindowText(h, sb, sb.Capacity); var title = sb.ToString(); if (string.IsNullOrWhiteSpace(title)) return true; list.Add(new TopWindow { Hwnd = h, ProcessId = pid, Title = title }); return true; }, IntPtr.Zero); return list; }
        #endregion

        #region Restore + Activate
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr hWnd);
        private const int SW_RESTORE = 9; private const int SW_SHOW = 5;
        private void RestoreAndActivate(IntPtr hwnd) { try { if (IsIconic(hwnd)) ShowWindow(hwnd, SW_RESTORE); else ShowWindow(hwnd, SW_SHOW); SetForegroundWindow(hwnd); } catch { } }
        #endregion

        #region Icon extraction helpers
        [DllImport("user32.dll", CharSet = CharSet.Auto)] private static extern IntPtr ExtractIcon(IntPtr hInst, string lpszExeFileName, int nIconIndex);
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImport("user32.dll")] private static extern IntPtr GetClassLongPtr(IntPtr hWnd, int nIndex);
        private const int WM_GETICON = 0x007F; private const int ICON_SMALL2 = 2; private const int GCL_HICON = -14;
        private BitmapSource TryGetProcessIcon(Process proc)
        { try { var path = proc?.MainModule?.FileName; if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return null; var h = ExtractIcon(IntPtr.Zero, path, 0); if (h == IntPtr.Zero) return null; var bmp = System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(h, Int32Rect.Empty, BitmapSizeOptions.FromWidthAndHeight(24, 24)); if (bmp.CanFreeze) bmp.Freeze(); return bmp; } catch { return null; } }
        private BitmapSource GetWindowIcon(IntPtr hwnd)
        {
            try
            {
                if (hwnd == IntPtr.Zero) return null;
                IntPtr hIcon = SendMessage(hwnd, WM_GETICON, ICON_SMALL2, 0);
                if (hIcon == IntPtr.Zero)
                {
                    hIcon = SafeGetClassLongIcon(hwnd, GCL_HICON);
                }
                if (hIcon == IntPtr.Zero) return null;
                var bmp = System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(hIcon, Int32Rect.Empty, BitmapSizeOptions.FromWidthAndHeight(16, 16));
                if (bmp.CanFreeze) bmp.Freeze();
                return bmp;
            }
            catch { return null; }
        }
        // Add SafeGetClassLongIcon helper
#if X64 || AMD64 || WIN64
        [DllImport("user32.dll", EntryPoint = "GetClassLongPtrW", SetLastError = true)]
        private static extern IntPtr GetClassLongPtr64(IntPtr hWnd, int nIndex);
#else
        [DllImport("user32.dll", EntryPoint = "GetClassLongW", SetLastError = true)]
        private static extern uint GetClassLong32(IntPtr hWnd, int nIndex);
#endif
        private static IntPtr SafeGetClassLongIcon(IntPtr hWnd, int index)
        {
            try
            {
#if X64 || AMD64 || WIN64
                return GetClassLongPtr64(hWnd, index);
#else
                uint val = GetClassLong32(hWnd, index);
                return new IntPtr(unchecked((int)val));
#endif
            }
            catch { return IntPtr.Zero; }
        }
        #endregion

        #region Volume Control
        [ComImport] [Guid("5CDF2C82-841E-4546-9722-0CF74078229A")] [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)] private interface IAudioEndpointVolume { int RegisterControlChangeNotify(IntPtr pNotify); int UnregisterControlChangeNotify(IntPtr pNotify); int GetChannelCount(out uint pnChannelCount); int SetMasterVolumeLevel(float fLevelDB, Guid pguidEventContext); int SetMasterVolumeLevelScalar(float fLevel, Guid pguidEventContext); int GetMasterVolumeLevel(out float pfLevelDB); int GetMasterVolumeLevelScalar(out float pfLevel); int SetChannelVolumeLevel(uint nChannel, float fLevelDB, Guid pguidEventContext); int SetChannelVolumeLevelScalar(uint nChannel, float fLevel, Guid pguidEventContext); int GetChannelVolumeLevel(uint nChannel, out float pfLevelDB); int GetChannelVolumeLevelScalar(uint nChannel, out float pfLevel); int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, Guid pguidEventContext); int GetMute(out bool pbMute); int GetVolumeStepInfo(out uint pnStep, out uint pnStepCount); int VolumeStepUp(Guid pguidEventContext); int VolumeStepDown(Guid pguidEventContext); int QueryHardwareSupport(out uint pdwHardwareSupportMask); int GetVolumeRange(out float pflVolumeMindB, out float pflVolumeMaxdB, out float pflVolumeIncrementdB); }
        [ComImport] [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")] private class MMDeviceEnumeratorComObject { }
        [ComImport] [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")] [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)] private interface IMMDeviceEnumerator { int NotImpl1(); int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice ppDevice); }
        [ComImport] [Guid("D666063F-1587-4E43-81F1-B948E807363F")] [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)] private interface IMMDevice { int Activate(ref Guid iid, int dwClsCtx, IntPtr pActivationParams, out IAudioEndpointVolume ppInterface); }
        private static IAudioEndpointVolume GetEndpointVolume() { var en = (IMMDeviceEnumerator)new MMDeviceEnumeratorComObject(); IMMDevice dev; en.GetDefaultAudioEndpoint(0, 1, out dev); var g = typeof(IAudioEndpointVolume).GUID; dev.Activate(ref g, 23, IntPtr.Zero, out var epv); return epv; }
        private static float GetMasterVolume() { try { var epv = GetEndpointVolume(); epv.GetMasterVolumeLevelScalar(out float v); Marshal.ReleaseComObject(epv); return v; } catch { return 0f; } }
        private static void SetMasterVolume(float value) { try { var epv = GetEndpointVolume(); epv.SetMasterVolumeLevelScalar(value, Guid.Empty); Marshal.ReleaseComObject(epv); } catch { } }
        private void VolumeSlider_ValueChanged(object s, RoutedPropertyChangedEventArgs<double> e) { if (!IsLoaded) return; SetMasterVolume((float)(VolumeSlider.Value / 100.0)); UpdateVolumeIcon((int)VolumeSlider.Value); }
        private void UpdateVolumeIcon(int v) { if (VolumeIcon == null) return; VolumeIcon.FontFamily = new FontFamily("Segoe MDL2 Assets"); if (v == 0) VolumeIcon.Text = "?"; else if (v < 30) VolumeIcon.Text = "?"; else if (v < 70) VolumeIcon.Text = "?"; else VolumeIcon.Text = "?"; }
        private void VolumeButton_Click(object s, RoutedEventArgs e) { VolumePopup.IsOpen = !VolumePopup.IsOpen; VolumePopup.PlacementTarget = VolumeButton; }
        #endregion

        #region Real Tray (Toolbar) Enumeration (unsupported / heuristic)
        // Attempt to read actual notification area toolbar(s) of explorer.exe.
        // This is unsupported and may break in future Windows versions.
        // We only extract icon handles from the imagelist (TB_GETIMAGELIST) per button index.

        // Win32
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string className, string windowText);
        [DllImport("user32.dll")] private static extern bool GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr OpenProcess(uint access, bool inherit, uint pid);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool CloseHandle(IntPtr hObject);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr VirtualAllocEx(IntPtr hProcess, IntPtr lpAddress, IntPtr dwSize, uint flAllocationType, uint flProtect);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool VirtualFreeEx(IntPtr hProcess, IntPtr lpAddress, IntPtr dwSize, uint dwFreeType);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool ReadProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, byte[] lpBuffer, IntPtr nSize, out IntPtr lpNumberOfBytesRead);
        [DllImport("user32.dll", CharSet = CharSet.Auto)] private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
        [DllImport("comctl32.dll")] private static extern IntPtr ImageList_GetIcon(IntPtr himl, int i, int flags);

        [DllImport("user32.dll")] private static extern bool DestroyIcon(IntPtr hIcon);

        private const int TB_BUTTONCOUNT = 0x0418;
        private const int TB_GETBUTTON = 0x0417;
        private const int TB_GETIMAGELIST = 0x0431;
        private const int ILD_NORMAL = 0x0000;

        private const uint PROCESS_VM_OPERATION = 0x0008;
        private const uint PROCESS_VM_READ = 0x0010;
        private const uint PROCESS_VM_WRITE = 0x0020;
        private const uint PROCESS_QUERY_INFORMATION = 0x0400;
        private const uint PROCESS_ALL = 0x1F0FFF; // fallback

        private const uint MEM_COMMIT = 0x1000;
        private const uint MEM_RELEASE = 0x8000;
        private const uint PAGE_READWRITE = 0x04;

        [StructLayout(LayoutKind.Sequential)] private struct TBBUTTON32
        { public int iBitmap; public int idCommand; public byte fsState; public byte fsStyle; public byte bReserved0; public byte bReserved1; public IntPtr dwData; public IntPtr iString; }
        [StructLayout(LayoutKind.Sequential)] private struct TBBUTTON64
        { public int iBitmap; public int idCommand; public byte fsState; public byte fsStyle; public byte bReserved0; public byte bReserved1; public IntPtr dwData; public IntPtr iString; }

        private void RefreshTrayIcons()
        {
            if (_trayAreaPanel == null) return; EnsureDefaultIcons();
            try
            {
                _trayAreaPanel.Children.Clear();
                var toolbars = LocateTrayToolbars();
                foreach (var tb in toolbars)
                {
                    foreach (var icon in EnumerateToolbarIcons(tb))
                    {
                        var img = new Image
                        {
                            Width = 16,
                            Height = 16,
                            Margin = new Thickness(2, 0, 2, 0),
                            Source = icon ?? _defaultWindowIcon16
                        };
                        _trayAreaPanel.Children.Add(img);
                        if (_trayAreaPanel.Children.Count > 30) break;
                    }
                    if (_trayAreaPanel.Children.Count > 30) break;
                }

                if (_trayAreaPanel.Children.Count == 0)
                {
                    _trayAreaPanel.Children.Add(new TextBlock
                    {
                        Text = "(no tray icons)",
                        Foreground = Brushes.Gray,
                        FontSize = 10,
                        Margin = new Thickness(4, 0, 4, 0),
                        VerticalAlignment = VerticalAlignment.Center
                    });
                }
            }
            catch (System.ComponentModel.Win32Exception wex)
            {
                Debug.WriteLine($"[Tray] Win32Exception {wex.NativeErrorCode}: {wex.Message}");
                if (_trayAreaPanel.Children.Count == 0)
                {
                    _trayAreaPanel.Children.Add(new TextBlock { Text = "(tray error)", Foreground = Brushes.OrangeRed, FontSize = 10, Margin = new Thickness(4,0,4,0) });
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Tray] General error: {ex.Message}");
                if (_trayAreaPanel.Children.Count == 0)
                    _trayAreaPanel.Children.Add(new TextBlock { Text = "(tray failure)", Foreground = Brushes.OrangeRed, FontSize = 10, Margin = new Thickness(4,0,4,0) });
            }
        }

        private IEnumerable<BitmapSource> EnumerateToolbarIcons(IntPtr toolbar)
        {
            var result = new List<BitmapSource>();
            try
            {
                int count = (int)SendMessage(toolbar, TB_BUTTONCOUNT, IntPtr.Zero, IntPtr.Zero);
                if (count <= 0 || count > 150) return result; // sanity
                IntPtr himl = SendMessage(toolbar, TB_GETIMAGELIST, IntPtr.Zero, IntPtr.Zero);
                if (himl == IntPtr.Zero) return result;

                GetWindowThreadProcessId(toolbar, out uint pid);
                var hProc = OpenProcess(PROCESS_VM_OPERATION | PROCESS_VM_READ | PROCESS_QUERY_INFORMATION, false, pid);
                if (hProc == IntPtr.Zero) return result; // skip if no access
                int btnSize = (IntPtr.Size == 8 ? Marshal.SizeOf(typeof(TBBUTTON64)) : Marshal.SizeOf(typeof(TBBUTTON32)));
                IntPtr remote = IntPtr.Zero;
                try
                {
                    remote = VirtualAllocEx(hProc, IntPtr.Zero, (IntPtr)btnSize, MEM_COMMIT, PAGE_READWRITE);
                    if (remote == IntPtr.Zero) return result;
                    var buffer = new byte[btnSize];
                    for (int i = 0; i < count; i++)
                    {
                        if (SendMessage(toolbar, TB_GETBUTTON, (IntPtr)i, remote) == IntPtr.Zero) continue;
                        if (!ReadProcessMemory(hProc, remote, buffer, (IntPtr)btnSize, out var read) || read.ToInt32() != btnSize) continue;
                        int iBitmap = BitConverter.ToInt32(buffer, 0);
                        if (iBitmap < 0 || iBitmap > 5000) continue;
                        var hIcon = ImageList_GetIcon(himl, iBitmap, ILD_NORMAL);
                        if (hIcon == IntPtr.Zero) continue;
                        try
                        {
                            var bmp = System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(hIcon, Int32Rect.Empty, BitmapSizeOptions.FromWidthAndHeight(16, 16));
                            if (bmp.CanFreeze) bmp.Freeze();
                            result.Add(bmp);
                        }
                        finally { DestroyIcon(hIcon); }
                        if (result.Count >= 30) break;
                    }
                }
                finally
                {
                    if (remote != IntPtr.Zero) VirtualFreeEx(hProc, remote, IntPtr.Zero, MEM_RELEASE);
                    CloseHandle(hProc);
                }
            }
            catch (System.ComponentModel.Win32Exception wex)
            {
                Debug.WriteLine($"[Tray] EnumerateToolbarIcons Win32Exception {wex.NativeErrorCode}: {wex.Message}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Tray] EnumerateToolbarIcons error: {ex.Message}");
            }
            return result;
        }

        private IEnumerable<IntPtr> LocateTrayToolbars()
        {
            var list = new List<IntPtr>();
            try
            {
                var tray = FindWindow("Shell_TrayWnd", null);
                if (tray != IntPtr.Zero)
                {
                    var notify = FindWindowEx(tray, IntPtr.Zero, "TrayNotifyWnd", null);
                    var sysPager = notify != IntPtr.Zero ? FindWindowEx(notify, IntPtr.Zero, "SysPager", null) : IntPtr.Zero;
                    var toolbar = sysPager != IntPtr.Zero ? FindWindowEx(sysPager, IntPtr.Zero, "ToolbarWindow32", null) : IntPtr.Zero;
                    if (toolbar != IntPtr.Zero) list.Add(toolbar);
                }
                var overflow = FindWindow("NotifyIconOverflowWindow", null);
                if (overflow != IntPtr.Zero)
                {
                    var ovTb = FindWindowEx(overflow, IntPtr.Zero, "ToolbarWindow32", null);
                    if (ovTb != IntPtr.Zero) list.Add(ovTb);
                }
            }
            catch { }
            return list;
        }

        #endregion // close Real Tray enumeration region

        // P/Invoke for broadcasting TaskbarCreated and basic window styles
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern uint RegisterWindowMessage(string lpString);
        [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
        private static readonly IntPtr HWND_BROADCAST = new IntPtr(0xFFFF);

        // Called after UI init
        private void InitShellTrayHost()
        {
            try
            {
                if (_msgTaskbarCreated == 0)
                    _msgTaskbarCreated = RegisterWindowMessage("TaskbarCreated");

                // Primary tray
                if (_shellTraySource == null)
                    _shellTraySource = CreateHiddenShellWindow("Shell_TrayWnd", ShellTrayWndProc);
                // Secondary tray host (even if only one monitor; some software enumerates it)
                if (_shellSecondaryTraySource == null)
                    _shellSecondaryTraySource = CreateHiddenShellWindow("Shell_SecondaryTrayWnd", ShellTrayWndProc);
                // Overflow host (apps may post messages to it)
                if (_notifyOverflowSource == null)
                    _notifyOverflowSource = CreateHiddenShellWindow("NotifyIconOverflowWindow", ShellTrayWndProc);

                // Broadcast so existing apps can re-register icons
                if (_msgTaskbarCreated != 0)
                    PostMessage(HWND_BROADCAST, _msgTaskbarCreated, IntPtr.Zero, IntPtr.Zero);

                Debug.WriteLine("[ShellHost] Created Shell_TrayWnd + Shell_SecondaryTrayWnd + NotifyIconOverflowWindow hosts.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[ShellHost] Init failed: " + ex.Message);
            }
        }

        private HwndSource CreateHiddenShellWindow(string className, HwndSourceHook hook)
        {
            var p = new HwndSourceParameters(className)
            {
                Width = 0,
                Height = 0,
                PositionX = -12000,
                PositionY = -12000,
                WindowStyle = unchecked((int)0x80000000) /*WS_POPUP*/ | 0x10000000 /*WS_VISIBLE*/,
                UsesPerPixelOpacity = false
            };
            var src = new HwndSource(p);
            if (hook != null) src.AddHook(hook);
            return src;
        }

        private IntPtr ShellTrayWndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            try
            {
                if (_msgTaskbarCreated != 0 && msg == _msgTaskbarCreated)
                {
                    Dispatcher.InvokeAsync(RefreshTrayIcons);
                }
            }
            catch { }
            return IntPtr.Zero;
        }

        // On window close, dispose host
        protected override void OnClosed(EventArgs e)
        {
            try { _shellTraySource?.RemoveHook(ShellTrayWndProc); _shellTraySource?.Dispose(); } catch { }
            try { _shellSecondaryTraySource?.RemoveHook(ShellTrayWndProc); _shellSecondaryTraySource?.Dispose(); } catch { }
            try { _notifyOverflowSource?.RemoveHook(ShellTrayWndProc); _notifyOverflowSource?.Dispose(); } catch { }
            base.OnClosed(e);
        }
    }
}
