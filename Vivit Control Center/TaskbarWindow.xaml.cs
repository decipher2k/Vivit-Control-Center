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
using System.Text; // (für StringBuilder in neuen Methoden)

namespace Vivit_Control_Center
{   
    public partial class TaskbarWindow : Window
    {
        // Win32 RECT + GetWindowRect (single declaration here)
        [StructLayout(LayoutKind.Sequential)] private struct RECT { public int left; public int top; public int right; public int bottom; }
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

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

        private uint _msgTaskbarCreated; // jetzt verwendet
        private DispatcherTimer _taskbarHideWatchdog; // neu: sorgt dafür, dass Taskleiste nicht zurückkommt

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
                var bg = new SolidColorBrush(Color.FromRgb(0, 0, 0));
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
            try { (Application.Current?.MainWindow as MainWindow)?.ActivateModule("Programs"); } catch { }
        }

        private void TaskbarWindow_Loaded(object sender, RoutedEventArgs e)
        {
            EnsureDefaultIcons();
            PositionAtBottom();

            StartTrayDiscoveryAsync();

            _clockTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(10) };
            _clockTimer.Tick += (_, __) => UpdateClock(); _clockTimer.Start(); UpdateClock();

            _processRefreshTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(3) };
            _processRefreshTimer.Tick += (_, __) => RefreshProcessIcons(); _processRefreshTimer.Start(); RefreshProcessIcons();

            _trayRefreshTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(15) };
            _trayRefreshTimer.Tick += (_, __) => RefreshTrayIcons(); _trayRefreshTimer.Start();

            try { var vol = (int)(GetMasterVolume() * 100); VolumeSlider.Value = vol; UpdateVolumeIcon(vol); } catch { }

            // Shell-Mode: Explorer-Taskbar ausblenden (falls aktiv)
            if (IsShellMode())
            {
                HideExplorerTaskbars(); // einmalig (OnSourceInitialized greift ebenfalls)
                StartTaskbarWatchdog();
            }
        }

        private async void StartTrayDiscoveryAsync()
        {
            for (int i = 0; i < 25; i++)
            {
                RefreshTrayIcons();
                if (_trayAreaPanel != null && _trayAreaPanel.Children.OfType<Image>().Any())
                {
                    Debug.WriteLine("[Tray] Icons gefunden nach Versuch " + (i + 1));
                    return;
                }
                await Task.Delay(800);
            }
            Debug.WriteLine("[Tray] Keine Icons nach Timeout – evtl. Windows 11 Struktur oder Rechteproblem.");
        }

        private void PositionAtBottom()
        {
            Left = 0;
            Width = SystemParameters.PrimaryScreenWidth;
            var src = PresentationSource.FromVisual(this);
            double dipHeight = App.ShellTaskbarHeightPx * (src?.CompositionTarget?.TransformFromDevice.M22 ?? 1.0);
            if (double.IsNaN(dipHeight) || dipHeight <= 0) dipHeight = 40; // fallback
            Height = Math.Max(24, dipHeight);
            Top = SystemParameters.PrimaryScreenHeight - Height;
        }
        private void UpdateClock() { var now = DateTime.Now; TimeText.Text = now.ToString("HH:mm"); DateText.Text = now.ToString("dd.MM.yyyy"); }

        #region Process + Window enumeration
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
            var btn = new Button { Width = 34, Height = 34, Padding = new Thickness(0), Margin = new Thickness(2, 3, 2, 3), Background = Brushes.Transparent, BorderBrush = Brushes.Transparent, ToolTip = pinfo.DisplayName, Tag = pinfo };
            btn.MouseEnter += ProcessButton_MouseEnter; btn.Click += ProcessButton_Click;
            var iconSource = TryGetProcessIcon(pinfo.Process) ?? GetWindowIcon(pinfo.FirstWindow) ?? _defaultProcessIcon24;
            btn.Content = new Image { Source = iconSource, Width = 24, Height = 24 };
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
        [DllImport("user32.dll")][return: MarshalAs(UnmanagedType.Bool)] private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll")] private static extern IntPtr GetShellWindow();
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMax);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        private const uint GW_OWNER = 4;
        private class TopWindow { public IntPtr Hwnd; public int ProcessId; public string Title; }
        private static List<TopWindow> EnumerateTopLevelWindows()
        {
            var list = new List<TopWindow>(); var shell = GetShellWindow();
            EnumWindows((h, l) =>
            {
                if (h == shell) return true;
                if (!IsWindowVisible(h)) return true;
                if (GetWindow(h, GW_OWNER) != IntPtr.Zero) return true;
                int pid; GetWindowThreadProcessId(h, out pid);
                var sb = new System.Text.StringBuilder(512);
                GetWindowText(h, sb, sb.Capacity);
                var title = sb.ToString();
                if (string.IsNullOrWhiteSpace(title)) return true;
                list.Add(new TopWindow { Hwnd = h, ProcessId = pid, Title = title });
                return true;
            }, IntPtr.Zero);
            return list;
        }

        private IntPtr GetExplorerShellTrayWnd()
        {
            try
            {
                int currentPid = Process.GetCurrentProcess().Id;
                IntPtr found = IntPtr.Zero;
                EnumWindows((h, l) =>
                {
                    var cls = new System.Text.StringBuilder(64);
                    if (GetClassName(h, cls, cls.Capacity) > 0 && cls.ToString() == "Shell_TrayWnd")
                    {
                        int pid; GetWindowThreadProcessId(h, out pid);
                        if (pid != currentPid) { found = h; return false; }
                    }
                    return true;
                }, IntPtr.Zero);
                return found;
            }
            catch { return IntPtr.Zero; }
        }
        #endregion

        #region Restore + Activate
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr hWnd);
        private const int SW_RESTORE = 9; private const int SW_SHOW = 5;
        private void RestoreAndActivate(IntPtr hwnd)
        {
            try
            {
                if (IsIconic(hwnd)) ShowWindow(hwnd, SW_RESTORE); else ShowWindow(hwnd, SW_SHOW);
                SetForegroundWindow(hwnd);
            }
            catch { }
        }
        #endregion

        #region Icon helpers
        [DllImport("user32.dll", CharSet = CharSet.Auto)] private static extern IntPtr ExtractIcon(IntPtr hInst, string file, int index);
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImport("user32.dll")] private static extern IntPtr GetClassLongPtr(IntPtr hWnd, int nIndex);
        private const int WM_GETICON = 0x007F; private const int ICON_SMALL2 = 2; private const int GCL_HICON = -14;
        private BitmapSource TryGetProcessIcon(Process proc)
        {
            try
            {
                var path = proc?.MainModule?.FileName;
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return null;
                var h = ExtractIcon(IntPtr.Zero, path, 0);
                if (h == IntPtr.Zero) return null;
                var bmp = Imaging.CreateBitmapSourceFromHIcon(h, Int32Rect.Empty, BitmapSizeOptions.FromWidthAndHeight(24, 24));
                if (bmp.CanFreeze) bmp.Freeze(); return bmp;
            }
            catch { return null; }
        }
        private BitmapSource GetWindowIcon(IntPtr hwnd)
        {
            try
            {
                if (hwnd == IntPtr.Zero) return null;
                IntPtr hIcon = SendMessage(hwnd, WM_GETICON, ICON_SMALL2, 0);
                if (hIcon == IntPtr.Zero) hIcon = SafeGetClassLongIcon(hwnd, GCL_HICON);
                if (hIcon == IntPtr.Zero) return null;
                var bmp = Imaging.CreateBitmapSourceFromHIcon(hIcon, Int32Rect.Empty, BitmapSizeOptions.FromWidthAndHeight(16, 16));
                if (bmp.CanFreeze) bmp.Freeze();
                return bmp;
            }
            catch { return null; }
        }
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

        #region Volume
        [ComImport]
        [Guid("5CDF2C82-841E-4546-9722-0CF74078229A")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IAudioEndpointVolume
        {
            int RegisterControlChangeNotify(IntPtr pNotify); int UnregisterControlChangeNotify(IntPtr pNotify);
            int GetChannelCount(out uint pnChannelCount); int SetMasterVolumeLevel(float fLevelDB, Guid pguidEventContext);
            int SetMasterVolumeLevelScalar(float fLevel, Guid pguidEventContext); int GetMasterVolumeLevel(out float pfLevelDB);
            int GetMasterVolumeLevelScalar(out float pfLevel); int SetChannelVolumeLevel(uint nChannel, float fLevelDB, Guid pguidEventContext);
            int SetChannelVolumeLevelScalar(uint nChannel, float fLevel, Guid pguidEventContext); int GetChannelVolumeLevel(uint nChannel, out float pfLevelDB);
            int GetChannelVolumeLevelScalar(uint nChannel, out float pfLevel); int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, Guid pguidEventContext);
            int GetMute(out bool pbMute); int GetVolumeStepInfo(out uint pnStep, out uint pnStepCount); int VolumeStepUp(Guid pguidEventContext);
            int VolumeStepDown(Guid pguidEventContext); int QueryHardwareSupport(out uint mask); int GetVolumeRange(out float mindB, out float maxdB, out float incrdB);
        }
        [ComImport][Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")] private class MMDeviceEnumeratorComObject { }
        [ComImport][Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")][InterfaceType(ComInterfaceType.InterfaceIsIUnknown)] private interface IMMDeviceEnumerator { int NotImpl1(); int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice ppDevice); }
        [ComImport][Guid("D666063F-1587-4E43-81F1-B948E807363F")][InterfaceType(ComInterfaceType.InterfaceIsIUnknown)] private interface IMMDevice { int Activate(ref Guid iid, int dwClsCtx, IntPtr pAct, out IAudioEndpointVolume ep); }
        private static IAudioEndpointVolume GetEndpointVolume() { var en = (IMMDeviceEnumerator)new MMDeviceEnumeratorComObject(); IMMDevice dev; en.GetDefaultAudioEndpoint(0, 1, out dev); var g = typeof(IAudioEndpointVolume).GUID; dev.Activate(ref g, 23, IntPtr.Zero, out var epv); return epv; }
        private static float GetMasterVolume() { try { var epv = GetEndpointVolume(); epv.GetMasterVolumeLevelScalar(out float v); Marshal.ReleaseComObject(epv); return v; } catch { return 0f; } }
        private static void SetMasterVolume(float value) { try { var epv = GetEndpointVolume(); epv.SetMasterVolumeLevelScalar(value, Guid.Empty); Marshal.ReleaseComObject(epv); } catch { } }
        private void VolumeSlider_ValueChanged(object s, RoutedPropertyChangedEventArgs<double> e) { if (!IsLoaded) return; SetMasterVolume((float)(VolumeSlider.Value / 100.0)); UpdateVolumeIcon((int)VolumeSlider.Value); }
        private void UpdateVolumeIcon(int v) { if (VolumeIcon == null) return; VolumeIcon.FontFamily = new FontFamily("Segoe MDL2 Assets"); VolumeIcon.Text = v == 0 ? "\uE198" : v < 30 ? "\uE15D" : v < 70 ? "\uE995" : "\uE15E"; }
        private void VolumeButton_Click(object s, RoutedEventArgs e) { VolumePopup.IsOpen = !VolumePopup.IsOpen; VolumePopup.PlacementTarget = VolumeButton; }
        #endregion

        #region Tray enumeration (improved)
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr FindWindow(string cls, string win);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string cls, string txt);
        [DllImport("user32.dll")] private static extern bool GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr OpenProcess(uint access, bool inherit, uint pid);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool CloseHandle(IntPtr hObject);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr VirtualAllocEx(IntPtr hProcess, IntPtr addr, IntPtr size, uint type, uint protect);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool VirtualFreeEx(IntPtr hProcess, IntPtr addr, IntPtr size, uint freeType);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool ReadProcessMemory(IntPtr hProcess, IntPtr baseAddr, byte[] buffer, IntPtr size, out IntPtr read);
        [DllImport("user32.dll", CharSet = CharSet.Auto)] private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
        [DllImport("comctl32.dll")] private static extern IntPtr ImageList_GetIcon(IntPtr himl, int i, int flags);
        [DllImport("user32.dll")] private static extern bool DestroyIcon(IntPtr hIcon);

        private const int TB_BUTTONCOUNT = 0x0418;
        private const int TB_GETBUTTON = 0x0417;
        private const int TB_GETIMAGELIST = 0x0431;
        private const int ILD_NORMAL = 0x0000;
        private const uint PROCESS_VM_OPERATION = 0x0008;
        private const uint PROCESS_VM_READ = 0x0010;
        private const uint PROCESS_QUERY_INFORMATION = 0x0400;
        private const uint MEM_COMMIT = 0x1000;
        private const uint MEM_RELEASE = 0x8000;
        private const uint PAGE_READWRITE = 0x04;

        [StructLayout(LayoutKind.Sequential)]
        private struct TBBUTTON32
        { public int iBitmap; public int idCommand; public byte fsState; public byte fsStyle; public byte bReserved0; public byte bReserved1; public IntPtr dwData; public IntPtr iString; }
        [StructLayout(LayoutKind.Sequential)]
        private struct TBBUTTON64
        { public int iBitmap; public int idCommand; public byte fsState; public byte fsStyle; public byte bReserved0; public byte bReserved1; public IntPtr dwData; public IntPtr iString; }

        private void RefreshTrayIcons()
        {
            if (_trayAreaPanel == null) return; EnsureDefaultIcons();
            try
            {
                _trayAreaPanel.Children.Clear();
                var toolbars = LocateTrayToolbars();
                Debug.WriteLine("[Tray] Gefundene Toolbars: " + string.Join(", ", toolbars.Select(h => h.ToString("X"))));
                foreach (var tb in toolbars)
                {
                    foreach (var icon in EnumerateToolbarIcons(tb))
                    {
                        _trayAreaPanel.Children.Add(new Image
                        {
                            Width = 16,
                            Height = 16,
                            Margin = new Thickness(2, 0, 2, 0),
                            Source = icon ?? _defaultWindowIcon16
                        });
                        if (_trayAreaPanel.Children.Count >= 40) break;
                    }
                    if (_trayAreaPanel.Children.Count >= 40) break;
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
            catch (Exception ex)
            {
                Debug.WriteLine("[Tray] Fehler: " + ex.Message);
                if (_trayAreaPanel.Children.Count == 0)
                    _trayAreaPanel.Children.Add(new TextBlock { Text = "(tray failure)", Foreground = Brushes.OrangeRed, FontSize = 10, Margin = new Thickness(4, 0, 4, 0) });
            }
        }

        private IEnumerable<BitmapSource> EnumerateToolbarIcons(IntPtr toolbar)
        {
            var result = new List<BitmapSource>();
            try
            {
                int count = (int)SendMessage(toolbar, TB_BUTTONCOUNT, IntPtr.Zero, IntPtr.Zero);
                if (count <= 0 || count > 300) return result;
                IntPtr himl = SendMessage(toolbar, TB_GETIMAGELIST, IntPtr.Zero, IntPtr.Zero);
                if (himl == IntPtr.Zero) return result;

                GetWindowThreadProcessId(toolbar, out uint pid);
                var hProc = OpenProcess(PROCESS_VM_OPERATION | PROCESS_VM_READ | PROCESS_QUERY_INFORMATION, false, pid);
                if (hProc == IntPtr.Zero) return result;
                int btnSize = (IntPtr.Size == 8 ? Marshal.SizeOf(typeof(TBBUTTON64)) : Marshal.SizeOf(typeof(TBBUTTON32)));
                IntPtr remote = IntPtr.Zero;
                var buffer = new byte[btnSize];
                try
                {
                    remote = VirtualAllocEx(hProc, IntPtr.Zero, (IntPtr)btnSize, MEM_COMMIT, PAGE_READWRITE);
                    if (remote == IntPtr.Zero) return result;
                    for (int i = 0; i < count; i++)
                    {
                        if (SendMessage(toolbar, TB_GETBUTTON, (IntPtr)i, remote) == IntPtr.Zero) continue;
                        IntPtr read;
                        if (!ReadProcessMemory(hProc, remote, buffer, (IntPtr)btnSize, out read) || read.ToInt32() != btnSize) continue;
                        int iBitmap = BitConverter.ToInt32(buffer, 0);
                        if (iBitmap < 0 || iBitmap > 8000) continue;
                        var hIcon = ImageList_GetIcon(himl, iBitmap, ILD_NORMAL);
                        if (hIcon == IntPtr.Zero) continue;
                        try
                        {
                            var bmp = Imaging.CreateBitmapSourceFromHIcon(hIcon, Int32Rect.Empty, BitmapSizeOptions.FromWidthAndHeight(16, 16));
                            if (bmp.CanFreeze) bmp.Freeze();
                            result.Add(bmp);
                        }
                        finally { DestroyIcon(hIcon); }
                        if (result.Count >= 50) break;
                    }
                }
                finally
                {
                    if (remote != IntPtr.Zero) VirtualFreeEx(hProc, remote, IntPtr.Zero, MEM_RELEASE);
                    CloseHandle(hProc);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[Tray] EnumerateToolbarIcons: " + ex.Message);
            }
            return result;
        }

        private IEnumerable<IntPtr> LocateTrayToolbars()
        {
            var list = new List<IntPtr>();
            try
            {
                var tray = GetExplorerShellTrayWnd();
                if (tray == IntPtr.Zero) return list;

                var notify = FindWindowEx(tray, IntPtr.Zero, "TrayNotifyWnd", null);
                if (notify == IntPtr.Zero) return list;

                var sysPager = FindWindowEx(notify, IntPtr.Zero, "SysPager", null);
                if (sysPager != IntPtr.Zero)
                {
                    var tb = FindWindowEx(sysPager, IntPtr.Zero, "ToolbarWindow32", null);
                    if (tb != IntPtr.Zero) list.Add(tb);
                }

                EnumChildWindows(notify, (h, l) =>
                {
                    var cls = new System.Text.StringBuilder(64);
                    if (GetClassName(h, cls, cls.Capacity) > 0 && cls.ToString() == "ToolbarWindow32")
                    {
                        if (!list.Contains(h)) list.Add(h);
                    }
                    return true;
                }, IntPtr.Zero);

                var overflow = FindWindow("NotifyIconOverflowWindow", null);
                if (overflow != IntPtr.Zero)
                {
                    EnumChildWindows(overflow, (h, l) =>
                    {
                        var cls = new System.Text.StringBuilder(64);
                        if (GetClassName(h, cls, cls.Capacity) > 0 && cls.ToString() == "ToolbarWindow32")
                        {
                            if (!list.Contains(h)) list.Add(h);
                        }
                        return true;
                    }, IntPtr.Zero);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[Tray] LocateTrayToolbars Fehler: " + ex.Message);
            }
            return list;
        }
        #endregion

        // === Explorer Tray Toggle Support ===
        [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        [DllImport("gdi32.dll")] private static extern IntPtr CreateRectRgn(int left, int top, int right, int bottom);
        [DllImport("user32.dll")] private static extern int SetWindowRgn(IntPtr hWnd, IntPtr hRgn, bool bRedraw);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern uint RegisterWindowMessage(string lpString); // NEU
        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const int SW_SHOWNA = 8;
        private const int SW_HIDE = 0; // NEU
        private IntPtr _cachedExplorerTray = IntPtr.Zero;
        private bool _explorerTrayVisible = false;
        private void ShowExplorerTrayButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_cachedExplorerTray == IntPtr.Zero || !IsWindowVisible(_cachedExplorerTray))
                {
                    _cachedExplorerTray = GetExplorerShellTrayWnd();
                }
                if (_cachedExplorerTray == IntPtr.Zero) return;

                IntPtr trayNotify = FindWindowEx(_cachedExplorerTray, IntPtr.Zero, "TrayNotifyWnd", null);
                if (trayNotify == IntPtr.Zero)
                {
                    if (!_explorerTrayVisible)
                    {
                        var src = PresentationSource.FromVisual(this);
                        int hPx = Math.Max(24, App.ShellTaskbarHeightPx);
                        int yPx = (int)Math.Round(SystemParameters.PrimaryScreenHeight / (src?.CompositionTarget?.TransformFromDevice.M22 ?? 1.0) - hPx);
                        SetWindowPos(_cachedExplorerTray, IntPtr.Zero, 0, yPx, (int)(SystemParameters.PrimaryScreenWidth / (src?.CompositionTarget?.TransformFromDevice.M11 ?? 1.0)), hPx, SWP_NOZORDER | SWP_NOACTIVATE);
                        ShowWindow(_cachedExplorerTray, SW_SHOWNA);
                        _explorerTrayVisible = true;
                    }
                    else
                    {
                        SetWindowPos(_cachedExplorerTray, IntPtr.Zero, -5000, (int)SystemParameters.PrimaryScreenHeight - 2, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_NOACTIVATE);
                        _explorerTrayVisible = false;
                    }
                    return;
                }

                IntPtr iconToolbar = IntPtr.Zero;
                IntPtr sysPager = FindWindowEx(trayNotify, IntPtr.Zero, "SysPager", null);
                if (sysPager != IntPtr.Zero)
                {
                    iconToolbar = FindWindowEx(sysPager, IntPtr.Zero, "ToolbarWindow32", null);
                }
                if (iconToolbar == IntPtr.Zero)
                {
                    iconToolbar = FindWindowEx(trayNotify, IntPtr.Zero, "ToolbarWindow32", null);
                }

                RECT shellRc; GetWindowRect(_cachedExplorerTray, out shellRc);
                RECT trayRc; GetWindowRect(trayNotify, out trayRc);

                RECT targetRc = trayRc;
                if (iconToolbar != IntPtr.Zero)
                {
                    RECT tbRc; GetWindowRect(iconToolbar, out tbRc);
                    targetRc = tbRc;
                }

                int width = targetRc.right - targetRc.left;
                int height = targetRc.bottom - targetRc.top;
                if (height <= 0) height = Math.Max(24, App.ShellTaskbarHeightPx);
                if (width <= 0 || width > SystemParameters.PrimaryScreenWidth) width = 260;

                if (!_explorerTrayVisible)
                {
                    int relLeft = targetRc.left - shellRc.left;
                    int relTop = targetRc.top - shellRc.top;
                    int relRight = relLeft + width;
                    int relBottom = relTop + height;
                    if (relLeft < 0 || relTop < 0)
                    {
                        relLeft = 0;
                        relTop = shellRc.bottom - shellRc.top - height;
                        relRight = relLeft + width;
                        relBottom = relTop + height;
                    }

                    IntPtr rgn = CreateRectRgn(relLeft, relTop, relRight, relBottom);
                    if (rgn != IntPtr.Zero)
                    {
                        SetWindowRgn(_cachedExplorerTray, rgn, true);
                    }

                    int screenW = (int)(SystemParameters.PrimaryScreenWidth / (PresentationSource.FromVisual(this)?.CompositionTarget?.TransformFromDevice.M11 ?? 1.0));
                    int screenH = (int)(SystemParameters.PrimaryScreenHeight / (PresentationSource.FromVisual(this)?.CompositionTarget?.TransformFromDevice.M22 ?? 1.0));
                    int newX = screenW - width;
                    int newY = screenH - height;
                    SetWindowPos(_cachedExplorerTray, IntPtr.Zero, newX, newY, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
                    ShowWindow(_cachedExplorerTray, SW_SHOWNA);
                    _explorerTrayVisible = true;
                }
                else
                {
                    SetWindowRgn(_cachedExplorerTray, IntPtr.Zero, true);
                    SetWindowPos(_cachedExplorerTray, IntPtr.Zero, -5000, (int)SystemParameters.PrimaryScreenHeight - 2, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_NOACTIVATE);
                    _explorerTrayVisible = false;
                }
            }
            catch { }
        }
        // === End Explorer Tray Toggle Support ===

        // === NEU: Sicheres Ausblenden der Windows-Taskbar im Shell-Modus ===
        private bool IsShellMode()
        {
            return App.IsShellMode;
        }

        private void StartTaskbarWatchdog()
        {
            if (_taskbarHideWatchdog != null) return;
            _taskbarHideWatchdog = new DispatcherTimer { Interval = TimeSpan.FromSeconds(5) };
            _taskbarHideWatchdog.Tick += (s, e) => { HideExplorerTaskbars(); };
            _taskbarHideWatchdog.Start();
        }

        private void HideExplorerTaskbars()
        {
            try
            {
                IntPtr primary = FindWindow("Shell_TrayWnd", null);
                if (primary != IntPtr.Zero)
                {
                    ShowWindow(primary, SW_HIDE);
                }

                // Sekundäre Taskleisten (Multi-Monitor)
                EnumWindows((h, l) =>
                {
                    var sb = new StringBuilder(64);
                    if (GetClassName(h, sb, sb.Capacity) > 0)
                    {
                        var cls = sb.ToString();
                        if (cls == "Shell_SecondaryTrayWnd")
                        {
                            ShowWindow(h, SW_HIDE);
                        }
                    }
                    return true;
                }, IntPtr.Zero);

                // Optional: Startmenü-Fenster (Windows 10/11) verbergen
                IntPtr start = FindWindow("Windows.UI.Core.CoreWindow", "Start");
                if (start != IntPtr.Zero) ShowWindow(start, SW_HIDE);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[ShellHide] Fehler: " + ex.Message);
            }
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            try
            {
                var src = (HwndSource)PresentationSource.FromVisual(this);
                if (src != null)
                {
                    _msgTaskbarCreated = RegisterWindowMessage("TaskbarCreated");
                    src.AddHook(WndProc);
                }
                if (IsShellMode())
                {
                    HideExplorerTaskbars();
                }
            }
            catch { }
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (_msgTaskbarCreated != 0 && (uint)msg == _msgTaskbarCreated)
            {
                // Explorer wurde neu gestartet -> erneut verbergen
                Dispatcher.BeginInvoke((Action)HideExplorerTaskbars);
            }
            return IntPtr.Zero;
        }
        // === Ende Neuer Code ===

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
        }

        private void AppButtonLogout_Click(object sender, RoutedEventArgs e)
        {
            try { Process.Start("shutdown", "/l"); } catch { }
        }
        private void AppButtonRestart_Click(object sender, RoutedEventArgs e)
        {
            try { Process.Start("shutdown", "/r /t 0"); } catch { }
        }
        private void AppButtonShutdown_Click(object sender, RoutedEventArgs e)
        {
            try { Process.Start("shutdown", "/s /t 0"); } catch { }
        }
    }
}
