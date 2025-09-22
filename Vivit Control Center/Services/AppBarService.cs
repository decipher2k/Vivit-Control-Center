using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace Vivit_Control_Center.Services
{
    // Reserves desktop work area using real AppBars (Explorer-respected)
    public static class AppBarService
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct APPBARDATA
        {
            public int cbSize;                 // must be Int32 for correct struct size
            public IntPtr hWnd;                // HWND
            public uint uCallbackMessage;      // message id for appbar notifications
            public uint uEdge;                 // ABE_*
            public RECT rc;                    // bounding rectangle
            public IntPtr lParam;              // pointer-sized
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int left; public int top; public int right; public int bottom; }

        [StructLayout(LayoutKind.Sequential)]
        private struct MONITORINFO
        {
            public int cbSize;
            public RECT rcMonitor; // total monitor bounds (px)
            public RECT rcWork;    // monitor work area (px)
            public uint dwFlags;
        }

        #pragma warning disable 618
        [DllImport("shell32.dll", CallingConvention = CallingConvention.StdCall)] private static extern IntPtr SHAppBarMessage(uint dwMessage, ref APPBARDATA pData);
        #pragma warning restore 618
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern uint RegisterWindowMessage(string lpString);
        [DllImport("user32.dll")] private static extern int GetSystemMetrics(int nIndex);
        [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        [DllImport("user32.dll", SetLastError = true)] private static extern bool SystemParametersInfo(int uiAction, int uiParam, ref RECT pvParam, int fWinIni);
        [DllImport("user32.dll")] private static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);
        [DllImport("user32.dll")] private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);

        private const int SM_CXSCREEN = 0;
        private const int SM_CYSCREEN = 1;

        private const int SPI_GETWORKAREA = 0x0030;
        private const uint MONITOR_DEFAULTTONEAREST = 2;

        private const uint ABM_NEW = 0x00000000;
        private const uint ABM_REMOVE = 0x00000001;
        private const uint ABM_QUERYPOS = 0x00000002;
        private const uint ABM_SETPOS = 0x00000003;
        private const uint ABM_ACTIVATE = 0x00000006;
        private const uint ABM_WINDOWPOSCHANGED = 0x00000009;

        private const int ABE_LEFT = 0;
        private const int ABE_TOP = 1;

        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_SHOWWINDOW = 0x0040;

        private static readonly IntPtr HWND_BOTTOM = new IntPtr(1);

        // Extended styles
        private const int WS_EX_TOOLWINDOW = 0x00000080;
        private const int WS_EX_NOACTIVATE = 0x08000000;
        private const int WS_EX_TRANSPARENT = 0x00000020;

        // Left appbar host
        private static HwndSource _sourceLeft;
        private static IntPtr _hwndLeft = IntPtr.Zero;
        private static bool _registeredLeft;
        private static int _reservedLeftWidth;

        // Top appbar host
        private static HwndSource _sourceTop;
        private static IntPtr _hwndTop = IntPtr.Zero;
        private static bool _registeredTop;
        private static int _reservedTopHeight;

        private static uint _msgId;

        public static void EnsureOrUpdateLeft(int widthPx)
        {
            try
            {
                if (widthPx <= 0)
                {
                    RemoveLeft();
                    return;
                }

                if (_sourceLeft == null || _hwndLeft == IntPtr.Zero)
                {
                    _sourceLeft = CreateHostWindow("VivitAppBarHost.Left");
                    _hwndLeft = _sourceLeft.Handle;
                }

                if (!_registeredLeft)
                {
                    RegisterAppBar(_hwndLeft);
                    _registeredLeft = true;
                }

                if (widthPx != _reservedLeftWidth)
                {
                    _reservedLeftWidth = widthPx;
                    SetPositionLeft(widthPx);
                }
            }
            catch { }
        }

        public static void EnsureOrUpdateTop(int heightPx)
        {
            try
            {
                if (heightPx <= 0)
                {
                    RemoveTop();
                    return;
                }

                if (_sourceTop == null || _hwndTop == IntPtr.Zero)
                {
                    _sourceTop = CreateHostWindow("VivitAppBarHost.Top");
                    _hwndTop = _sourceTop.Handle;
                }

                if (!_registeredTop)
                {
                    RegisterAppBar(_hwndTop);
                    _registeredTop = true;
                }

                if (heightPx != _reservedTopHeight)
                {
                    _reservedTopHeight = heightPx;
                    SetPositionTop(heightPx);
                }
            }
            catch { }
        }

        public static void Remove()
        {
            RemoveLeft();
            RemoveTop();
        }

        private static void RemoveLeft()
        {
            try
            {
                if (_registeredLeft && _hwndLeft != IntPtr.Zero)
                {
                    var abd = new APPBARDATA { cbSize = Marshal.SizeOf(typeof(APPBARDATA)), hWnd = _hwndLeft };
                    SHAppBarMessage(ABM_REMOVE, ref abd);
                    _registeredLeft = false;
                }
            }
            catch { }
            finally
            {
                try { _sourceLeft?.Dispose(); } catch { }
                _sourceLeft = null;
                _hwndLeft = IntPtr.Zero;
                _reservedLeftWidth = 0;
            }
        }

        private static void RemoveTop()
        {
            try
            {
                if (_registeredTop && _hwndTop != IntPtr.Zero)
                {
                    var abd = new APPBARDATA { cbSize = Marshal.SizeOf(typeof(APPBARDATA)), hWnd = _hwndTop };
                    SHAppBarMessage(ABM_REMOVE, ref abd);
                    _registeredTop = false;
                }
            }
            catch { }
            finally
            {
                try { _sourceTop?.Dispose(); } catch { }
                _sourceTop = null;
                _hwndTop = IntPtr.Zero;
                _reservedTopHeight = 0;
            }
        }

        private static HwndSource CreateHostWindow(string name)
        {
            var p = new HwndSourceParameters(name)
            {
                Width = 1,
                Height = 1,
                PositionX = 0,
                PositionY = 0,
                WindowStyle = unchecked((int)0x80000000) /* WS_POPUP */,
                // Toolwindow, non-activating, click-through; not topmost
                ExtendedWindowStyle = WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE | WS_EX_TRANSPARENT
            };
            var src = new HwndSource(p);
            if (_msgId == 0) _msgId = RegisterWindowMessage("VIVIT_APPBAR_CALLBACK");
            src.AddHook(WndProc);
            return src;
        }

        private static void RegisterAppBar(IntPtr hwnd)
        {
            var abd = new APPBARDATA
            {
                cbSize = Marshal.SizeOf(typeof(APPBARDATA)),
                hWnd = hwnd,
                uCallbackMessage = _msgId
            };
            SHAppBarMessage(ABM_NEW, ref abd);
        }

        private static void GetMonitorBoundsForMainWindow(out RECT mon, out RECT work)
        {
            mon = new RECT { left = 0, top = 0, right = GetSystemMetrics(SM_CXSCREEN), bottom = GetSystemMetrics(SM_CYSCREEN) };
            work = mon;
            try
            {
                var main = Application.Current?.MainWindow;
                var hwndOwner = main != null ? new WindowInteropHelper(main).Handle : IntPtr.Zero;
                var hMon = MonitorFromWindow(hwndOwner != IntPtr.Zero ? hwndOwner : (_hwndLeft != IntPtr.Zero ? _hwndLeft : _hwndTop), MONITOR_DEFAULTTONEAREST);
                var mi = new MONITORINFO { cbSize = Marshal.SizeOf(typeof(MONITORINFO)) };
                if (GetMonitorInfo(hMon, ref mi))
                {
                    mon = mi.rcMonitor;
                    work = mi.rcWork;
                    return;
                }
            }
            catch { }

            // Fallback: SPI work area on primary
            var wa = new RECT();
            if (SystemParametersInfo(SPI_GETWORKAREA, 0, ref wa, 0))
            {
                work = wa;
            }
        }

        private static void SetPositionLeft(int widthPx)
        {
            if (_hwndLeft == IntPtr.Zero) return;

            GetMonitorBoundsForMainWindow(out var mon, out var work);

            var abd = new APPBARDATA
            {
                cbSize = Marshal.SizeOf(typeof(APPBARDATA)),
                hWnd = _hwndLeft,
                uEdge = ABE_LEFT,
                rc = new RECT { left = mon.left, top = work.top, right = mon.left + Math.Max(1, widthPx), bottom = work.bottom }
            };

            SHAppBarMessage(ABM_QUERYPOS, ref abd);
            abd.rc.right = abd.rc.left + Math.Max(1, widthPx);
            SHAppBarMessage(ABM_SETPOS, ref abd);

            SetWindowPos(_hwndLeft, HWND_BOTTOM, abd.rc.left, abd.rc.top, abd.rc.right - abd.rc.left, abd.rc.bottom - abd.rc.top, SWP_NOACTIVATE | SWP_SHOWWINDOW);

            var notify = new APPBARDATA { cbSize = Marshal.SizeOf(typeof(APPBARDATA)), hWnd = _hwndLeft };
            SHAppBarMessage(ABM_WINDOWPOSCHANGED, ref notify);
            SHAppBarMessage(ABM_ACTIVATE, ref abd);
        }

        private static void SetPositionTop(int heightPx)
        {
            if (_hwndTop == IntPtr.Zero) return;

            GetMonitorBoundsForMainWindow(out var mon, out var work);

            int h = Math.Max(1, heightPx);
            var abd = new APPBARDATA
            {
                cbSize = Marshal.SizeOf(typeof(APPBARDATA)),
                hWnd = _hwndTop,
                uEdge = ABE_TOP,
                // Anchor at monitor top so it sits below the title bar
                rc = new RECT { left = mon.left, top = mon.top, right = mon.right, bottom = mon.top + h }
            };

            SHAppBarMessage(ABM_QUERYPOS, ref abd);
            // Keep the requested height by adjusting bottom relative to whatever top Explorer decides
            abd.rc.bottom = abd.rc.top + h;
            SHAppBarMessage(ABM_SETPOS, ref abd);

            SetWindowPos(_hwndTop, HWND_BOTTOM, abd.rc.left, abd.rc.top, abd.rc.right - abd.rc.left, abd.rc.bottom - abd.rc.top, SWP_NOACTIVATE | SWP_SHOWWINDOW);

            var notify = new APPBARDATA { cbSize = Marshal.SizeOf(typeof(APPBARDATA)), hWnd = _hwndTop };
            SHAppBarMessage(ABM_WINDOWPOSCHANGED, ref notify);
            SHAppBarMessage(ABM_ACTIVATE, ref abd);
        }

        private static IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            // React to position changes requested by the shell
            if (_msgId != 0 && msg == _msgId)
            {
                // ABN_POSCHANGED = 0x00000001
                if (wParam.ToInt32() == 0x00000001)
                {
                    if (_registeredLeft) SetPositionLeft(Math.Max(1, _reservedLeftWidth));
                    if (_registeredTop) SetPositionTop(Math.Max(1, _reservedTopHeight));
                }
            }
            return IntPtr.Zero;
        }
    }
}
