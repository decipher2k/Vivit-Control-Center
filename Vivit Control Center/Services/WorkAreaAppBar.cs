using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace Vivit_Control_Center.Services
{
    // Disabled: legacy AppBar window approach superseded by AppBarService
    public static class WorkAreaAppBar
    {
        public static void Ensure(int widthPx, bool shellMode, int shellTaskbarHeightPx) { /* disabled */ }
        public static void EnsureForWindow(int widthPx, Window owner) { /* disabled */ }
        public static void Release() { /* disabled */ }
    }
}
