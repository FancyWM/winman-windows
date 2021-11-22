using System;
using System.Collections.Generic;

namespace WinMan.Windows
{
    internal interface IWin32VirtualDesktopService
    {
        int GetCurrentDesktopIndex(IntPtr hMon);

        int GetDesktopCount(IntPtr hMon);

        bool IsWindowPinned(IntPtr hWnd);

        bool IsCurrentDesktop(IntPtr hMon, object desktop);

        string GetDesktopName(object desktop);

        void SwitchToDesktop(IntPtr hMon, object desktop);

        void MoveToDesktop(IntPtr hWnd, object desktop);

        bool IsWindowOnCurrentDesktop(IntPtr hWnd);

        bool HasWindow(object desktop, IntPtr hWnd);
        
        object GetDesktopByIndex(IntPtr hMon, int index);

        List<object> GetVirtualDesktops(IntPtr hMon);

        int GetDesktopIndex(IntPtr hMon, object m_desktop);
    }
}
