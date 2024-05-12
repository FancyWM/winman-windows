using System;
using System.Collections.Generic;

namespace WinMan.Windows
{
    internal interface IWin32VirtualDesktopService
    {
        class Desktop
        {
            public IntPtr Monitor { get; init; }
            public Guid Guid { get; init; }

            public Desktop(Guid guid)
            {
                Guid = guid;
            }

            public Desktop(IntPtr monitor, Guid guid)
            {
                Monitor = monitor;
                Guid = guid;
            }
        }

        void Connect();

        int GetCurrentDesktopIndex(IntPtr hMon);

        int GetDesktopCount(IntPtr hMon);

        bool IsWindowPinned(IntPtr hWnd);

        bool IsCurrentDesktop(IntPtr hMon, Desktop desktop);

        string GetDesktopName(Desktop desktop);

        void SwitchToDesktop(IntPtr hMon, Desktop desktop);

        void MoveToDesktop(IntPtr hWnd, Desktop desktop);

        bool IsWindowOnCurrentDesktop(IntPtr hWnd);

        bool HasWindow(Desktop desktop, IntPtr hWnd);

        Desktop GetDesktopByIndex(IntPtr hMon, int index);

        List<Desktop> GetVirtualDesktops(IntPtr hMon);

        int GetDesktopIndex(IntPtr hMon, Desktop m_desktop);
    }
}
