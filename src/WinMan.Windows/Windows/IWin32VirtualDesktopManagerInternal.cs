using System;


namespace WinMan.Windows
{
    internal interface IWin32VirtualDesktopManagerInternal
    {
        void CheckVirtualDesktopChanges();

        bool IsNotOnCurrentDesktop(IntPtr hwnd);
    }
}
