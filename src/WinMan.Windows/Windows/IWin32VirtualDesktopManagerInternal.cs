using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace WinMan.Windows.Windows
{
    internal interface IWin32VirtualDesktopManagerInternal
    {
        void CheckVirtualDesktopChanges();

        bool IsNotOnCurrentDesktop(IntPtr hwnd);
    }
}
