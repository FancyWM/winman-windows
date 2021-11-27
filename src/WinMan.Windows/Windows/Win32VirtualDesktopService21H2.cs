using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

using WinMan.Windows.Com;
using WinMan.Windows.Com.Build21H2;

using WinMan.Windows.DllImports;
using static WinMan.Windows.DllImports.NativeMethods;

namespace WinMan.Windows
{
    internal class Win32VirtualDesktopService21H2 : IWin32VirtualDesktopService
    {
        private readonly IComVirtualDesktopManagerInternal VirtualDesktopManagerInternal;
        private readonly IComVirtualDesktopManager VirtualDesktopManager;
        private readonly IComApplicationViewCollection ApplicationViewCollection;
        private readonly IComVirtualDesktopPinnedApps VirtualDesktopPinnedApps;

        public Win32VirtualDesktopService21H2()
        {
            var shell = (IComServiceProvider10)Activator.CreateInstance(Type.GetTypeFromCLSID(ComGuids.CLSID_ImmersiveShell));
            VirtualDesktopManagerInternal = (IComVirtualDesktopManagerInternal)shell.QueryService(ComGuids.CLSID_VirtualDesktopManagerInternal, typeof(IComVirtualDesktopManagerInternal).GUID);
            VirtualDesktopManager = (IComVirtualDesktopManager)Activator.CreateInstance(Type.GetTypeFromCLSID(ComGuids.CLSID_VirtualDesktopManager));
            ApplicationViewCollection = (IComApplicationViewCollection)shell.QueryService(typeof(IComApplicationViewCollection).GUID, typeof(IComApplicationViewCollection).GUID);
            VirtualDesktopPinnedApps = (IComVirtualDesktopPinnedApps)shell.QueryService(ComGuids.CLSID_VirtualDesktopPinnedApps, typeof(IComVirtualDesktopPinnedApps).GUID);
        }

        public int GetCurrentDesktopIndex(IntPtr hMon)
        {
            return GetDesktopIndex(hMon, VirtualDesktopManagerInternal.GetCurrentDesktop(hMon));
        }

        public int GetDesktopCount(IntPtr hMon)
        {
            return VirtualDesktopManagerInternal.GetCount(hMon);
        }

        public string GetDesktopName(object desktop)
        {
            IComVirtualDesktop comDesktop = (IComVirtualDesktop)desktop;
            // get desktop name
            string desktopName = null;
            try
            {
                desktopName = MarshalHSTRING(new(comDesktop.GetName()));
            }
            catch { }

            // no name found, generate generic name
            if (string.IsNullOrEmpty(desktopName))
            { // create name "Desktop n" (n = number starting with 1)
                desktopName = "Desktop " + (GetDesktopIndex(IntPtr.Zero, comDesktop) + 1).ToString();
            }
            return desktopName;
        }

        private string MarshalHSTRING(HSTRING hStr)
        {
            unsafe
            {
                uint length = 0;
                PCWSTR pBuffer = WindowsGetStringRawBuffer(hStr, &length);
                string str = new((char*)pBuffer, 0, (int)length);
                NativeMethods.WindowsDeleteString(hStr);
                return str;
            }
        }

        public bool HasWindow(object desktop, IntPtr window)
        {
            return ((IComVirtualDesktop)desktop).GetId() == VirtualDesktopManager.GetWindowDesktopId(window);
        }

        public bool IsCurrentDesktop(IntPtr hMon, object desktop)
        {
            return ReferenceEquals((IComVirtualDesktop)desktop, VirtualDesktopManagerInternal.GetCurrentDesktop(hMon));
        }

        public bool IsWindowOnCurrentDesktop(IntPtr window)
        {
            return VirtualDesktopManager.IsWindowOnCurrentVirtualDesktop(window);
        }

        public bool IsWindowPinned(IntPtr window)
        {
            ApplicationViewCollection.GetViewForHwnd(window, out var view);
            return VirtualDesktopPinnedApps.IsViewPinned(view);
        }

        public void SwitchToDesktop(IntPtr hMon, object desktop)
        {
            VirtualDesktopManagerInternal.SwitchDesktop(hMon, (IComVirtualDesktop)desktop);
        }

        public List<object> GetVirtualDesktops(IntPtr hMon)
        {
            return EnumerateVirtualDesktops(hMon).Cast<object>().ToList();
        }

        public int GetDesktopIndex(IntPtr hMon, object desktop)
        {
            Guid idSearch = ((IComVirtualDesktop)desktop).GetId();
            int i = 0;
            foreach (var d in EnumerateVirtualDesktops(hMon))
            {
                if (idSearch == d.GetId())
                {
                    return i;
                }
                i++;
            }

            return -1;
        }

        public object GetDesktopByIndex(IntPtr hMon, int index)
        {
            int count = VirtualDesktopManagerInternal.GetCount(hMon);
            if (index < 0 || index >= count) throw new ArgumentOutOfRangeException(nameof(index));

            VirtualDesktopManagerInternal.GetDesktops(hMon, out IComObjectArray desktops);
            desktops.GetAt(index, typeof(IComVirtualDesktop).GUID, out object objdesktop);
            Marshal.ReleaseComObject(desktops);
            return (IComVirtualDesktop)objdesktop;
        }

        public void MoveToDesktop(IntPtr hWnd, object desktop)
        {
            ApplicationViewCollection.GetViewForHwnd(hWnd, out var view);
            VirtualDesktopManagerInternal.MoveViewToDesktop(view, (IComVirtualDesktop)desktop);
        }

        private IEnumerable<IComVirtualDesktop> EnumerateVirtualDesktops(IntPtr hMon)
        {
            VirtualDesktopManagerInternal.GetDesktops(hMon, out IComObjectArray desktops);
            for (int i = 0; i < VirtualDesktopManagerInternal.GetCount(hMon); i++)
            {
                desktops.GetAt(i, typeof(IComVirtualDesktop).GUID, out object objdesktop);
                yield return (IComVirtualDesktop)objdesktop;
            }
            _ = Marshal.ReleaseComObject(desktops);
        }
    }
}
