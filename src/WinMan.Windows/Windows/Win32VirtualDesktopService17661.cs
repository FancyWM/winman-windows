using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

using WinMan.Windows.Com;
using WinMan.Windows.Com.Build17661;

using static WinMan.Windows.IWin32VirtualDesktopService;

namespace WinMan.Windows
{
    internal class Win32VirtualDesktopService17661 : IWin32VirtualDesktopService
    {
        private IComVirtualDesktopManagerInternal VirtualDesktopManagerInternal;
        private IComVirtualDesktopManager VirtualDesktopManager;
        private IComApplicationViewCollection ApplicationViewCollection;
        private IComVirtualDesktopPinnedApps VirtualDesktopPinnedApps;

        public Win32VirtualDesktopService17661()
        {
            Connect();
        }

        public void Connect()
        {
            var shell = (IComServiceProvider10?)Activator.CreateInstance(Type.GetTypeFromCLSID(ComGuids.CLSID_ImmersiveShell, true)!)
                ?? throw new COMException($"Failed to create instance of {ComGuids.CLSID_ImmersiveShell}");
            VirtualDesktopManagerInternal = (IComVirtualDesktopManagerInternal?)shell.QueryService(ComGuids.CLSID_VirtualDesktopManagerInternal, typeof(IComVirtualDesktopManagerInternal).GUID)
                ?? throw new COMException($"Failed to create instance of {ComGuids.CLSID_VirtualDesktopManagerInternal}");
            VirtualDesktopManager = (IComVirtualDesktopManager?)Activator.CreateInstance(Type.GetTypeFromCLSID(ComGuids.CLSID_VirtualDesktopManager, true)!)
                ?? throw new COMException($"Failed to create instance of {ComGuids.CLSID_VirtualDesktopManager}");
            ApplicationViewCollection = (IComApplicationViewCollection)shell.QueryService(typeof(IComApplicationViewCollection).GUID, typeof(IComApplicationViewCollection).GUID);
            VirtualDesktopPinnedApps = (IComVirtualDesktopPinnedApps)shell.QueryService(ComGuids.CLSID_VirtualDesktopPinnedApps, typeof(IComVirtualDesktopPinnedApps).GUID);
        }

        public int GetCurrentDesktopIndex(IntPtr hMon)
        {
            return GetDesktopIndex(hMon, new Desktop(VirtualDesktopManagerInternal.GetCurrentDesktop().GetId()));
        }

        public int GetDesktopCount(IntPtr _hMon)
        {
            return VirtualDesktopManagerInternal.GetCount();
        }

        public string GetDesktopName(Desktop desktop)
        {
            IComVirtualDesktop comDesktop = GetComDesktop(desktop);
            // return name of desktop or "Desktop n" if it has no name
            Guid guid = comDesktop.GetId();

            // read desktop name in registry
            string? desktopName = null;
            try
            {
                desktopName = (string?)Microsoft.Win32.Registry.GetValue("HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\VirtualDesktops\\Desktops\\{" + guid.ToString() + "}", "Name", null);
            }
            catch { }

            // no name found, generate generic name
            if (string.IsNullOrEmpty(desktopName))
            { // create name "Desktop n" (n = number starting with 1)
                desktopName = "Desktop " + (GetDesktopIndex(IntPtr.Zero, desktop) + 1).ToString();
            }
            return desktopName;
        }

        public bool HasWindow(Desktop desktop, IntPtr window)
        {
            return desktop.Guid == VirtualDesktopManager.GetWindowDesktopId(window);
        }

        public bool IsCurrentDesktop(IntPtr _hMon, Desktop desktop)
        {
            return desktop.Guid == VirtualDesktopManagerInternal.GetCurrentDesktop().GetId();
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

        public void SwitchToDesktop(IntPtr hMon, Desktop desktop)
        {
            VirtualDesktopManagerInternal.SwitchDesktop(GetComDesktop(desktop));
        }

        public List<Desktop> GetVirtualDesktops(IntPtr hMon)
        {
            return EnumerateVirtualDesktops().Select(objdesktop => new Desktop(objdesktop.GetId())).ToList();
        }

        public int GetDesktopIndex(IntPtr hWnd, Desktop desktop)
        {
            Guid idSearch = desktop.Guid;
            int i = 0;
            foreach (var d in EnumerateVirtualDesktops())
            {
                if (idSearch == d.GetId())
                {
                    return i;
                }
                i++;
            }

            return -1;
        }

        public Desktop GetDesktopByIndex(IntPtr hMon, int index)
        {
            int count = VirtualDesktopManagerInternal.GetCount();
            if (index < 0 || index >= count) throw new ArgumentOutOfRangeException(nameof(index));

            VirtualDesktopManagerInternal.GetDesktops(out IComObjectArray desktops);
            desktops.GetAt(index, typeof(IComVirtualDesktop).GUID, out object objdesktop);
            Marshal.ReleaseComObject(desktops);
            return new Desktop(((IComVirtualDesktop)objdesktop).GetId());
        }

        public void MoveToDesktop(IntPtr hWnd, Desktop desktop)
        {
            ApplicationViewCollection.GetViewForHwnd(hWnd, out var view);
            VirtualDesktopManagerInternal.MoveViewToDesktop(view, GetComDesktop(desktop));
        }

        private IEnumerable<IComVirtualDesktop> EnumerateVirtualDesktops()
        {
            VirtualDesktopManagerInternal.GetDesktops(out IComObjectArray desktops);
            for (int i = 0; i < VirtualDesktopManagerInternal.GetCount(); i++)
            {
                desktops.GetAt(i, typeof(IComVirtualDesktop).GUID, out object objdesktop);
                yield return (IComVirtualDesktop)objdesktop;
            }
            _ = Marshal.ReleaseComObject(desktops);
        }

        private IComVirtualDesktop GetComDesktop(Desktop desktop)
        {
            var id = desktop.Guid;

            VirtualDesktopManagerInternal.GetDesktops(out IComObjectArray desktops);
            for (int i = 0; i < VirtualDesktopManagerInternal.GetCount(); i++)
            {
                desktops.GetAt(i, typeof(IComVirtualDesktop).GUID, out object objdesktop);
                if (objdesktop is IComVirtualDesktop d && d.GetId() == id)
                {
                    _ = Marshal.ReleaseComObject(desktops);
                    return d;
                }
            }
            _ = Marshal.ReleaseComObject(desktops);
            throw new COMException("Desktop does not exist!");
        }
    }
}
