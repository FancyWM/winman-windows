using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

using WinMan.Windows.Com;
using WinMan.Windows.Com.Build22000;

using WinMan.Windows.DllImports;

namespace WinMan.Windows
{
    internal class Win32VirtualDesktopService22000 : IWin32VirtualDesktopService
    {
        private class Desktop
        {
            public Guid Guid { get; init; }
            public IntPtr Monitor { get; init; }

            public Desktop(IntPtr hMon, IComVirtualDesktop comDesktop)
            {
                Guid = comDesktop.GetId();
                Monitor = hMon;
            }
        }

        private IComVirtualDesktopManagerInternal VirtualDesktopManagerInternal;
        private IComVirtualDesktopManager VirtualDesktopManager;
        private IComApplicationViewCollection ApplicationViewCollection;
        private IComVirtualDesktopPinnedApps VirtualDesktopPinnedApps;

        public Win32VirtualDesktopService22000()
        {
            Connect();
        }

        public void Connect()
        {
            var shell = (IComServiceProvider10?)Activator.CreateInstance(Type.GetTypeFromCLSID(ComGuids.CLSID_ImmersiveShell, true)!)
                ?? throw new COMException($"Failed to create instance of {ComGuids.CLSID_ImmersiveShell}");
            VirtualDesktopManagerInternal = (IComVirtualDesktopManagerInternal)shell.QueryService(ComGuids.CLSID_VirtualDesktopManagerInternal, typeof(IComVirtualDesktopManagerInternal).GUID);
            VirtualDesktopManager = (IComVirtualDesktopManager?)Activator.CreateInstance(Type.GetTypeFromCLSID(ComGuids.CLSID_VirtualDesktopManager, true)!)
                ?? throw new COMException($"Failed to create instance of {ComGuids.CLSID_VirtualDesktopManager}");
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
            IComVirtualDesktop comDesktop = GetComDesktop(desktop);
            // get desktop name
            string? desktopName = null;
            try
            {
                desktopName = new HSTRING(comDesktop.GetName()).MarshalIntoString();
            }
            catch { }

            // no name found, generate generic name
            if (string.IsNullOrEmpty(desktopName))
            { // create name "Desktop n" (n = number starting with 1)
                desktopName = "Desktop " + (GetDesktopIndex(IntPtr.Zero, comDesktop) + 1).ToString();
            }
            return desktopName;
        }

        public bool HasWindow(object desktop, IntPtr window)
        {
            return ((Desktop)desktop).Guid == VirtualDesktopManager.GetWindowDesktopId(window);
        }

        public bool IsCurrentDesktop(IntPtr hMon, object desktop)
        {
            return ReferenceEquals(GetComDesktop(desktop), VirtualDesktopManagerInternal.GetCurrentDesktop(hMon));
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
            VirtualDesktopManagerInternal.SwitchDesktop(hMon, GetComDesktop(desktop));
        }

        public List<object> GetVirtualDesktops(IntPtr hMon)
        {
            return EnumerateVirtualDesktops(hMon)
                .Select(x => new Desktop(hMon, x))
                .Cast<object>().ToList();
        }

        public int GetDesktopIndex(IntPtr hMon, object desktop)
        {
            Guid idSearch = desktop is IComVirtualDesktop comDesktop
                ? comDesktop.GetId()
                : ((Desktop)desktop).Guid;
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
            return new Desktop(hMon, (IComVirtualDesktop)objdesktop);
        }

        public void MoveToDesktop(IntPtr hWnd, object desktop)
        {
            ApplicationViewCollection.GetViewForHwnd(hWnd, out var view);
            VirtualDesktopManagerInternal.MoveViewToDesktop(view, GetComDesktop(desktop));
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

        private IComVirtualDesktop GetComDesktop(object desktopObj)
        {
            var desktop = (Desktop)desktopObj;
            var id = desktop.Guid;
            var hMon = desktop.Monitor;

            VirtualDesktopManagerInternal.GetDesktops(hMon, out IComObjectArray desktops);
            for (int i = 0; i < VirtualDesktopManagerInternal.GetCount(hMon); i++)
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
