using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

using WinMan.Windows.Com;
using WinMan.Windows.Com.BuildGeneric;

using static WinMan.Windows.IWin32VirtualDesktopService;

namespace WinMan.Windows.Windows
{
    internal class Win32VirtualDesktopServiceGeneric : IWin32VirtualDesktopService
    {
        private VirtualDesktopManagerInternalProxy VirtualDesktopManagerInternal;
        private IComVirtualDesktopManager VirtualDesktopManager;
        private IComApplicationViewCollection ApplicationViewCollection;
        private IComVirtualDesktopPinnedApps VirtualDesktopPinnedApps;

        public Win32VirtualDesktopServiceGeneric()
        {
            Connect();
        }

        public void Connect()
        {
            VirtualDesktopManagerInternal = VirtualDesktopManagerInternalProxy.TryCreateSlow()
                ?? throw new COMException($"Failed to determine internal COM GUIDs!");
            var shell = (IComServiceProvider10?)Activator.CreateInstance(Type.GetTypeFromCLSID(ComGuids.CLSID_ImmersiveShell, true)!)
                ?? throw new COMException($"Failed to create instance of {ComGuids.CLSID_ImmersiveShell}");
            VirtualDesktopManager = (IComVirtualDesktopManager?)Activator.CreateInstance(Type.GetTypeFromCLSID(ComGuids.CLSID_VirtualDesktopManager, true)!)
                ?? throw new COMException($"Failed to create instance of {ComGuids.CLSID_VirtualDesktopManager}");
            ApplicationViewCollection = (IComApplicationViewCollection)shell.QueryService(typeof(IComApplicationViewCollection).GUID, typeof(IComApplicationViewCollection).GUID);
            VirtualDesktopPinnedApps = (IComVirtualDesktopPinnedApps)shell.QueryService(ComGuids.CLSID_VirtualDesktopPinnedApps, typeof(IComVirtualDesktopPinnedApps).GUID);
        }

        public int GetCurrentDesktopIndex(IntPtr hMon)
        {
            return GetDesktopIndex(hMon, new Desktop(hMon, VirtualDesktopManagerInternal.GetCurrentDesktop(hMon).GetId()));
        }

        public Guid GetCurrentDesktopGuid(IntPtr hMon)
        {
            return VirtualDesktopManagerInternal.GetCurrentDesktop(hMon).GetId();
        }

        public Guid GetWindowDesktopGuid(IntPtr hWnd)
        {
            return VirtualDesktopManager.GetWindowDesktopId(hWnd);
        }

        public int GetDesktopCount(IntPtr hMon)
        {
            return VirtualDesktopManagerInternal.GetCount(hMon);
        }

        public string GetDesktopName(Desktop desktop)
        {
            return "Desktop " + (GetDesktopIndex(IntPtr.Zero, desktop) + 1).ToString();
        }

        public bool HasWindow(Desktop desktop, IntPtr window)
        {
            return desktop.Guid == VirtualDesktopManager.GetWindowDesktopId(window);
        }

        public bool IsCurrentDesktop(IntPtr hMon, Desktop desktop)
        {
            return desktop.Guid == VirtualDesktopManagerInternal.GetCurrentDesktop(hMon).GetId();
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
            VirtualDesktopManagerInternal.SwitchDesktop(hMon, GetComDesktop(hMon, desktop));
        }

        public List<Desktop> GetVirtualDesktops(IntPtr hMon)
        {
            return EnumerateVirtualDesktops(hMon).ToList();
        }

        public int GetDesktopIndex(IntPtr hMon, Desktop desktop)
        {
            Guid idSearch = desktop.Guid;
            int i = 0;
            foreach (var d in EnumerateVirtualDesktops(hMon))
            {
                if (idSearch == d.Guid)
                {
                    return i;
                }
                i++;
            }

            return -1;
        }

        public Desktop GetDesktopByIndex(IntPtr hMon, int index)
        {
            int count = VirtualDesktopManagerInternal.GetCount(hMon);
            if (index < 0 || index >= count) throw new ArgumentOutOfRangeException(nameof(index));

            var desktop = VirtualDesktopManagerInternal.GetDesktopAtIndex(hMon, index);
            return new Desktop(hMon, desktop.GetId());
        }

        public void MoveToDesktop(IntPtr hWnd, Desktop desktop)
        {
            ApplicationViewCollection.GetViewForHwnd(hWnd, out var view);
            VirtualDesktopManagerInternal.MoveViewToDesktop(view, GetComDesktop(hWnd, desktop));
        }

        private IEnumerable<Desktop> EnumerateVirtualDesktops(IntPtr hMon)
        {
            var desktops = VirtualDesktopManagerInternal.GetDesktops(hMon);
            for (int i = 0; i < VirtualDesktopManagerInternal.GetCount(hMon); i++)
            {
                yield return new Desktop(hMon, desktops[i].GetId());
            }
        }

        private VirtualDesktopProxy GetComDesktop(IntPtr hMon, Desktop desktop)
        {
            var desktops = VirtualDesktopManagerInternal.GetDesktops(hMon);
            for (int i = 0; i < desktops.Count; i++)
            {
                if (desktops[i].GetId() == desktop.Guid)
                {
                    return desktops[i];
                }
            }
            throw new COMException("Desktop does not exist!");
        }
    }
}
