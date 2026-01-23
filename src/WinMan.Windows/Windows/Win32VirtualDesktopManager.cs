using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;

using WinMan.Windows.DllImports;
using WinMan.Windows.Utilities;

namespace WinMan.Windows
{
    public class Win32VirtualDesktopManager : IVirtualDesktopManager, IWin32VirtualDesktopManagerInternal
    {
        private readonly object m_syncRoot = new object();

        private readonly Win32Workspace m_workspace;

        private readonly List<Win32VirtualDesktop> m_desktops;

        private int m_currentDesktop;

        private readonly IWin32VirtualDesktopService m_vds;

        private readonly IntPtr m_hMon;

        private readonly Func<Guid> m_getCurrentDesktopGuid;

        public IWorkspace Workspace => m_workspace;

        public bool CanManageVirtualDesktops => true;

        public IReadOnlyList<IVirtualDesktop> Desktops
        {
            get
            {
                lock (m_syncRoot)
                {
                    return m_desktops.ToArray();
                }
            }
        }

        public IVirtualDesktop CurrentDesktop
        {
            get
            {
                return GetCurrentDesktop();
            }
        }

        public event EventHandler<DesktopChangedEventArgs>? DesktopAdded;
        public event EventHandler<DesktopChangedEventArgs>? DesktopRemoved;
        public event EventHandler<CurrentDesktopChangedEventArgs>? CurrentDesktopChanged;

        internal Win32VirtualDesktopManager(Win32Workspace workspace, IWin32VirtualDesktopService vds, IntPtr hMon)
        {
            m_workspace = workspace;
            m_desktops = new List<Win32VirtualDesktop>();
            m_hMon = hMon;
            m_vds = vds;

            foreach (var d in m_vds.GetVirtualDesktops(m_hMon))
            {
                m_desktops.Add(new Win32VirtualDesktop(workspace, m_vds, d));
            }

            m_currentDesktop = m_vds.GetCurrentDesktopIndex(m_hMon);

            try
            {
                var comGuid = GetCurrentDesktopGuidFromCom();
                var registryGuid = VirtualDesktopRegistryHelper.GetCurrentDesktopGuid();
                // This should be true on Windows 11. Windows 10 may have it in the expected registry location.
                if (comGuid == registryGuid)
                {
                    m_getCurrentDesktopGuid = GetCurrentDesktopGuidFromRegistry;
                }
                else
                {
                    m_getCurrentDesktopGuid = GetCurrentDesktopGuidFromCom;
                }
            }
            catch
            {
                m_getCurrentDesktopGuid = GetCurrentDesktopGuidFromCom;
            }
        }

        public IVirtualDesktop CreateDesktop()
        {
            throw new NotImplementedException();
        }

        public bool IsWindowPinned(IWindow window)
        {
            if (window == null)
            {
                throw new ArgumentNullException(nameof(window));
            }

            try
            {
                return m_vds.IsWindowPinned(window.Handle);
            }
            catch (COMException e) when ((uint)e.HResult == /*TYPE_E_ELEMENTNOTFOUND*/ 0x8002802B)
            {
                return false;
            }
        }

        public void PinWindow(IWindow window)
        {
            throw new NotImplementedException();
        }

        public void UnpinWindow(IWindow window)
        {
            throw new NotImplementedException();
        }

        bool IWin32VirtualDesktopManagerInternal.IsNotOnCurrentDesktop(IntPtr hwnd)
        {
            try
            {
                return !m_vds.IsWindowOnCurrentDesktop(hwnd);
            }
            catch (COMException e) when ((uint)e.HResult == /*TYPE_E_ELEMENTNOTFOUND*/ 0x8002802B)
            {
                return true;
            }
        }

        void IWin32VirtualDesktopManagerInternal.CheckVirtualDesktopChanges()
        {
            List<Win32VirtualDesktop> removedDesktops = new List<Win32VirtualDesktop>();
            List<Win32VirtualDesktop> addedDesktops = new List<Win32VirtualDesktop>();

            int newDesktopCount = m_vds.GetDesktopCount(m_hMon);
            int oldDesktopCount;
            lock (m_syncRoot)
            {
                oldDesktopCount = m_desktops.Count;
                if (oldDesktopCount > newDesktopCount)
                {
                    for (int i = newDesktopCount; i < oldDesktopCount; i++)
                    {
                        removedDesktops.Add(m_desktops[i]);
                    }
                    m_desktops.RemoveRange(oldDesktopCount - 1, oldDesktopCount - newDesktopCount);
                }
                else if (oldDesktopCount < newDesktopCount)
                {
                    for (int i = oldDesktopCount; i < newDesktopCount; i++)
                    {
                        var newInstance = new Win32VirtualDesktop(m_workspace, m_vds, m_vds.GetDesktopByIndex(m_hMon, i));
                        m_desktops.Add(newInstance);
                        addedDesktops.Add(newInstance);
                    }
                }
            }

            List<Exception> exs = new List<Exception>();

            foreach (var desktop in removedDesktops)
            {
                try
                {
                    desktop.OnRemoved();
                }
                catch (Exception e)
                {
                    exs.Add(e);
                }
                finally
                {
                    try
                    {
                        DesktopRemoved?.Invoke(this, new DesktopChangedEventArgs(desktop));
                    }
                    catch (Exception e)
                    {
                        exs.Add(e);
                    }
                }
            }

            foreach (var desktop in addedDesktops)
            {
                try
                {
                    DesktopAdded?.Invoke(this, new DesktopChangedEventArgs(desktop));
                }
                catch (Exception e)
                {
                    exs.Add(e);
                }
            }

            if (exs.Count > 0)
            {
                throw new AggregateException(exs);
            }

            int newCurrentDesktop = m_vds.GetCurrentDesktopIndex(m_hMon);
            int oldCurrentDesktop;
            IVirtualDesktop newDesktop;
            IVirtualDesktop? oldDesktop;

            lock (m_syncRoot)
            {
                oldCurrentDesktop = m_currentDesktop;
                oldDesktop = oldCurrentDesktop < m_desktops.Count ? m_desktops[oldCurrentDesktop] : null;
                if (newCurrentDesktop != oldCurrentDesktop)
                {
                    m_currentDesktop = newCurrentDesktop;
                }
                newDesktop = m_desktops[newCurrentDesktop];
            }

            if (oldCurrentDesktop != newCurrentDesktop)
            {
                CurrentDesktopChanged?.Invoke(this, new CurrentDesktopChangedEventArgs(newDesktop, oldDesktop));
            }
        }

        private Guid GetCurrentDesktopGuidFromRegistry()
        {
#if DEBUG
            var expectedGuid = m_vds.GetCurrentDesktopGuid(m_hMon);
#endif

            var guidFromRegistryTimer = BlockTimer.Create();
            Guid guid = VirtualDesktopRegistryHelper.GetCurrentDesktopGuid();
            guidFromRegistryTimer.LogIfExceeded(15, nameof(VirtualDesktopRegistryHelper.GetCurrentDesktopGuid));

            if (guid != default)
            {
                var guidTimer = BlockTimer.Create();
                guid = m_vds.GetCurrentDesktopGuid(m_hMon);
                guidTimer.LogIfExceeded(15, nameof(m_vds.GetCurrentDesktopGuid));
            }
#if DEBUG
            else
            {
                var comGuid = m_vds.GetCurrentDesktopGuid(m_hMon);
                if (comGuid == expectedGuid)
                {
                    Debug.Assert(guid == expectedGuid, $"{guid} == {expectedGuid}");
                }
            }
#endif
            return guid;
        }


        private Guid GetCurrentDesktopGuidFromCom()
        {
            return m_vds.GetCurrentDesktopGuid(m_hMon);
        }

        private Win32VirtualDesktop GetCurrentDesktop()
        {
            var guid = m_getCurrentDesktopGuid();
            lock (m_desktops)
            {
                for (int i = 0; i < m_desktops.Count; i++)
                {
                    if (m_desktops[i].Guid == guid)
                    {
                        return m_desktops[i];
                    }
                }
                // If the above operation fails, that means that CheckVirtualDesktopsChanged
                // hasn't had the chance to run yet, so return the last recorded virtual desktop
                // or whatever is valid.
                return m_desktops[Math.Min(m_currentDesktop, m_desktops.Count - 1)];
            }
        }
    }
}