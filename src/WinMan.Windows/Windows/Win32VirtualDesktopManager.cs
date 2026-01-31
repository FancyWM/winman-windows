using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

using WinMan.Windows.Utilities;

namespace WinMan.Windows
{
    public class Win32VirtualDesktopManager : IVirtualDesktopManager, IWin32VirtualDesktopManagerInternal
    {
        private readonly object m_syncRoot = new object();

        private readonly Win32Workspace m_workspace;

        private List<Win32VirtualDesktop> m_desktops;

        private Guid m_currentDesktopGuid;

        private Win32VirtualDesktop? m_currentDesktop;

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

            m_currentDesktopGuid = m_vds.GetCurrentDesktopGuid(m_hMon);
            m_currentDesktop = m_desktops.FirstOrDefault(x => x.Guid == m_currentDesktopGuid);
            if (m_currentDesktop == null)
            {
                m_currentDesktopGuid = m_desktops[0].Guid;
                m_currentDesktop = m_desktops[0];
            }

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

        private void CheckVirtualDesktopListChanges()
        {
            var newDesktops = m_vds.GetVirtualDesktops(m_hMon);
            List<Guid> freshGuids = [.. newDesktops.Select(x => x.Guid)];

            var addedDesktops = new List<Win32VirtualDesktop>();
            var removedDesktops = new List<Win32VirtualDesktop>();

            lock (m_syncRoot)
            {
                if (m_desktops.Select(x => x.Guid).SequenceEqual(freshGuids))
                {
                    return;
                }

                List<Win32VirtualDesktop> updatedDesktops = [];
                for (int i = 0; i < freshGuids.Count; i++)
                {
                    var freshGuid = freshGuids[i];
                    var knownDesktop = m_desktops.FirstOrDefault(x => x.Guid == freshGuid);
                    if (knownDesktop != null)
                    {
                        updatedDesktops.Add(knownDesktop);
                    }
                    else
                    {
                        var newDesktop = new Win32VirtualDesktop(m_workspace, m_vds, new IWin32VirtualDesktopService.Desktop(m_hMon, freshGuid));
                        updatedDesktops.Add(newDesktop);
                        addedDesktops.Add(newDesktop);
                    }
                }

                removedDesktops = [.. m_desktops.Where(x => !freshGuids.Contains(x.Guid))];
                m_desktops = updatedDesktops;
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
        }

        void IWin32VirtualDesktopManagerInternal.CheckVirtualDesktopChanges()
        {
            var savedCurrentDesktop = CurrentDesktop;
            CheckVirtualDesktopListChanges();

            Guid newCurrentDesktop = m_vds.GetCurrentDesktopGuid(m_hMon);
            Guid oldCurrentDesktop;
            Win32VirtualDesktop? newDesktop = null;
            Win32VirtualDesktop? oldDesktop = null;

            lock (m_syncRoot)
            {
                oldCurrentDesktop = m_currentDesktopGuid;
                foreach (var d in m_desktops)
                {
                    if (d.Guid == oldCurrentDesktop)
                    {
                        oldDesktop = d;
                    }
                    if (d.Guid == newCurrentDesktop)
                    {
                        newDesktop = d;
                    }
                }
                if (newCurrentDesktop != oldCurrentDesktop)
                {
                    m_currentDesktopGuid = newCurrentDesktop;
                }
                if (newDesktop != null)
                {
                    m_currentDesktop = newDesktop;
                }
            }

            Debug.Assert(newDesktop != null);
            if (newDesktop == null)
            {
                // The current desktop GUID is not in the list of current desktops.
                // This can only happen during a race condition, which we cannot avoid.
                // The same race condition is seen in GetCurrentDesktop.
                // Best bet is to retry.
                (this as IWin32VirtualDesktopManagerInternal).CheckVirtualDesktopChanges();
                return;
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
                // hasn't had the chance to run yet, so return the previous desktop.
                Debug.Assert(m_currentDesktop != null);
                return m_currentDesktop!;
            }
        }
    }
}