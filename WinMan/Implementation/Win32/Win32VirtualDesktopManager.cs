using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

using WinMan.Implementation.Win32.VirtualDesktop;

namespace WinMan.Implementation.Win32
{
    internal class Win32VirtualDesktopManager : IVirtualDesktopManager
    {
        private readonly object m_syncRoot = new object();

        private readonly Win32Workspace m_workspace;

        private readonly List<Win32VirtualDesktop> m_desktops;

        private int m_currentDesktop;

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
                return m_desktops[Desktop.FromDesktop(Desktop.Current)];
            }
        }

        public event EventHandler<DesktopChangedEventArgs> DesktopAdded;
        public event EventHandler<DesktopChangedEventArgs> DesktopRemoved;
        public event EventHandler<CurrentDesktopChangedEventArgs> CurrentDesktopChanged;

        public Win32VirtualDesktopManager(Win32Workspace workspace)
        {
            m_workspace = workspace;
            m_desktops = new List<Win32VirtualDesktop>();
            for (int i = 0; i < Desktop.Count; i++)
            {
                m_desktops.Add(new Win32VirtualDesktop(workspace, Desktop.FromIndex(i)));
            }
            m_currentDesktop = Desktop.FromDesktop(Desktop.Current);
        }

        public IVirtualDesktop CreateDesktop()
        {
            var desktopInternal = Desktop.Create();
            var desktop = new Win32VirtualDesktop(m_workspace, desktopInternal);
            lock (m_syncRoot)
            {
                m_desktops.Add(desktop);
            }
            return desktop;
        }

        public bool IsWindowPinned(IWindow window)
        {
            try
            {
                return Desktop.IsWindowPinned(window.Handle);
            }
            catch (COMException e) when ((uint)e.HResult == /*TYPE_E_ELEMENTNOTFOUND*/ 0x8002802B)
            {
                if (!window.IsAlive)
                {
                    throw new InvalidWindowReferenceException(window.Handle);
                }
                return false;
            }
        }

        public void PinWindow(IWindow window)
        {
            try
            {
                Desktop.PinWindow(window.Handle);
            }
            catch (COMException e) when ((uint)e.HResult == /*TYPE_E_ELEMENTNOTFOUND*/ 0x8002802B)
            {
                if (!window.IsAlive)
                {
                    throw new InvalidWindowReferenceException(window.Handle);
                }
            }
        }

        public void UnpinWindow(IWindow window)
        {
            try 
            { 
                Desktop.UnpinWindow(window.Handle);
            }
            catch (COMException e) when ((uint)e.HResult == /*TYPE_E_ELEMENTNOTFOUND*/ 0x8002802B)
            {
                if (!window.IsAlive)
                {
                    throw new InvalidWindowReferenceException(window.Handle);
                }
            }
        }

        internal bool IsNotOnCurrentDesktop(IntPtr hwnd)
        {
            try
            {
                return !DesktopManager.VirtualDesktopManager.IsWindowOnCurrentVirtualDesktop(hwnd);
            }
            catch (COMException e) when ((uint)e.HResult == /*TYPE_E_ELEMENTNOTFOUND*/ 0x8002802B)
            {
                return true;
            }
        }

        internal void CheckVirtualDesktopChanges()
        {
            List<Win32VirtualDesktop> removedDesktops = new List<Win32VirtualDesktop>();
            List<Win32VirtualDesktop> addedDesktops = new List<Win32VirtualDesktop>();

            int newDesktopCount = Desktop.Count;
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
                        var newInstance = new Win32VirtualDesktop(m_workspace, Desktop.FromIndex(i));
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

            int newCurrentDesktop = Desktop.FromDesktop(Desktop.Current);
            int oldCurrentDesktop;
            IVirtualDesktop newDesktop;
            IVirtualDesktop oldDesktop;

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
    }
}