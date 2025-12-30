using System;
using System.Linq;
using System.Runtime.InteropServices;

using static WinMan.Windows.IWin32VirtualDesktopService;

namespace WinMan.Windows
{
    public class Win32VirtualDesktop : IVirtualDesktop
    {
        private readonly Win32Workspace m_workspace;
        private readonly IWin32VirtualDesktopService m_vds;
        private readonly Desktop m_desktop;
        private readonly IntPtr m_hMon;

        public event EventHandler<DesktopChangedEventArgs>? Removed;

        internal Win32VirtualDesktop(Win32Workspace workspace, IWin32VirtualDesktopService vds, Desktop desktop)
        {
            m_workspace = workspace;
            m_vds = vds;
            m_desktop = desktop;
        }

        public bool IsAlive => m_workspace.VirtualDesktopManager.Desktops.Contains(this);

        public bool IsCurrent
        {
            get
            {
                try
                {
                    return m_vds.IsCurrentDesktop(m_hMon, m_desktop);
                }
                catch (COMException)
                {
                    return false;
                }
            }
        }

        public int Index
        {
            get
            {
                try
                {
                    return m_vds.GetDesktopIndex(m_hMon, m_desktop);
                }
                catch (COMException)
                {
                    return m_vds.GetDesktopCount(m_hMon) - 1;
                }
            }
        }

        public string Name
        {
            get
            {
                try
                {
                    return m_vds.GetDesktopName(m_desktop);
                }
                catch (COMException)
                {
                    return $"Desktop {m_vds.GetDesktopCount(m_hMon) - 1}";
                }
            }
        }

        public IWorkspace Workspace => m_workspace;

        internal Guid Guid => m_desktop.Guid;

        public bool HasWindow(IWindow window)
        {
            try
            {
                return m_vds.HasWindow(m_desktop, window.Handle) || m_vds.IsWindowPinned(window.Handle);
            }
            catch (COMException e) when ((uint)e.HResult == /*TYPE_E_ELEMENTNOTFOUND*/ 0x8002802B)
            {
                return false;
            }
        }

        public void SwitchTo()
        {
            m_vds.SwitchToDesktop(m_hMon, m_desktop);
        }

        public override bool Equals(object? obj)
        {
            return obj is Win32VirtualDesktop desktop &&
                   m_desktop.Equals(desktop.m_desktop);
        }

        public override int GetHashCode()
        {
            return -1324198676 + m_desktop.GetHashCode();
        }

        public override string ToString()
        {
            return $"Win32VirtualDesktop {{ Name = {Name}, IsCurrent = {IsCurrent} }}";
        }

        internal void OnRemoved()
        {
            try
            {
                Removed?.Invoke(this, new DesktopChangedEventArgs(this));
            }
            finally
            {
                Removed = null;
            }
        }

        public void MoveWindow(IWindow window)
        {
            try
            {
                m_vds.MoveToDesktop(window.Handle, m_desktop);
            }
            catch (Exception e) when (!window.IsAlive)
            {
                throw new InvalidWindowReferenceException(window.Handle, e);
            }
        }

        public void SetName(string newName)
        {
            throw new NotImplementedException();
        }

        public void Remove()
        {
            throw new NotImplementedException();
        }
    }
}