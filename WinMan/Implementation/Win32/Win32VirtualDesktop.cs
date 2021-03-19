using System;
using System.Runtime.InteropServices;
using WinMan.Implementation.Win32.VirtualDesktop;

namespace WinMan.Implementation.Win32
{
    internal class Win32VirtualDesktop : IVirtualDesktop
    {
        private readonly Win32Workspace m_workspace;
        private readonly Desktop m_desktop;

        public event EventHandler<DesktopChangedEventArgs> Removed;

        public Win32VirtualDesktop(Win32Workspace workspace, Desktop desktop)
        {
            m_workspace = workspace;
            m_desktop = desktop;
        }

        public bool IsCurrent => m_desktop.IsVisible;

        public int Index => Desktop.FromDesktop(m_desktop);

        public string Name => Desktop.DesktopNameFromDesktop(m_desktop);

        public IWorkspace Workspace => m_workspace;

        public void MoveWindow(IWindow window)
        {
            try
            {
                m_desktop.MoveWindow(window.Handle);
            }
            catch (COMException e) when ((uint)e.HResult == /*TYPE_E_ELEMENTNOTFOUND*/ 0x8002802B)
            {
                if (!window.IsAlive)
                {
                    throw new InvalidWindowReferenceException(window.Handle);
                }
            }
        }

        public bool HasWindow(IWindow window)
        {
            try
            {
                return m_desktop.HasWindow(window.Handle) || m_workspace.VirtualDesktopManager.IsWindowPinned(window);
            }
            catch (COMException e) when ((uint)e.HResult == /*TYPE_E_ELEMENTNOTFOUND*/ 0x8002802B)
            {
                return false;
            }
        }

        public void SwitchTo()
        {
            m_desktop.MakeVisible();
        }

        public void SetName(string newName)
        {
            m_desktop.SetName(newName);
        }

        public void Remove()
        {
            m_desktop.Remove();
        }

        public override bool Equals(object obj)
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
    }
}