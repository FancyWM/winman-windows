using System;

namespace WinMan.Implementation.Win32
{
    internal class Win32Display : IDisplay
    {
        public Rectangle WorkArea => m_manager.GetWorkArea(m_hMonitor);

        public Rectangle Bounds => m_manager.GetBounds(m_hMonitor);

        public IWorkspace Workspace => m_manager.Workspace;

        public event DisplayChangedHandler Removed;

        private readonly Win32DisplayManager m_manager;
        private readonly IntPtr m_hMonitor;

        internal Win32Display(Win32DisplayManager manager, IntPtr hMonitor)
        {
            m_manager = manager;
            m_hMonitor = hMonitor;
        }
    }
}