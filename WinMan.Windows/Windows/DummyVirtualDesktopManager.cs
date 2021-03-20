using System;
using System.Collections.Generic;

namespace WinMan.Windows
{
    public class DummyVirtualDesktopManager : IVirtualDesktopManager
    {
        public IWorkspace Workspace { get; }

        public bool CanManageVirtualDesktops => false;

        public IReadOnlyList<IVirtualDesktop> Desktops { get; }

        public IVirtualDesktop CurrentDesktop => Desktops[0];


#pragma warning disable
        public event EventHandler<CurrentDesktopChangedEventArgs> CurrentDesktopChanged;
        public event EventHandler<DesktopChangedEventArgs> DesktopAdded;
        public event EventHandler<DesktopChangedEventArgs> DesktopRemoved;
#pragma warning restore

        public DummyVirtualDesktopManager(IWorkspace workspace)
        {
            Workspace = workspace;
            Desktops = new[]
            {
                new DummyVirtualDesktop(workspace),
            };
        }

        public bool IsWindowPinned(IWindow window)
        {
            return false;
        }

        public void PinWindow(IWindow window)
        {
        }

        public void UnpinWindow(IWindow window)
        {
        }

        public IVirtualDesktop CreateDesktop()
        {
            throw new NotSupportedException();
        }
    }
}
