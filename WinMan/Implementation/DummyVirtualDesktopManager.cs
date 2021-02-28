using System;
using System.Collections.Generic;

namespace WinMan.Implementation
{
    internal class DummyVirtualDesktopManager : IVirtualDesktopManager
    {
        public IWorkspace Workspace { get; }

        public bool CanManageVirtualDesktops => false;

        public IReadOnlyList<IVirtualDesktop> Desktops { get; }

        public IVirtualDesktop CurrentDesktop => Desktops[0];


#pragma warning disable
        public event VirtualDesktopChangedEventHandler CurrentDesktopChanged;
        public event VirtualDesktopAddedEventHandler DesktopAdded;
        public event VirtualDesktopRemovedEventHandler DesktopRemoved;
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
