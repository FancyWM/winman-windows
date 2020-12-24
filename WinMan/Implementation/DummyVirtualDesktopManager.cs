using System;
using System.Collections.Generic;

namespace WinMan.Implementation
{
    internal class DummyVirtualDesktopManager : IVirtualDesktopManager
    {
        public IWorkspace Workspace { get; }

        public bool IsVirtualDesktopsSupported => false;

        public IReadOnlyList<IVirtualDesktop> Desktops { get; }

        public IVirtualDesktop CurrentDesktop => Desktops[0];


        public event VirtualDesktopChangedEventHandler CurrentDesktopChanged;
        public event VirtualDesktopChangedEventHandler DesktopAdded;
        public event VirtualDesktopChangedEventHandler DesktopRemoved;

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
            return true;
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
