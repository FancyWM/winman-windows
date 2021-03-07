﻿using System.Collections.Generic;

namespace WinMan
{
    public interface IVirtualDesktopManager
    {
        IWorkspace Workspace { get; }

        bool CanManageVirtualDesktops { get; }

        event VirtualDesktopAddedEventHandler DesktopAdded;
        event VirtualDesktopRemovedEventHandler DesktopRemoved;
        event VirtualDesktopChangedEventHandler CurrentDesktopChanged;

        IReadOnlyList<IVirtualDesktop> Desktops { get; }

        IVirtualDesktop CurrentDesktop { get; }

        IVirtualDesktop CreateDesktop();

        void PinWindow(IWindow window);

        void UnpinWindow(IWindow window);

        bool IsWindowPinned(IWindow window);
    }
}