using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;

using static WinMan.Implementation.Win32.NativeMethods;

namespace WinMan.Implementation.Win32
{
    internal class Win32DisplayManager : IDisplayManager
    {
        // TODO: Implements hotplug detection
        public event DisplayChangedHandler Added;
        public event DisplayChangedHandler Removed;

        public Rectangle VirtualDisplayBounds
        {
            get
            {
                int x = GetSystemMetrics(SystemMetric.SM_XVIRTUALSCREEN);
                int y = GetSystemMetrics(SystemMetric.SM_YVIRTUALSCREEN);
                int width = GetSystemMetrics(SystemMetric.SM_CXVIRTUALSCREEN);
                int height = GetSystemMetrics(SystemMetric.SM_CYVIRTUALSCREEN);

                return new Rectangle(x, y, x + width, y + height);
            }
        }

        public IDisplay PrimaryDisplay => GetDisplays().First(x => x.Bounds.TopLeft == new Point(0, 0));

        public IReadOnlyList<IDisplay> Displays => GetDisplays();

        public IWorkspace Workspace { get; }

        public Win32DisplayManager(IWorkspace workspace)
        {
            Workspace = workspace;
        }

        private IReadOnlyList<Win32Display> GetDisplays()
        {
            List<Win32Display> displays = new List<Win32Display>();
            if (!EnumDisplayMonitors(IntPtr.Zero, IntPtr.Zero, delegate(IntPtr hMonitor, IntPtr hdcMonitor, ref RECT lprcMonitor, IntPtr dwData)
            {
                if (IsVisibleMonitor(hMonitor))
                {
                    displays.Add(new Win32Display(this, hMonitor));
                }
                return true;
            }, IntPtr.Zero))
            {
                throw new Win32Exception();
            }
            return displays;
        }

        internal Rectangle GetWorkArea(IntPtr hMonitor)
        {
            var rect = GetMonitorInfo(hMonitor).rcWork;
            return new Rectangle(rect.LEFT, rect.TOP, rect.RIGHT, rect.BOTTOM);
        }

        internal Rectangle GetBounds(IntPtr hMonitor)
        {
            var rect = GetMonitorInfo(hMonitor).rcMonitor;
            return new Rectangle(rect.LEFT, rect.TOP, rect.RIGHT, rect.BOTTOM);
        }

        private bool IsVisibleMonitor(IntPtr hMonitor)
        {
            return (GetMonitorInfo(hMonitor).dwFlags & DISPLAY_DEVICE_MIRRORING_DRIVER) == 0;
        }

        private MONITORINFO GetMonitorInfo(IntPtr hMonitor)
        {
            MONITORINFO mi = default;
            mi.cbSize = Marshal.SizeOf<MONITORINFO>();

            if (!NativeMethods.GetMonitorInfo(hMonitor, ref mi))
            {
                throw new Win32Exception();
            }

            return mi;
        }
    }
}