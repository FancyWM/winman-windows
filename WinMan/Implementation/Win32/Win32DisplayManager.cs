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
        public event EventHandler<DisplayChangedEventArgs> Added;
        public event EventHandler<DisplayChangedEventArgs> Removed;
        public event EventHandler<DisplayRectangleChangedEventArgs> VirtualDisplayBoundsChanged;
        public event EventHandler<PrimaryDisplayChangedEventArgs> PrimaryDisplayChanged;

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

        public IDisplay PrimaryDisplay => Displays.First(x => x.Bounds.TopLeft == new Point(0, 0));

        public IReadOnlyList<IDisplay> Displays
        {
            get
            {
                lock (m_displays)
                {
                    return m_displays.ToList();
                }
            }
        }

        public IWorkspace Workspace => m_workspace;

        private readonly Win32Workspace m_workspace;

        private readonly HashSet<Win32Display> m_displays;

        public Win32DisplayManager(Win32Workspace workspace)
        {
            m_workspace = workspace;
            m_displays = GetMonitors().Select(x => new Win32Display(this, x)).ToHashSet();
        }

        private List<IntPtr> GetMonitors()
        {
            List<IntPtr> monitors = new List<IntPtr>();
            if (!EnumDisplayMonitors(IntPtr.Zero, IntPtr.Zero, delegate (IntPtr hMonitor, IntPtr hdcMonitor, ref RECT lprcMonitor, IntPtr dwData)
            {
                if (IsVisibleMonitor(hMonitor))
                {
                    monitors.Add(hMonitor);
                }
                return true;
            }, IntPtr.Zero))
            {
                throw new Win32Exception();
            }
            return monitors;
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

        internal void OnDisplayChange()
        {
            var addedDisplays = new List<Win32Display>();
            var removedDisplays = new List<Win32Display>();

            lock (m_displays)
            {
                var newMonitors = GetMonitors();
                var handles = m_displays.Select(x => x.Handle);

                var added = newMonitors.Except(handles).ToList();
                var removed = handles.Except(newMonitors).ToList();

                foreach (var hMonitor in removed)
                {
                    var disp = m_displays.First(x => x.Handle == hMonitor);
                    removedDisplays.Add(disp);
                }
                m_displays.RemoveWhere(x => removed.Contains(x.Handle));

                foreach (var hMonitor in added)
                {
                    var disp = new Win32Display(this, hMonitor);
                    m_displays.Add(disp);
                    addedDisplays.Add(disp);
                }
            }

            try
            {
                foreach (var added in addedDisplays)
                {
                    Added?.Invoke(added, new DisplayChangedEventArgs(added));
                }
            }
            finally
            {
                foreach (var removed in removedDisplays)
                {
                    try
                    {
                        removed.OnRemoved();
                    }
                    catch
                    {
                        Removed?.Invoke(removed, new DisplayChangedEventArgs(removed));
                    }
                }
            }
        }

        internal void OnSettingChange()
        {
            List<Win32Display> displays;
            lock (m_displays)
            {
                displays = m_displays.ToList();
            }

            foreach (var d in displays)
            {
                d.OnSettingChange();
            }
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
                try
                {
                    throw new Win32Exception();
                }
                catch (Win32Exception e) when (e.IsInvalidMonitorHandleException())
                {
                    throw new InvalidDisplayReferenceException(hMonitor, e);
                }
            }

            return mi;
        }
    }
}