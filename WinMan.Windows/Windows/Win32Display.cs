using System;
using System.Collections.Generic;

using static WinMan.Windows.NativeMethods;

namespace WinMan.Windows
{
    public class Win32Display : IDisplay
    {
        private static readonly bool IsPerMonitorDPISupported = Environment.OSVersion.Version.Major > 6
            || (Environment.OSVersion.Version.Major == 6 && Environment.OSVersion.Version.Minor >= 3);

        public event EventHandler<DisplayChangedEventArgs>? Removed;

        public event EventHandler<DisplayRectangleChangedEventArgs>? WorkAreaChanged;

        public event EventHandler<DisplayRectangleChangedEventArgs>? BoundsChanged;

        public event EventHandler<DisplayScalingChangedEventArgs>? ScalingChanged;

        public Rectangle WorkArea
        {
            get
            {
                lock (m_syncRoot)
                {
                    return m_workArea;
                }
            }
        }

        public Rectangle Bounds
        {
            get
            {
                lock (m_syncRoot)
                {
                    return m_bounds;
                }
            }
        }

        public double Scaling
        {
            get
            {
                lock (m_syncRoot)
                {
                    return m_scaling;
                }
            }
        }

        public IWorkspace Workspace => m_manager.Workspace;

        internal IntPtr Handle => m_hMonitor;

        private readonly Win32DisplayManager m_manager;
        private readonly IntPtr m_hMonitor;

        private readonly object m_syncRoot = new object();
        private Rectangle m_workArea;
        private Rectangle m_bounds;
        private double m_scaling;

        internal Win32Display(Win32DisplayManager manager, IntPtr hMonitor)
        {
            m_manager = manager;
            m_hMonitor = hMonitor;

            m_workArea = m_manager.GetWorkArea(m_hMonitor);
            m_bounds = m_manager.GetBounds(m_hMonitor);
            m_scaling = GetDpiScale();
        }

        internal void OnSettingChange()
        {
            Rectangle newWorkArea;
            Rectangle newBounds;
            double newScaling;
            try
            {
                newWorkArea = m_manager.GetWorkArea(m_hMonitor);
                newBounds = m_manager.GetBounds(m_hMonitor);
                newScaling = GetDpiScale();
            }
            catch (InvalidDisplayReferenceException)
            {
                return;
            }

            if (newWorkArea != m_workArea)
            {
                Rectangle oldWorkArea;
                lock (m_syncRoot)
                {
                    oldWorkArea = m_workArea;
                    m_workArea = newWorkArea;
                }
                WorkAreaChanged?.Invoke(this, new DisplayRectangleChangedEventArgs(this, newWorkArea, oldWorkArea));
            }

            if (newBounds != m_bounds)
            {
                Rectangle oldBounds;
                lock (m_syncRoot)
                {
                    oldBounds = m_bounds;
                    m_bounds = newBounds;
                }
                BoundsChanged?.Invoke(this, new DisplayRectangleChangedEventArgs(this, newBounds, oldBounds));
            }

            if (newScaling != m_scaling)
            {
                double oldScaling;
                lock (m_syncRoot)
                {
                    oldScaling = m_scaling;
                    m_scaling = newScaling;
                }
                ScalingChanged?.Invoke(this, new DisplayScalingChangedEventArgs(this, newScaling, oldScaling));
            }
        }

        internal void OnRemoved()
        {
            Removed?.Invoke(this, new DisplayChangedEventArgs(this));
        }

        private double GetDpiScale()
        {
            if (IsPerMonitorDPISupported)
            {
                NT_6_3.GetDpiForMonitor(m_hMonitor, NT_6_3.MDT_EFFECTIVE_DPI, out uint dpiX, out _);
                return dpiX / 96.0;
            }
            else
            {
                return 1.0;
            }
        }

        public override bool Equals(object obj)
        {
            return obj is Win32Display display &&
                   EqualityComparer<IntPtr>.Default.Equals(Handle, display.Handle);
        }

        public override int GetHashCode()
        {
            return 1786700523 + Handle.GetHashCode();
        }
    }
}