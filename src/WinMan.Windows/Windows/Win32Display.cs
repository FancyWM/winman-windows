using System;
using System.Collections.Generic;

namespace WinMan.Windows
{
    public class Win32Display : IDisplay
    {
        public event EventHandler<DisplayChangedEventArgs>? Removed;

        public event EventHandler<DisplayRectangleChangedEventArgs>? WorkAreaChanged;

        public event EventHandler<DisplayRectangleChangedEventArgs>? BoundsChanged;

        public event EventHandler<DisplayScalingChangedEventArgs>? ScalingChanged;

        public event EventHandler<DisplayRefreshRateChangedEventArgs>? RefreshRateChanged;

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

        public int RefreshRate
        {
            get
            {
                // int read is atomic
                return m_refreshRate;
            }
        }

        public IWorkspace Workspace => m_manager.Workspace;

        internal IntPtr Handle => m_hMonitor;

        private readonly Win32DisplayManager m_manager;
        private readonly IntPtr m_hMonitor;

        private readonly object m_syncRoot = new object();
        private string m_deviceName;
        private Rectangle m_workArea;
        private Rectangle m_bounds;
        private double m_scaling;
        private int m_refreshRate;

        internal Win32Display(Win32DisplayManager manager, IntPtr hMonitor)
        {
            m_manager = manager;
            m_hMonitor = hMonitor;

            var (deviceName, workArea, bounds, scaling, refreshRate) = m_manager.GetMonitorSettings(m_hMonitor);
            m_deviceName = deviceName;
            m_workArea = workArea;
            m_bounds = bounds;
            m_scaling = scaling;
            m_refreshRate = refreshRate;
        }

        internal void OnSettingChange()
        {
            Rectangle newWorkArea;
            Rectangle newBounds;
            double newScaling;
            int newRefreshRate;
            try
            {
                // Device name cannot change
                var (_, workArea, bounds, scaling, refreshRate) = m_manager.GetMonitorSettings(m_hMonitor);
                newWorkArea = workArea;
                newBounds = bounds;
                newScaling = scaling;
                newRefreshRate = refreshRate;
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

            if (newRefreshRate != m_refreshRate)
            {
                int oldRefreshRate;
                lock (m_syncRoot)
                {
                    oldRefreshRate = m_refreshRate;
                    m_refreshRate = newRefreshRate;
                }
                RefreshRateChanged?.Invoke(this, new DisplayRefreshRateChangedEventArgs(this, newRefreshRate, oldRefreshRate));
            }
        }

        internal void OnRemoved()
        {
            Removed?.Invoke(this, new DisplayChangedEventArgs(this));
            Removed = null;
            WorkAreaChanged = null;
            BoundsChanged = null;
            ScalingChanged = null;
            RefreshRateChanged = null;
        }

        public bool Equals(IDisplay? other)
        {
            return Equals((object?)other);
        }

        public override bool Equals(object? obj)
        {
            return obj is Win32Display display &&
                   EqualityComparer<IntPtr>.Default.Equals(Handle, display.Handle);
        }

        public override int GetHashCode()
        {
            return 1786700523 + Handle.GetHashCode();
        }

        public override string ToString()
        {
            return $"Win32Display {{ {m_deviceName} }}";
        }
    }
}
