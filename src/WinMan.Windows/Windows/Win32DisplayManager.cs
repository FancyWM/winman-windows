using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

using Microsoft.Win32;

using WinMan.Windows.DllImports;
using static WinMan.Windows.DllImports.Constants;
using static WinMan.Windows.DllImports.NativeMethods;

namespace WinMan.Windows
{
    public class Win32DisplayManager : IDisplayManager
    {
        // TODO: Implements hotplug detection
        public event EventHandler<DisplayChangedEventArgs>? Added;
        public event EventHandler<DisplayChangedEventArgs>? Removed;
#pragma warning disable CS0067
        public event EventHandler<DisplayRectangleChangedEventArgs>? VirtualDisplayBoundsChanged;
#pragma warning restore CS0067
        public event EventHandler<PrimaryDisplayChangedEventArgs>? PrimaryDisplayChanged;

        private static readonly bool IsPerMonitorDPISupported = Environment.OSVersion.Version.Major > 6
            || (Environment.OSVersion.Version.Major == 6 && Environment.OSVersion.Version.Minor >= 3);

        public Rectangle VirtualDisplayBounds
        {
            get
            {
                int x = GetSystemMetrics(GetSystemMetrics_nIndexFlags.SM_XVIRTUALSCREEN);
                int y = GetSystemMetrics(GetSystemMetrics_nIndexFlags.SM_YVIRTUALSCREEN);
                int width = GetSystemMetrics(GetSystemMetrics_nIndexFlags.SM_CXVIRTUALSCREEN);
                int height = GetSystemMetrics(GetSystemMetrics_nIndexFlags.SM_CYVIRTUALSCREEN);

                return new Rectangle(x, y, x + width, y + height);
            }
        }

        public IDisplay PrimaryDisplay { get; private set; }

        public IReadOnlyList<IDisplay> Displays
        {
            get
            {
                lock (m_displays)
                {
                    return m_displays.OrderBy(x => x.DeviceID).ToList();
                }
            }
        }

        public IWorkspace Workspace => m_workspace;

        // Appears to be the device name used when the user is not logged in or
        // for RDP connections. I could not find any documentation, and it appears
        // that this device always returns a resoltion of 1024x768.
        // I suspect it is also used when the graphics driver crashes.
        private const string VirtualDeviceName = "WinDisc";
        // Fake DeviceID we use when we cannot find any displays.
        private const string NoMonitorID = "NoMonitor";

        private const int VirtualDeviceDefaultRefreshRate = 30;

        private readonly Win32Workspace m_workspace;

        private readonly HashSet<Win32Display> m_displays;

        public Win32DisplayManager(Win32Workspace workspace)
        {
            m_workspace = workspace;
            m_displays = new HashSet<Win32Display>(GetVisibleDisplayMonitors().Select(x => new Win32Display(this, GetDeviceID(x.deviceName))));
            PrimaryDisplay = m_displays.First(x => x.Bounds.TopLeft == new Point(0, 0));
        }

        internal void OnDisplayChange()
        {
            var addedDisplays = new List<Win32Display>();
            var removedDisplays = new List<Win32Display>();

            IDisplay oldPrimaryDisplay;
            IDisplay newPrimaryDisplay;
            lock (m_displays)
            {
                oldPrimaryDisplay = PrimaryDisplay;
                string[] freshDeviceIds = WaitForVisibleDisplayMonitors()
                    .Select(x => GetDeviceIDOrNull(x.deviceName))
                    .Where(x => x != null)
                    .ToArray()!;
                if (freshDeviceIds.Length == 0)
                {
                    freshDeviceIds = [NoMonitorID];
                }

                IEnumerable<string> existingDeviceIds = m_displays.Select(x => x.DeviceID);

                string[] added = freshDeviceIds.Except(existingDeviceIds).ToArray();
                string[] removed = existingDeviceIds.Except(freshDeviceIds).ToArray();

                foreach (var id in removed)
                {
                    var disp = m_displays.First(x => x.DeviceID == id);
                    removedDisplays.Add(disp);
                }
                m_displays.RemoveWhere(x => removed.Contains(x.DeviceID));

                foreach (var id in added)
                {
                    try
                    {
                        var disp = new Win32Display(this, id);
                        m_displays.Add(disp);
                        addedDisplays.Add(disp);
                    }
                    catch (InvalidDisplayReferenceException)
                    {
                        // ignore
                    }
                }

                newPrimaryDisplay = m_displays.FirstOrDefault(x => x.Bounds.TopLeft == new Point(0, 0))
                    ?? m_displays.First();
                if (!oldPrimaryDisplay.Equals(newPrimaryDisplay))
                {
                    PrimaryDisplay = newPrimaryDisplay;
                }
            }

            try
            {
                // Added events
                try
                {
                    foreach (var added in addedDisplays)
                    {
                        Added?.Invoke(added, new DisplayChangedEventArgs(added));
                    }
                }
                finally
                {
                    // Removed events
                    foreach (var removed in removedDisplays)
                    {
                        try
                        {
                            removed.OnRemoved();
                        }
                        finally
                        {
                            Removed?.Invoke(removed, new DisplayChangedEventArgs(removed));
                        }
                    }
                }
            }
            finally
            {
                if (!oldPrimaryDisplay.Equals(newPrimaryDisplay))
                {
                    PrimaryDisplayChanged?.Invoke(this, new PrimaryDisplayChangedEventArgs(newPrimaryDisplay, oldPrimaryDisplay));
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
        internal struct DisplayInfo
        {
            public string DeviceID;
            public Rectangle WorkArea;
            public Rectangle Bounds;
            public double DPIScale;
            public int RefreshRate;
        };

        internal DisplayInfo GetDisplayInfo(string deviceID)
        {
            if (deviceID == NoMonitorID)
            {
                return new DisplayInfo
                {
                    DeviceID = NoMonitorID,
                    Bounds = new Rectangle(0, 0, 1024, 768),
                    WorkArea = new Rectangle(0, 0, 1024, 768),
                    DPIScale = 1,
                    RefreshRate = 60,
                };
            }

            try
            {
                var hMonitor = GetMonitorHandle(deviceID);
                return GetDisplayInfo(hMonitor);
            }
            catch (Win32Exception)
            {
                throw new InvalidDisplayReferenceException(deviceID);
            }
        }

        private DisplayInfo GetDisplayInfo(IntPtr hMonitor)
        {
            try
            {
                var (mi, device, refreshRate) = GetDisplayInfoInternal(hMonitor);
                var dpiScale = GetDPIScale(hMonitor);

                return new DisplayInfo
                {
                    DeviceID = device,
                    WorkArea = new Rectangle(mi.rcWork.left, mi.rcWork.top, mi.rcWork.right, mi.rcWork.bottom),
                    Bounds = new Rectangle(mi.rcMonitor.left, mi.rcMonitor.top, mi.rcMonitor.right, mi.rcMonitor.bottom),
                    DPIScale = dpiScale,
                    RefreshRate = refreshRate,
                };
            }
            catch (Win32Exception e) when (e.IsInvalidMonitorHandleException() || !IsMonitorValid(hMonitor))
            {
                throw new InvalidDisplayReferenceException(hMonitor);
            }
        }

        private (MONITORINFO mi, string deviceId, int refreshRate) GetDisplayInfoInternal(IntPtr hMonitor)
        {
            var (mi, deviceName) = GetMonitorInfoEx(hMonitor);
            DEVMODEW devMode = default;
            if (!EnumDisplaySettings(deviceName, ENUM_DISPLAY_SETTINGS_MODE.ENUM_CURRENT_SETTINGS, ref devMode))
            {
                if (!IsMonitorValid(hMonitor))
                {
                    throw new InvalidDisplayReferenceException(hMonitor);
                }
                if (deviceName == VirtualDeviceName)
                {
                    return (mi, GetDeviceID(deviceName), GetVirtualMonitorRefreshRate());
                }
                else
                {
                    throw new Win32Exception($"Could not read the settings for monitor \"{deviceName}\".");
                }
            }

            return (mi, GetDeviceID(deviceName), (int)devMode.dmDisplayFrequency);
        }

        private List<IntPtr> GetAllDisplayMonitors()
        {
            List<IntPtr> monitors = [];
            unsafe
            {
                // Handling the error code from EnumDisplayMonitors is problematic, because not all fail modes are
                // documented. Microsoft's own samples and WPF sources do not handle the errors from this call.
                EnumDisplayMonitors(new(), (RECT*)null, delegate (HMONITOR hMonitor, HDC hdcMonitor, RECT* lprcMonitor, LPARAM dwData)
                {
                    monitors.Add(hMonitor);
                    return true;
                }, new LPARAM());

                if (monitors.Count == 0)
                {
                    throw new Win32Exception().WithMessage("Could not enumerate the display monitors attached to the system!");
                }
            }
            return monitors;
        }

        private List<(IntPtr hMonitor, MONITORINFO info, string deviceName)> GetVisibleDisplayMonitors()
        {
            return GetAllDisplayMonitors().Select(hMonitor =>
            {
                var (info, name) = GetMonitorInfoEx(hMonitor);
                return (hMonitor, info, name);
            }).Where(x => IsVisibleMonitor(x.info)).ToList();
        }

        private List<(IntPtr hMonitor, MONITORINFO info, string deviceName)> WaitForVisibleDisplayMonitors()
        {
            for (int i = 0; i < 10; i++)
            {
                if (i != 0)
                {
                    Thread.Sleep(500);
                }
                var monitors = GetVisibleDisplayMonitors();
                if (monitors.Count > 0)
                {
                    return monitors;
                }
            }
            throw new IndexOutOfRangeException("EnumDisplayMonitors returned no monitors. The operation was retied but failed.");
        }

        private bool IsVisibleMonitor(MONITORINFO info)
        {
            return (info.dwFlags & DISPLAY_DEVICE_MIRRORING_DRIVER) == 0;
        }

        private double GetDPIScale(IntPtr hMonitor)
        {
            if (IsPerMonitorDPISupported)
            {
                try
                {
                    NT_6_3.GetDpiForMonitor(new(hMonitor), MONITOR_DPI_TYPE.MDT_EFFECTIVE_DPI, out uint dpiX, out _);
                    return dpiX / 96.0;
                }
                catch (Win32Exception e) when (e.IsInvalidMonitorHandleException() || !IsMonitorValid(hMonitor))
                {
                    throw new InvalidDisplayReferenceException(hMonitor);
                }
            }
            else
            {
                return 1.0;
            }
        }

        private (MONITORINFO info, string deviceName) GetMonitorInfoEx(IntPtr hMonitor)
        {
            unsafe
            {
                MONITORINFOEXW miEx = default;
                MONITORINFO* pmi = (MONITORINFO*)&miEx;
                (*pmi).cbSize = (uint)sizeof(MONITORINFOEXW);
                if (!NativeMethods.GetMonitorInfo(new(hMonitor), pmi))
                {
                    try
                    {
                        throw new Win32Exception().WithMessage($"Could not read the monitor information for HMONITOR={hMonitor:X8}!");
                    }
                    catch (Win32Exception e) when (e.IsInvalidMonitorHandleException() || !IsMonitorValid(hMonitor))
                    {
                        throw new InvalidDisplayReferenceException(hMonitor, e);
                    }
                }

                char* pszDevice = (char*)&miEx.szDevice;
                return (*pmi, new string(pszDevice));
            }
        }


        private string GetDeviceID(IntPtr hMonitor)
        {
            var (_, deviceName) = GetMonitorInfoEx(hMonitor);
            return GetDeviceID(deviceName);
        }

        private string? GetDeviceIDOrNull(string deviceName)
        {
            DISPLAY_DEVICEW ddInterface = default;
            ddInterface.cb = (uint)Marshal.SizeOf<DISPLAY_DEVICEW>();
            if (!EnumDisplayDevices(deviceName, 0, ref ddInterface, EDD_GET_DEVICE_INTERFACE_NAME))
            {
                return null;
            }

            return MarshalExtensions.MarshalIntoString(ddInterface.DeviceID.AsSpan());
        }

        private string GetDeviceID(string deviceName)
        {
            if (GetDeviceIDOrNull(deviceName) is string deviceId)
            {
                return deviceId;
            }
            throw new Win32Exception($"Could not read the device ID for monitor \"{deviceName}\".");
        }

        private IntPtr GetMonitorHandle(string deviceID)
        {
            foreach (var hMonitor in GetAllDisplayMonitors())
            {
                if (GetDeviceID(hMonitor) == deviceID)
                {
                    return hMonitor;
                }
            }
            throw new InvalidDisplayReferenceException(deviceID);
        }

        private bool IsMonitorValid(IntPtr hMonitor)
        {
            return GetAllDisplayMonitors().Contains(hMonitor);
        }

        private int GetVirtualMonitorRefreshRate()
        {
            try
            {
                using var rdpConfig = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations");
                var frameIntervalString = rdpConfig!.GetValue("DWMFRAMEINTERVAL", null)?.ToString();
                if (frameIntervalString == null)
                    return VirtualDeviceDefaultRefreshRate;
                int frameInterval = int.Parse(frameIntervalString);
                return (int)(1000.0 / frameInterval);
            }
            catch
            {
                return VirtualDeviceDefaultRefreshRate;
            }
        }
    }
}
