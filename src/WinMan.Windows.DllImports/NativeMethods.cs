using System.Runtime.InteropServices;

namespace WinMan.Windows
{
    internal static partial class NativeMethods
    {
        internal static class NT_6_3
        {
            [DllImport("shcore.dll")]
            internal static extern uint GetDpiForMonitor(HMONITOR hmonitor, MONITOR_DPI_TYPE dpiType, out uint dpiX, out uint dpiY);
        }
    }
}
