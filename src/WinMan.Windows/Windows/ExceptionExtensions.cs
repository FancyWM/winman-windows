using System;
using System.ComponentModel;

using WinMan.Windows.DllImports;

namespace WinMan.Windows
{
    internal static class ExceptionExtensions
    {
        internal static bool IsInvalidWindowHandleException(this Exception e)
        {
            return e is Win32Exception ex && (ex.NativeErrorCode == Constants.ERROR_INVALID_HANDLE || ex.NativeErrorCode == Constants.ERROR_INVALID_WINDOW_HANDLE);
        }
        internal static bool IsInvalidMonitorHandleException(this Exception e)
        {
            return e is Win32Exception ex && (ex.NativeErrorCode == Constants.ERROR_INVALID_HANDLE || ex.NativeErrorCode == Constants.ERROR_INVALID_MONITOR_HANDLE);
        }
    }
}
