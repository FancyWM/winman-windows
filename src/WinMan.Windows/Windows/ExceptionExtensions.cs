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
        internal static bool IsAccessDeniedException(this Exception e)
        {
            return e is Win32Exception ex && ((uint)ex.HResult == /*E_ACCESSDENIED*/ 0x80070005 || ex.NativeErrorCode == /* ERROR_ACCESS_DENIED */ 5);
        }
    }
}
