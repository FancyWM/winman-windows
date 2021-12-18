using System;
using System.ComponentModel;

using WinMan.Windows.DllImports;

namespace WinMan.Windows
{
    internal static class ExceptionExtensions
    {
        internal static bool IsInvalidWindowHandleException(this Win32Exception e)
        {
            return e.NativeErrorCode == Constants.ERROR_INVALID_HANDLE || e.NativeErrorCode == Constants.ERROR_INVALID_WINDOW_HANDLE;
        }
        
        internal static bool IsInvalidMonitorHandleException(this Win32Exception e)
        {
            return e.NativeErrorCode == Constants.ERROR_INVALID_HANDLE || e.NativeErrorCode == Constants.ERROR_INVALID_MONITOR_HANDLE;
        }

        internal static bool IsAccessDeniedException(this Win32Exception e)
        {
            return (uint)e.HResult == /*E_ACCESSDENIED*/ 0x80070005 || e.NativeErrorCode == /* ERROR_ACCESS_DENIED */ 5;
        }

        internal static bool IsTimeoutException(this Win32Exception e)
        {
            return e.NativeErrorCode == Constants.ERROR_TIMEOUT;
        }

        internal static Win32Exception WithMessage(this Win32Exception e, string message)
        {
            throw new Win32Exception(e.NativeErrorCode, $"{message}: {e.Message}");
        }
    }
}
