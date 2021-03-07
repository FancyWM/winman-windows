using System;
using System.ComponentModel;

namespace WinMan.Implementation.Win32
{
    internal static class ExceptionExtensions
    {
        internal static bool IsInvalidHandleException(this Exception e)
        {
            return e is Win32Exception && e.Message == "Invalid window handle";
        }
    }
}
