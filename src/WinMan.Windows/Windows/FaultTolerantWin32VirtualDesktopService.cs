using System;
using System.Collections.Generic;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Threading;

using static WinMan.Windows.IWin32VirtualDesktopService;

namespace WinMan.Windows
{
    internal class FaultTolerantWin32VirtualDesktopService : IWin32VirtualDesktopService
    {
        private const int RPC_S_SERVER_UNAVAILABLE = unchecked((int)0x800706BA);
        private const int RPC_S_CALL_FAILED = unchecked((int)0x800706BE);
        private const int REGDB_E_CLASSNOTREG = unchecked((int)0x80040154);

        private readonly IWin32VirtualDesktopService m_vds;

        public FaultTolerantWin32VirtualDesktopService(IWin32VirtualDesktopService vds)
        {
            m_vds = vds;
        }

        public void Connect()
        {
            m_vds.Connect();
        }

        public int GetCurrentDesktopIndex(IntPtr hMon)
        {
            return ExecuteWithRetry(() => m_vds.GetCurrentDesktopIndex(hMon));
        }

        public Desktop GetDesktopByIndex(IntPtr hMon, int index)
        {
            return ExecuteWithRetry(() => m_vds.GetDesktopByIndex(hMon, index));
        }

        public int GetDesktopCount(IntPtr hMon)
        {
            return ExecuteWithRetry(() => m_vds.GetDesktopCount(hMon));
        }

        public int GetDesktopIndex(IntPtr hMon, Desktop m_desktop)
        {
            return ExecuteWithRetry(() => m_vds.GetDesktopIndex(hMon, m_desktop));
        }

        public string GetDesktopName(Desktop desktop)
        {
            return ExecuteWithRetry(() => m_vds.GetDesktopName(desktop));
        }

        public List<Desktop> GetVirtualDesktops(IntPtr hMon)
        {
            return ExecuteWithRetry(() => m_vds.GetVirtualDesktops(hMon));
        }

        public bool HasWindow(Desktop desktop, IntPtr hWnd)
        {
            return ExecuteWithRetry(() => m_vds.HasWindow(desktop, hWnd));
        }

        public bool IsCurrentDesktop(IntPtr hMon, Desktop desktop)
        {
            return ExecuteWithRetry(() => m_vds.IsCurrentDesktop(hMon, desktop));
        }

        public bool IsWindowOnCurrentDesktop(IntPtr hWnd)
        {
            return ExecuteWithRetry(() => m_vds.IsWindowOnCurrentDesktop(hWnd));
        }

        public bool IsWindowPinned(IntPtr hWnd)
        {
            return ExecuteWithRetry(() => m_vds.IsWindowPinned(hWnd));
        }

        public void MoveToDesktop(IntPtr hWnd, Desktop desktop)
        {
            ExecuteWithRetry(() => m_vds.MoveToDesktop(hWnd, desktop));
        }

        public void SwitchToDesktop(IntPtr hMon, Desktop desktop)
        {
            ExecuteWithRetry(() => m_vds.SwitchToDesktop(hMon, desktop));
        }

        private T ExecuteWithRetry<T>(Func<T> func)
        {
            ExceptionDispatchInfo? exception = null;
            for (int i = 0; i < 10; i++)
            {
                try
                {
                    if (i != 0)
                    {
                        Thread.Sleep(500);
                        m_vds.Connect();
                    }
                    return func();
                }
                catch (COMException e) when (e.HResult == RPC_S_SERVER_UNAVAILABLE || e.HResult == RPC_S_CALL_FAILED || e.HResult == REGDB_E_CLASSNOTREG)
                {
                    exception = ExceptionDispatchInfo.Capture(e);
                }
                catch (NotImplementedException e)
                {
                    exception = ExceptionDispatchInfo.Capture(e);
                }
            }
            exception!.Throw();
            throw new InvalidProgramException();
        }

        private void ExecuteWithRetry(Action action)
        {
            ExceptionDispatchInfo? exception = null;
            for (int i = 0; i < 10; i++)
            {
                try
                {
                    if (i != 0)
                    {
                        Thread.Sleep(500);
                        m_vds.Connect();
                    }
                    action();
                    return;
                }
                catch (COMException e) when (e.HResult == RPC_S_SERVER_UNAVAILABLE || e.HResult == RPC_S_CALL_FAILED || e.HResult == REGDB_E_CLASSNOTREG)
                {
                    exception = ExceptionDispatchInfo.Capture(e);
                }
                catch (NotImplementedException e)
                {
                    exception = ExceptionDispatchInfo.Capture(e);
                }
            }
            exception!.Throw();
            throw new InvalidProgramException();
        }
    }
}
