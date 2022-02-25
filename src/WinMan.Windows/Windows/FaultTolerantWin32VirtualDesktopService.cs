using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace WinMan.Windows
{
    internal class FaultTolerantWin32VirtualDesktopService : IWin32VirtualDesktopService
    {
        private const int RPC_S_SERVER_UNAVAILABLE = unchecked((int)0x800706BA);
        private const int RPC_S_CALL_FAILED = unchecked((int)0x800706BE);

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

        public object GetDesktopByIndex(IntPtr hMon, int index)
        {
            return ExecuteWithRetry(() => m_vds.GetDesktopByIndex(hMon, index));
        }

        public int GetDesktopCount(IntPtr hMon)
        {
            return ExecuteWithRetry(() => m_vds.GetDesktopCount(hMon));
        }

        public int GetDesktopIndex(IntPtr hMon, object m_desktop)
        {
            return ExecuteWithRetry(() => m_vds.GetDesktopIndex(hMon, m_desktop));
        }

        public string GetDesktopName(object desktop)
        {
            return ExecuteWithRetry(() => m_vds.GetDesktopName(desktop));
        }

        public List<object> GetVirtualDesktops(IntPtr hMon)
        {
            return ExecuteWithRetry(() => m_vds.GetVirtualDesktops(hMon));
        }

        public bool HasWindow(object desktop, IntPtr hWnd)
        {
            return ExecuteWithRetry(() => m_vds.HasWindow(desktop, hWnd));
        }

        public bool IsCurrentDesktop(IntPtr hMon, object desktop)
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

        public void MoveToDesktop(IntPtr hWnd, object desktop)
        {
            ExecuteWithRetry(() => m_vds.MoveToDesktop(hWnd, desktop));
        }

        public void SwitchToDesktop(IntPtr hMon, object desktop)
        {
            ExecuteWithRetry(() => m_vds.SwitchToDesktop(hMon, desktop));
        }

        private T ExecuteWithRetry<T>(Func<T> func)
        {
            Exception? exception = null;
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
                catch (COMException e) when (e.HResult == RPC_S_SERVER_UNAVAILABLE || e.HResult == RPC_S_CALL_FAILED)
                {
                    exception = e;
                }
                catch (NotImplementedException e)
                {
                    exception = e;
                }
            }
            throw exception!;
        }

        private void ExecuteWithRetry(Action action)
        {
            Exception? exception = null;
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
                catch (COMException e) when (e.HResult == RPC_S_SERVER_UNAVAILABLE || e.HResult == RPC_S_CALL_FAILED)
                {
                    exception = e;
                }
                catch (NotImplementedException e)
                {
                    exception = e;
                }
            }
            throw exception!;
        }
    }
}
