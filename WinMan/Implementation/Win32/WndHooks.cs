using System;
using System.Runtime.InteropServices;

namespace WinMan.Implementation.Win32
{
    internal static class WndHooks
    {
        internal enum HookID : int
        {
            WH_CALLWNDPROC = 4,
            WH_CALLWNDPROCRET = 12,
        }

        internal static bool InstallWindowMessageHook(IntPtr hwndRecv, int wmRecv, HookID idHook, int wmMin, int wmMax)
        {
            if (Marshal.SizeOf<IntPtr>() == 4)
            {
                return WndHooksX32.InstallWindowMessageHook(hwndRecv.ToInt32(), wmRecv, (int)idHook, wmMin, wmMax);
            }
            else
            {
                return WndHooksX64.InstallWindowMessageHook(hwndRecv.ToInt64(), wmRecv, (int)idHook, wmMin, wmMax);
            }
        }

        internal static bool RemoveWindowMessageHook(IntPtr hwndRecv, int wmRecv, HookID idHook, int wmMin, int wmMax)
        {
            if (Marshal.SizeOf<IntPtr>() == 4)
            {
                return WndHooksX32.RemoveWindowMessageHook(hwndRecv.ToInt32(), wmRecv, (int)idHook, wmMin, wmMax);
            }
            else
            {
                return WndHooksX64.RemoveWindowMessageHook(hwndRecv.ToInt64(), wmRecv, (int)idHook, wmMin, wmMax);
            }
        }

        private static class WndHooksX32
        {
            [DllImport("HookProcDllx32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool InstallWindowMessageHook(int hwndRecv, int wmRecv, int idHook, int wmMin, int wmMax);

            [DllImport("HookProcDllx32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool RemoveWindowMessageHook(int hwndRecv, int wmRecv, int idHook, int wmMin, int wmMax);
        }

        private static class WndHooksX64
        {
            [DllImport("HookProcDllx64.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool InstallWindowMessageHook(long hwndRecv, int wmRecv, int idHook, int wmMin, int wmMax);

            [DllImport("HookProcDllx64.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool RemoveWindowMessageHook(long hwndRecv, int wmRecv, int idHook, int wmMin, int wmMax);
        }
    }
}
