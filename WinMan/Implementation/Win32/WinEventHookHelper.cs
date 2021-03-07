using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

using WinMan.Utilities;

using static WinMan.Implementation.Win32.NativeMethods;

namespace WinMan.Implementation.Win32
{
    internal delegate void WinEventCallback(uint eventType, IntPtr hwnd, int idObject, int idChild, uint idEventThread, uint dwmsEventTime);

    internal class WinEventHookHelper
    {
        public static Deleter CreateGlobalOutOfContextHook(ISet<uint> eventTypes, WinEventCallback handler)
        {
            uint min = eventTypes.Min();
            uint max = eventTypes.Max();

            WinEventProc callback = delegate (IntPtr hWinEventHook, uint eventType,
                IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
            {
                if (eventTypes.Contains(eventType))
                {
                    handler(eventType, hwnd, idObject, idChild, dwEventThread, dwmsEventTime);
                }
            };

            IntPtr hHook = SetWinEventHook(min, max, IntPtr.Zero, callback, idProcess: 0, idThread: 0, WINEVENT_OUTOFCONTEXT);
            if (IntPtr.Zero == hHook)
            {
                throw new Win32Exception();
            }

            return new Deleter(() =>
            {
                // Capture and null callback
                callback = null;
                UnhookWinEvent(hHook);
            });
        }
    }
}
