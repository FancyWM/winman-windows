﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

using WinMan.Windows.Utilities;

using WinMan.Windows.DllImports;
using static WinMan.Windows.DllImports.Constants;
using static WinMan.Windows.DllImports.NativeMethods;

namespace WinMan.Windows
{
    internal delegate void WinEventCallback(uint eventType, IntPtr hwnd, int idObject, int idChild, uint idEventThread, uint dwmsEventTime);

    internal class WinEventHookHelper
    {
        public static Deleter CreateGlobalOutOfContextHook(ISet<uint> eventTypes, WinEventCallback handler)
        {
            uint min = eventTypes.Min();
            uint max = eventTypes.Max();

            WINEVENTPROC? callback = delegate (HWINEVENTHOOK hWinEventHook, uint eventType,
                HWND hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
            {
                if (eventTypes.Contains(eventType))
                {
                    handler(eventType, hwnd, idObject, idChild, dwEventThread, dwmsEventTime);
                }
            };

            IntPtr hHook = SetWinEventHook(min, max, IntPtr.Zero, callback, idProcess: 0, idThread: 0, WINEVENT_OUTOFCONTEXT);
            if (IntPtr.Zero == hHook)
            {
                throw new Win32Exception().WithMessage("Could not set a window message event hook!");
            }

            return new Deleter(() =>
            {
                // Capture and null callback
                callback = null;
                UnhookWinEvent(new(hHook));
            });
        }
    }
}
