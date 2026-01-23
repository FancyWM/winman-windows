using System;
using System.Diagnostics;
using System.Reflection;
using System.Threading;

using WinMan.Windows.Utilities;

namespace WinMan.Windows
{
    /// <summary>
    /// Win32WindowHandles are much lighter in terms of memory consumption when compared to Win32Window.
    /// Since the user facing API will only ever see several dozen Win32Windows at any time, as compared
    /// to hundreds or thousands of references used in Win32Workspace, it makes sense to use this lighter abstraction
    /// and create the heavyweight objects only when necessary.
    /// </summary>
    internal class Win32WindowHandle
    {
        public IntPtr Handle { get; }

        public Win32Window? WindowObject => m_windowObject;

        private Win32Window? m_windowObject;

#if DEBUG
        public TimeSpan LastCheckedAt = TimeSpan.Zero;
        public int CheckCount = 0;

        private readonly TimeSpan m_createdAt = SteadyClock.Now;
#endif

        public Win32WindowHandle(IntPtr handle)
        {
            Handle = handle;
        }

        public void EnsureWindowObject(Win32Workspace workspace)
        {
            if (m_windowObject == null)
            {
#if DEBUG
                var now = SteadyClock.Now;
                var elapsed = now - m_createdAt;
#endif
                var window = new Win32Window(workspace, Handle);
                Interlocked.CompareExchange(ref m_windowObject, window, null);
#if DEBUG
                if (elapsed.TotalMilliseconds > 15)
                {
                    Debug.WriteLine($"[WRN] Block \"Create Win32Window ({window.Handle})\" took too long elapsed={elapsed.TotalMilliseconds}ms threshold=15ms");
                    Debug.WriteLine($"      Win32Window ({window.Handle}): ClassName={window.ClassName} Checks={CheckCount} SinceLastCheck={(now - LastCheckedAt).TotalMilliseconds}ms");
                }
#endif
            }
        }
    }
}
