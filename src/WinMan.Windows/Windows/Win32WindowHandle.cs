using System;
using System.Threading;

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

        public Win32WindowHandle(IntPtr handle)
        {
            Handle = handle;
        }

        public void EnsureWindowObject(Win32Workspace workspace)
        {
            if (m_windowObject == null)
            {
                var window = new Win32Window(workspace, Handle);
                Interlocked.CompareExchange(ref m_windowObject, window, null);
            }
        }
    }
}
