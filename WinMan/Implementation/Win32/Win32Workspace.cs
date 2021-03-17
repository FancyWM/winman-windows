using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading;

using WinMan.Utilities;

using static WinMan.Implementation.Win32.NativeMethods;

namespace WinMan.Implementation.Win32
{
    class Win32Workspace : IWorkspace
    {
        public event WindowPresenceChangedEventHandler WindowAdded;
        public event WindowPresenceChangedEventHandler WindowRemoved;
        public event WindowPresenceChangedEventHandler WindowDestroyed;
        public event WindowPresenceChangedEventHandler WindowManaging;
        public event UnhandledExceptionEventHandler UnhandledException;

        private static readonly IntPtr IDT_TIMER_WATCH = new IntPtr(1);
        private const uint TIMER_WATCH_INTERVAL = 50;

        private IntPtr m_msgWnd;
        private Thread m_eventLoopThread;
        private Deleter m_winEventHook;
        private IntPtr m_hTimer;
        private bool m_isShuttingDown = false;

        private HashSet<Win32Window> m_visibleWindows = new HashSet<Win32Window>();
        private List<Win32Window> m_windowList = new List<Win32Window>();
        private Dictionary<IntPtr, Win32Window> m_windowSet = new Dictionary<IntPtr, Win32Window>();
        private IntPtr m_hwndFocused = IntPtr.Zero;

        private IVirtualDesktopManager m_virtualDesktops = null;

        public IWindow FocusedWindow
        {
            get
            {
                CheckOpen();

                IntPtr hwnd = GetForegroundWindow();
                if (hwnd == IntPtr.Zero)
                {
                    return null;
                }
                return FindWindow(hwnd);
            }
        }

        public IVirtualDesktopManager VirtualDesktopManager => VirtualDesktopManagerLazy;

        private IVirtualDesktopManager VirtualDesktopManagerLazy
        {
            get
            {
                IVirtualDesktopManager Init()
                {
                    try
                    {
                        return new Win32VirtualDesktopManager(this);
                    }
                    catch
                    {
                        return new DummyVirtualDesktopManager(this);
                    }
                }

                if (m_virtualDesktops == null)
                {
                    Interlocked.CompareExchange(ref m_virtualDesktops, Init(), null);
                }

                return m_virtualDesktops;
            }
        }


        private Win32DisplayManager m_displayManager;

        public IDisplayManager DisplayManager => m_displayManager;

        public Point CursorLocation
        {
            get
            {
                if (!GetCursorPos(out POINT pt))
                {
                    throw new Win32Exception();
                }
                return new Point(pt.X, pt.Y);
            }
        }

        public Win32Workspace()
        {
        }

        public void Open()
        {
            m_displayManager = new Win32DisplayManager(this);

            m_eventLoopThread = new Thread(EventLoop);
            m_eventLoopThread.Name = "Win32WorkspaceEventLoop";

            foreach (var window in GetWindowListImpl())
            {
                m_windowSet.Add(window.Handle, window);
                m_windowList.Add(window);

                if (window.IsTopLevelVisible)
                {
                    lock (m_visibleWindows)
                    {
                        m_visibleWindows.Add(window);
                    }

                    try
                    {
                        window.OnAdded();
                    }
                    finally
                    {
                        WindowManaging?.Invoke(window);
                    }
                }
            }

            m_eventLoopThread.Start();
        }

        public IWindow FindWindow(IntPtr windowHandle)
        {
            CheckOpen();
            lock (m_windowList)
            {
                return m_windowList.FirstOrDefault(x => x.Handle == windowHandle);
            }
        }

        public IReadOnlyList<IWindow> GetSnapshot()
        {
            CheckOpen();
            Win32Window[] windows = GetVisibleWindowList();
            return windows.Where(x => x.IsTopLevelVisible).ToArray();
        }

        public IComparer<IWindow> CreateSnapshotZOrderComparer()
        {
            Dictionary<IntPtr, int> hwndToZOrder = new Dictionary<IntPtr, int>();

            int index = 0;
            bool success = EnumWindows(delegate (IntPtr hwnd, IntPtr _)
            {
                hwndToZOrder[hwnd] = index++;
                return true;
            }, IntPtr.Zero);

            return Comparer<IWindow>.Create((x, y) =>
            {
                if (hwndToZOrder.TryGetValue(x.Handle, out int zorderX))
                {
                    if (hwndToZOrder.TryGetValue(y.Handle, out int zorderY))
                    {
                        return hwndToZOrder[x.Handle] - hwndToZOrder[y.Handle];
                    }
                    return zorderX;
                }
                return 0;
            });
        }

        public ILiveThumbnail CreateLiveThumbnail(IWindow destWindow, IWindow srcWindow)
        {
            return new Win32LiveThumbnail(destWindow, srcWindow);
        }

        public void RefreshConfiguration()
        {
            CheckVirtualDesktops();
            CheckVisibilityChanges();
            OnSettingChange();
            OnDisplayChange();

            foreach (var window in GetVisibleWindowList())
            {
                window.CheckChanges();
            }
        }

        public IWindow UnsafeCreateFromHandle(IntPtr windowHandle)
        {
            return new Win32Window(this, windowHandle);
        }

        public void Dispose()
        {
            KillTimer(IntPtr.Zero, m_hTimer);
            m_isShuttingDown = true;
            m_winEventHook?.Dispose();
            m_eventLoopThread?.Join();
        }

        internal IWindow UnsafeGetWindow(IntPtr hwnd)
        {
            Win32Window window;
            lock (m_windowList)
            {
                window = m_windowList.FirstOrDefault(x => x.Handle == hwnd);
            }
            if (window?.IsTopLevelVisible == true)
            {
                return window;
            }
            return null;
        }

        [Obsolete]
        private void EventLoop()
        {
            m_msgWnd = CreateWindowEx(
                ExtendedWindowStyles.WS_EX_NOACTIVATE,
                "STATIC",
                "WinManMessageReceiver",
                WindowStyles.WS_DISABLED,
                0, 0, 0, 0,
                IntPtr.Zero,
                IntPtr.Zero,
                GetModuleHandle(null),
                IntPtr.Zero);

            m_winEventHook = WinEventHookHelper.CreateGlobalOutOfContextHook(new SortedSet<uint>
            {
                //EVENT_MIN,
                //EVENT_MAX,

                EVENT_OBJECT_CREATE,
                EVENT_OBJECT_DESTROY,

                //EVENT_SYSTEM_MINIMIZESTART,
                //EVENT_SYSTEM_MINIMIZEEND,

                EVENT_SYSTEM_MOVESIZESTART,
                EVENT_SYSTEM_MOVESIZEEND,

                EVENT_SYSTEM_FOREGROUND,
                EVENT_OBJECT_LOCATIONCHANGE,
            }, OnWinEvent);

            m_hTimer = SetTimer(m_msgWnd, IDT_TIMER_WATCH, TIMER_WATCH_INTERVAL, null);
            if (m_hTimer == IntPtr.Zero)
            {
                throw new Win32Exception();
            }

            while (!m_isShuttingDown && GetMessage(out MSG msg, m_msgWnd, 0, 0) > 0)
            {
                WndProc(m_msgWnd, msg.message, msg.wParam, msg.lParam);
            }
        }

        private void WndProc(IntPtr hwnd, uint msg, UIntPtr wParam, IntPtr lParam)
        {
            try
            {
                switch (msg)
                {
                    case WM_TIMER:
                        OnTimerWatch();
                        break;
                    case WM_DISPLAYCHANGE:
                        OnDisplayChange();
                        break;
                    case WM_SETTINGCHANGE:
                        OnSettingChange();
                        break;
                }
            }
            catch (Exception e)
            {
                if (UnhandledException != null)
                {
                    UnhandledException(e);
                    return;
                }
                else
                {
                    throw e;
                }
            }
        }

        private void OnSettingChange()
        {
            m_displayManager.OnSettingChange();
        }

        private void OnDisplayChange()
        {
            m_displayManager.OnDisplayChange();
        }

        private void OnTimerWatch()
        {
            RefreshConfiguration();
        }

        private void OnWinEvent(uint eventType, IntPtr hwnd, int idObject, int idChild, uint idEventThread, uint dwmsEventTime)
        {
            if (idObject != OBJID_WINDOW || idChild != 0)
            {
                return;
            }

            CatchWndProcException(() =>
            {
                Win32Window window;

                switch (eventType)
                {
                    case EVENT_OBJECT_CREATE:
                        OnWindowCreated(hwnd);
                        return;
                    case EVENT_OBJECT_DESTROY:
                        OnWindowDestroyed(hwnd);
                        return;

                    //case EVENT_SYSTEM_MINIMIZEEND:
                    //    if (m_windowSet.TryGetValue(hwnd, out window))
                    //    {
                    //        window.OnStateChanged();
                    //    }
                    //    return;

                    case EVENT_SYSTEM_MOVESIZESTART:
                        if (m_windowSet.TryGetValue(hwnd, out window))
                        {
                            m_windowSet[hwnd].OnMoveSizeStart();
                        }
                        return;
                    case EVENT_SYSTEM_MOVESIZEEND:
                        if (m_windowSet.TryGetValue(hwnd, out window))
                        {
                            m_windowSet[hwnd].OnMoveSizeEnd();
                        }
                        return;

                    case EVENT_SYSTEM_FOREGROUND:
                        OnWindowForeground(hwnd);
                        return;

                    case EVENT_OBJECT_LOCATIONCHANGE:
                        if (m_windowSet.TryGetValue(hwnd, out window))
                        {
                            window.OnPositionChanged();
                        }
                        return;

                    default:
                        return;
                }
            });
        }

        private void CatchWndProcException(Action action)
        {
            try
            {
                action();
            }
            catch (Exception e)
            {
                if (e.IsInvalidWindowHandleException())
                {
                    CatchWndProcException(() => CheckVisibilityChanges());
                    return;
                }
                else if (UnhandledException != null)
                {
                    UnhandledException(e);
                    return;
                }
                else
                {
                    throw e;
                }
            }
        }

        private void CheckVirtualDesktops()
        {
            if (VirtualDesktopManagerLazy is Win32VirtualDesktopManager vdm)
            {
                vdm.CheckVirtualDesktopChanges();
            }
        }

        private void OnWindowCreated(IntPtr hwnd)
        {
            Win32Window window = new Win32Window(this, hwnd);
            m_windowSet[hwnd] = window;
            lock (m_windowList)
            {
                m_windowList.Add(window);
            }

            try
            {
                if (window.IsTopLevelVisible)
                {
                    lock (m_visibleWindows)
                    {
                        m_visibleWindows.Add(window);
                    }
                    try
                    {
                        window.OnAdded();
                    }
                    finally
                    {
                        WindowAdded?.Invoke(window);
                    }
                }
            }
            finally
            {
                if (GetForegroundWindow() == hwnd)
                {
                    OnWindowForeground(hwnd);
                }
            }
        }

        private void OnWindowDestroyed(IntPtr hwnd)
        {
            if (m_windowSet.TryGetValue(hwnd, out Win32Window window))
            {
                m_windowSet.Remove(hwnd);

                lock (m_windowList)
                {
                    m_windowList.Remove(window);
                }

                try
                {
                    bool removed;
                    lock (m_visibleWindows)
                    {
                        removed = m_visibleWindows.Remove(window);
                    }
                    if (removed)
                    {
                        try
                        {
                            window.OnRemoved();
                        }
                        finally
                        {
                            WindowRemoved?.Invoke(window);
                        }
                    }
                }
                finally
                {
                    try
                    {
                        window.OnDestroyed();
                    }
                    finally
                    {
                        WindowDestroyed?.Invoke(window);
                    }
                }
            }
        }

        private void OnWindowForeground(IntPtr hwnd)
        {
            if (m_hwndFocused == hwnd)
            {
                return;
            }

            try
            {
                if (m_windowSet.TryGetValue(m_hwndFocused, out var window))
                {
                    window.OnBackground();
                }
                else
                {
                    IntPtr hwndRoot = GetAncestor(m_hwndFocused, GA.GetRoot);
                    if (m_windowSet.TryGetValue(hwndRoot, out window))
                    {
                        window.OnBackground();
                    }
                }
            }
            finally
            {
                m_hwndFocused = hwnd;

                if (m_windowSet.TryGetValue(hwnd, out var window))
                {
                    window.OnForeground();
                }
                else
                {
                    IntPtr hwndRoot = GetAncestor(hwnd, GA.GetRoot);
                    if (m_windowSet.TryGetValue(hwndRoot, out window))
                    {
                        window.OnForeground();
                    }
                }
            }
        }

        private void CheckVisibilityChanges()
        {
            foreach (var window in GetWindowListSnapshot())
            {
                bool isVisible = window.IsTopLevelVisible;
                bool isInList;
                lock (m_visibleWindows)
                {
                    isInList = m_visibleWindows.Contains(window);
                }

                if (isVisible != isInList)
                {
                    if (isVisible)
                    {
                        lock (m_visibleWindows)
                        {
                            m_visibleWindows.Add(window);
                        }
                        try
                        {
                            window.OnAdded();
                        }
                        finally
                        {
                            WindowAdded?.Invoke(window);
                        }
                    }
                    else
                    {
                        try
                        {
                            window.OnRemoved();
                        }
                        finally
                        {
                            lock (m_visibleWindows)
                            {
                                m_visibleWindows.Remove(window);
                            }
                            WindowRemoved?.Invoke(window);
                        }
                    }
                }
            }
        }

        private void CheckOpen()
        {
            if (m_eventLoopThread == null)
            {
                throw new InvalidOperationException("Call Open() first!");
            }
        }

        private IReadOnlyList<Win32Window> GetWindowListImpl()
        {
            List<Win32Window> windows = new List<Win32Window>();

            bool success = EnumWindows(delegate (IntPtr hwnd, IntPtr _)
            {
                windows.Add(new Win32Window(this, hwnd));
                return true; // Continue
            }, IntPtr.Zero);

            if (!success)
            {
                throw new Win32Exception();
            }

            return windows;
        }

        private Win32Window[] GetWindowListSnapshot()
        {
            Win32Window[] windowListCopy;
            lock (m_windowList)
            {
                windowListCopy = m_windowList.ToArray();
            }
            return windowListCopy;
        }

        private Win32Window[] GetVisibleWindowList()
        {
            Win32Window[] windowListCopy;
            lock (m_visibleWindows)
            {
                windowListCopy = m_visibleWindows.ToArray();
            }
            return windowListCopy;
        }
    }
}
