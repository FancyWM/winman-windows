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
        public event EventHandler<WindowChangedEventArgs> WindowAdded;
        public event EventHandler<WindowChangedEventArgs> WindowRemoved;
        public event EventHandler<WindowChangedEventArgs> WindowDestroyed;
        public event EventHandler<WindowChangedEventArgs> WindowManaging;
        public event UnhandledExceptionEventHandler UnhandledException;

        private static readonly IntPtr IDT_TIMER_WATCH = new IntPtr(1);

        private readonly object m_initSyncRoot = new object();

        private IntPtr m_msgWnd;
        private Thread m_eventLoopThread;
        private Deleter m_winEventHook;
        private IntPtr m_hTimer;
        private bool m_isShuttingDown = false;
        private TimeSpan m_watchInterval = TimeSpan.FromMilliseconds(250);

        private HashSet<Win32Window> m_visibleWindows = new HashSet<Win32Window>();
        private List<Win32WindowHandle> m_windowList = new List<Win32WindowHandle>();
        private Dictionary<IntPtr, Win32WindowHandle> m_windowSet = new Dictionary<IntPtr, Win32WindowHandle>();
        private IntPtr m_hwndFocused = IntPtr.Zero;

        private IVirtualDesktopManager m_virtualDesktops = null;

        public bool IsOpen => m_eventLoopThread != null;

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

        public TimeSpan WatchInterval
        {
            get => m_watchInterval;
            set 
            {
                lock (m_initSyncRoot)
                {
                    if (IsOpen)
                    {
                        throw new InvalidOperationException("Set this value before calling Open()!");
                    }
                    m_watchInterval = value;
                }
            }
        }

        public Win32Workspace()
        {
        }

        public void Open()
        {
            lock (m_initSyncRoot)
            {
                if (IsOpen)
                {
                    throw new InvalidOperationException("Workspace has already been Open()ed!");
                }

                m_displayManager = new Win32DisplayManager(this);

                m_eventLoopThread = new Thread(EventLoop);
                m_eventLoopThread.Name = "Win32WorkspaceEventLoop";

                foreach (var window in GetWindowListImpl())
                {
                    m_windowSet.Add(window.Handle, window);
                    m_windowList.Add(window);

                    if (GetVisibility(window))
                    {
                        lock (m_visibleWindows)
                        {
                            m_visibleWindows.Add(window.WindowObject);
                        }

                        try
                        {
                            window.WindowObject.OnAdded();
                        }
                        finally
                        {
                            WindowManaging?.Invoke(this, new WindowChangedEventArgs(window.WindowObject));
                        }
                    }
                }

                m_eventLoopThread.Start();
            }
        }

        public IWindow FindWindow(IntPtr windowHandle)
        {
            CheckOpen();
            lock (m_windowList)
            {
                return m_windowList.FirstOrDefault(x => x.Handle == windowHandle)?.WindowObject;
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
            Win32WindowHandle window;
            lock (m_windowList)
            {
                window = m_windowList.FirstOrDefault(x => x.Handle == hwnd);
            }
            if (window?.WindowObject?.IsTopLevelVisible == true)
            {
                return window.WindowObject;
            }
            return null;
        }

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

            // WndHooks.InstallWindowMessageHook(m_msgWnd, (int)(WM_USER + 100), WndHooks.HookID.WH_CALLWNDPROC, 0, 0);

            m_winEventHook = WinEventHookHelper.CreateGlobalOutOfContextHook(new SortedSet<uint>
            {
                EVENT_OBJECT_CREATE,
                EVENT_OBJECT_DESTROY,

                EVENT_SYSTEM_MOVESIZESTART,
                EVENT_SYSTEM_MOVESIZEEND,

                EVENT_SYSTEM_FOREGROUND,
                EVENT_OBJECT_LOCATIONCHANGE,

                EVENT_SYSTEM_DESKTOPSWITCH,

                EVENT_OBJECT_NAMECHANGE,
            }, OnWinEvent);

            m_hTimer = SetTimer(m_msgWnd, IDT_TIMER_WATCH, (uint)m_watchInterval.TotalMilliseconds, null);
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
                    UnhandledException(this, new UnhandledExceptionEventArgs(e, false));
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
            // Dirty checking is still needed, as some things do not have corresponding events. 
            // For example, virtual desktop addition/removal or windows changing their windowstyles at runtime
            // cannot be observed directly.
            RefreshConfiguration();
        }

        private uint m_lastEventType = EVENT_MIN - 1;

        private void OnWinEvent(uint eventType, IntPtr hwnd, int idObject, int idChild, uint idEventThread, uint dwmsEventTime)
        {
            if (idObject != OBJID_WINDOW || idChild != 0)
            {
                return;
            }

            bool isRepeatedEvent = false;
            if (m_lastEventType == eventType)
            {
                isRepeatedEvent = true;
            }
            m_lastEventType = eventType;

            CatchWndProcException(() =>
            {
                Win32WindowHandle window;

                switch (eventType)
                {
                    case EVENT_OBJECT_CREATE:
                        OnWindowCreated(hwnd);
                        return;
                    case EVENT_OBJECT_DESTROY:
                        OnWindowDestroyed(hwnd);
                        return;

                    case EVENT_SYSTEM_DESKTOPSWITCH:
                        CheckVirtualDesktops();
                        CheckVisibilityChanges();
                        return;

                    case EVENT_OBJECT_NAMECHANGE:
                        if (m_windowSet.TryGetValue(hwnd, out window))
                        {
                            window.WindowObject?.OnTitleChange();
                        }
                        return;

                    case EVENT_SYSTEM_MOVESIZESTART:
                        if (m_windowSet.TryGetValue(hwnd, out window))
                        {
                            window.WindowObject?.OnMoveSizeStart();
                        }
                        return;
                    case EVENT_SYSTEM_MOVESIZEEND:
                        if (m_windowSet.TryGetValue(hwnd, out window))
                        {
                            window.WindowObject?.OnMoveSizeEnd();
                        }
                        return;

                    case EVENT_SYSTEM_FOREGROUND:
                        OnWindowForeground(hwnd);
                        return;

                    case EVENT_OBJECT_LOCATIONCHANGE:
                        if (m_windowSet.TryGetValue(hwnd, out window))
                        {
                            window.WindowObject?.OnPositionChanged();
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
                    UnhandledException(this, new UnhandledExceptionEventArgs(e, false));
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
            Win32WindowHandle window = new Win32WindowHandle(hwnd);

            m_windowSet[hwnd] = window;
            lock (m_windowList)
            {
                m_windowList.Add(window);
            }

            try
            {
                if (GetVisibility(window))
                {
                    lock (m_visibleWindows)
                    {
                        m_visibleWindows.Add(window.WindowObject);
                    }
                    try
                    {
                        window.WindowObject.OnAdded();
                    }
                    finally
                    {
                        WindowAdded?.Invoke(this, new WindowChangedEventArgs(window.WindowObject));
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
            if (m_windowSet.TryGetValue(hwnd, out Win32WindowHandle window))
            {
                m_windowSet.Remove(hwnd);

                lock (m_windowList)
                {
                    m_windowList.Remove(window);
                }

                if (window.WindowObject == null)
                {
                    // Window was never visible, no need to continue.
                    return;
                }

                try
                {
                    bool removed;
                    lock (m_visibleWindows)
                    {
                        removed = m_visibleWindows.Remove(window.WindowObject);
                    }
                    if (removed)
                    {
                        try
                        {
                            window?.WindowObject.OnRemoved();
                        }
                        finally
                        {
                            WindowRemoved?.Invoke(this, new WindowChangedEventArgs(window.WindowObject));
                        }
                    }
                }
                finally
                {
                    try
                    {
                        window?.WindowObject.OnDestroyed();
                    }
                    finally
                    {
                        WindowDestroyed?.Invoke(this, new WindowChangedEventArgs(window.WindowObject));
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
                    window.WindowObject?.OnBackground();
                }
                else
                {
                    IntPtr hwndRoot = GetAncestor(m_hwndFocused, GA.GetRoot);
                    if (m_windowSet.TryGetValue(hwndRoot, out window))
                    {
                        window.WindowObject?.OnBackground();
                    }
                }
            }
            finally
            {
                m_hwndFocused = hwnd;

                if (m_windowSet.TryGetValue(hwnd, out var window))
                {
                    window.WindowObject?.OnForeground();
                }
                else
                {
                    IntPtr hwndRoot = GetAncestor(hwnd, GA.GetRoot);
                    if (m_windowSet.TryGetValue(hwndRoot, out window))
                    {
                        window.WindowObject?.OnForeground();
                    }
                }
            }
        }

        private bool GetVisibility(Win32WindowHandle window)
        {
            if (Win32Window.GetIsTopLevelVisible(this, window.Handle))
            {
                window.EnsureWindowObject(this);
                return true;
            }
            return false;
        }

        private void CheckVisibilityChanges()
        {
            foreach (var window in GetWindowListSnapshot())
            {
                CheckVisibilityChanges(window);
            }
        }

        private void CheckVisibilityChanges(Win32WindowHandle window)
        {
            bool isVisible = GetVisibility(window);
            bool isInList;
            lock (m_visibleWindows)
            {
                isInList = m_visibleWindows.Contains(window.WindowObject);
            }

            if (isVisible != isInList)
            {
                if (isVisible)
                {
                    lock (m_visibleWindows)
                    {
                        m_visibleWindows.Add(window.WindowObject);
                    }
                    try
                    {
                        window.WindowObject.OnAdded();
                    }
                    finally
                    {
                        WindowAdded?.Invoke(this, new WindowChangedEventArgs(window.WindowObject));
                    }
                }
                else
                {
                    try
                    {
                        window.WindowObject.OnRemoved();
                    }
                    finally
                    {
                        lock (m_visibleWindows)
                        {
                            m_visibleWindows.Remove(window.WindowObject);
                        }
                        WindowRemoved?.Invoke(this, new WindowChangedEventArgs(window.WindowObject));
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

        private IReadOnlyList<Win32WindowHandle> GetWindowListImpl()
        {
            List<Win32WindowHandle> windows = new List<Win32WindowHandle>();

            bool success = EnumWindows(delegate (IntPtr hwnd, IntPtr _)
            {
                try
                {
                    windows.Add(new Win32WindowHandle(hwnd));
                }
                catch (InvalidWindowReferenceException)
                {
                    // ignore
                }
                return true; // Continue
            }, IntPtr.Zero);

            if (!success)
            {
                throw new Win32Exception();
            }

            return windows;
        }

        private Win32WindowHandle[] GetWindowListSnapshot()
        {
            Win32WindowHandle[] windowListCopy;
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
