using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

using WinMan.Windows.DllImports;
using WinMan.Windows.Utilities;

using static WinMan.Windows.DllImports.NativeMethods;

namespace WinMan.Windows
{
    [DebuggerDisplay("Handle = {Handle}, Title = {Title}")]
    public class Win32Window : IWindow
    {
        private delegate void RaiseEvent();

        public event EventHandler<WindowChangedEventArgs>? Added;
        public event EventHandler<WindowChangedEventArgs>? Removed;
        public event EventHandler<WindowChangedEventArgs>? Destroyed;

        public event EventHandler<WindowFocusChangedEventArgs>? GotFocus;
        public event EventHandler<WindowFocusChangedEventArgs>? LostFocus;

        public event EventHandler<WindowPositionChangedEventArgs>? PositionChangeStart;
        public event EventHandler<WindowPositionChangedEventArgs>? PositionChangeEnd;
        public event EventHandler<WindowPositionChangedEventArgs>? PositionChanged;

        public event EventHandler<WindowStateChangedEventArgs>? StateChanged;
        public event EventHandler<WindowTopmostChangedEventArgs>? TopmostChanged;
        public event EventHandler<WindowTitleChangedEventArgs>? TitleChanged;

        public IntPtr Handle => m_hwnd;

        public object SyncRoot => m_syncRoot;

        public IWorkspace Workspace
        {
            get
            {
                return m_workspace;
            }
        }

        // Reference read is atomic
        public string Title => m_title.Result;

        public Rectangle Position
        {
            get
            {
                Rectangle GetValue()
                {
                    lock (m_syncRoot)
                    {
                        return m_position;
                    }
                }
                return UseDefaults(GetValue, Rectangle.Empty);
            }
        }

        public WindowState State
        {
            get
            {
                // int read is atomic
                return UseDefaults(() => m_state, WindowState.Minimized);
            }
        }

        public bool IsTopmost
        {
            get
            {
                // bool read is atomic
                return UseDefaults(() => m_isTopmost, false);
            }
        }

        public bool IsFocused => UseDefaults(() => m_isFocused, false);

        public bool CanResize => UseDefaults(() => GetStyle().HasFlag(WINDOWS_STYLE.WS_SIZEBOX) && CanAccess(), false);

        public bool CanMove => UseDefaults(() => CanAccess(), false);

        public bool CanReorder => UseDefaults(() => CanAccess(), false);

        public bool CanMinimize => UseDefaults(() => GetStyle().HasFlag(WINDOWS_STYLE.WS_MINIMIZEBOX) && CanAccess(), false);

        public bool CanMaximize => UseDefaults(() => GetStyle().HasFlag(WINDOWS_STYLE.WS_MAXIMIZEBOX) && CanAccess(), false);

        public bool CanCreateLiveThumbnail => true;

        public bool CanClose => UseDefaults(() => CanAccess(), false);

        public Point? MinSize => UseDefaults(() => !CanResize ? Position.Size : GetMinMaxSizeCached().minSize, default);

        public Point? MaxSize => UseDefaults(() => !CanResize ? Position.Size : GetMinMaxSizeCached().maxSize, default);

        public bool IsAlive
        {
            get
            {
                if (m_isDead)
                {
                    return false;
                }

                m_isDead = !IsWindow(new(m_hwnd));
                return !m_isDead;
            }
        }

        public Rectangle FrameMargins => UseDefaults(() => GetFrameMargins(), default);

        public string ClassName => UseDefaults(() => GetClassName(), "");

        internal bool IsTopLevelVisible => GetIsTopLevelVisible(m_workspace, m_hwnd);

        private readonly object m_syncRoot = new object();

        /// <summary>
        /// Reference to the IWorkpsace for this Win32Window.
        /// </summary>
        private readonly Win32Workspace m_workspace;
        /// <summary>
        /// The window handle.
        /// </summary>
        private readonly IntPtr m_hwnd;
        /// <summary>
        /// Window is dead. Handle is invalid.
        /// </summary>
        private volatile bool m_isDead;
        /// <summary>
        /// The title of the window.
        /// </summary>
        private Task<string> m_title;
        /// <summary>
        /// The current (last known) location of the window.
        /// </summary>
        private Rectangle m_position;
        /// <summary>
        /// Location before PositionChangeStart.
        /// </summary>
        private Rectangle m_initialPosition;
        /// <summary>
        /// The current (last known) state of the window.
        /// </summary>
        private WindowState m_state;
        /// <summary>
        /// True if the window is HWND_TOPMOST.
        /// </summary>
        private bool m_isTopmost;
        /// <summary>
        /// True of the window is focused.
        /// </summary>
        private bool m_isFocused;
        /// <summary>
        /// The last known previous window handle.
        /// </summary>
        private IntPtr m_oldHwndPrev;
        /// <summary>
        /// The revision of the window state.
        /// </summary>
        private long m_revision = 0;
        /// <summary>
        /// The cached style.
        /// </summary>
        private Cached<WINDOWS_STYLE> m_cachedStyle;
        /// <summary>
        /// The cached accessibility.
        /// </summary>
        private Cached<bool> m_cachedCanAccess;
        /// <summary>
        /// Cached min max window size.
        /// </summary>
        private Cached<(Point? minSize, Point? maxSize)> m_cachedMinMaxSize;

        internal Win32Window(Win32Workspace workspace, IntPtr hwnd)
        {
            m_workspace = workspace;
            m_hwnd = hwnd;
            m_isDead = false;
            m_oldHwndPrev = GetPreviousWindow()?.Handle ?? IntPtr.Zero;
            m_title = m_workspace.RunInBackgroundAsync(GetTitleWithFallback);
        }

        public void Close()
        {
            unsafe
            {
                var flags = SendMessageTimeout_fuFlags.SMTO_NORMAL | SendMessageTimeout_fuFlags.SMTO_ABORTIFHUNG;
                nuint result = 0;
                if (new LRESULT() == SendMessageTimeout(new(m_hwnd), Constants.WM_CLOSE, new(), new(), flags, 3000, &result))
                {
                    int err = Marshal.GetLastWin32Error();
                    err = err == 0 ? (int)Constants.ERROR_TIMEOUT : err;
                    CheckAlive();
                    throw new Win32Exception(err).WithMessage($"Could not close {this}");
                }
            }
        }

        public void SetPosition(Rectangle position)
        {
            SetPosition(position, redraw: true);
        }

        public void SetPosition(Rectangle position, bool redraw)
        {
            if (m_state != WindowState.Restored)
            {
                throw new InvalidOperationException("Cannot set the position of a window that is not in the restored state!");
            }

            var flags = SetWindowPos_uFlags.SWP_NOZORDER | SetWindowPos_uFlags.SWP_NOACTIVATE | SetWindowPos_uFlags.SWP_ASYNCWINDOWPOS;
            if (!redraw)
            {
                flags |= SetWindowPos_uFlags.SWP_NOREDRAW;
            }
            Rectangle currentPosition = Position;
            if (!CanResize)
            {
                flags |= SetWindowPos_uFlags.SWP_NOSIZE;
                position = Rectangle.OffsetAndSize(
                    position.Left,
                    position.Top,
                    currentPosition.Width,
                    currentPosition.Height
                );
            }
            if (!CanMove)
            {
                flags |= SetWindowPos_uFlags.SWP_NOMOVE;
            }

            if (!SetWindowPos(new(m_hwnd), new(), position.Left, position.Top, position.Width, position.Height, flags))
            {
                CheckAlive();
                // Never throw for other reasons here. (I think it sometimes fails sporadically.)
            }
            else
            {
                UpdatePositionAndNotify(position)?.Invoke();
            }
        }

        public void SetState(WindowState state)
        {
            SetState(state, skipTransition: false);
        }

        public void SetState(WindowState state, bool skipTransition)
        {
            try
            {
                SHOW_WINDOW_CMD sw;
                switch (state)
                {
                    case WindowState.Minimized:
                        if (!CanMinimize)
                        {
                            throw new InvalidOperationException("The window does not support minimization");
                        }
                        sw = SHOW_WINDOW_CMD.SW_SHOWMINNOACTIVE;
                        break;
                    case WindowState.Maximized:
                        if (!CanMaximize)
                        {
                            throw new InvalidOperationException("The window does not support maximization");
                        }
                        sw = SHOW_WINDOW_CMD.SW_MAXIMIZE;
                        break;
                    case WindowState.Restored:
                        sw = SHOW_WINDOW_CMD.SW_SHOWNOACTIVATE;
                        break;
                    default:
                        throw new InvalidProgramException();
                }

                if (skipTransition)
                {
                    var wplc = GetWindowPlacementSafe();
                    wplc.showCmd = sw;
                    if (!SetWindowPlacement(new (m_hwnd), in wplc))
                    {
                        // Returns when failed.
                        CheckAlive();
                    }
                }
                else
                {
                    if (!ShowWindow(new(m_hwnd), sw))
                    {
                        // Returns false when was previously visible or when failed.
                        CheckAlive();
                    }
                }

                UpdateStateAndNotify(state)?.Invoke();
            }
            catch (Win32Exception e) when (!e.IsInvalidWindowHandleException())
            {
                CheckAlive();
                throw;
            }
        }

        public void SetTopmost(bool topmost)
        {
            CheckAlive();

            try
            {
                IntPtr hwndAfter = topmost ? Constants.HWND_TOPMOST : Constants.HWND_TOP;
                InsertAfter(Handle, hwndAfter);

                UpdateTopmostAndNotify(topmost)?.Invoke();
            }
            catch (Win32Exception e) when (e.IsInvalidWindowHandleException())
            {
                m_isDead = true;
                CheckAlive();
            }
        }

        public bool RequestFocus()
        {
            if (!SetForegroundWindow(new HWND(m_hwnd)))
            {
                CheckAlive();
                return false;
            }
            return true;
        }

        public void InsertAfter(IWindow other)
        {
            CheckAlive();
            InsertAfter(m_hwnd, other.Handle);
        }

        public void SendToBack()
        {
            CheckAlive();
            InsertAfter(m_hwnd, Constants.HWND_BOTTOM);
        }

        public void BringToFront()
        {
            CheckAlive();
            InsertAfter(m_hwnd, Constants.HWND_TOP);
        }

        public IWindow? GetNextWindow()
        {
            return UseDefaults(() =>
            {
                IntPtr hwnd = m_hwnd;
                while ((hwnd = GetWindow(new(hwnd), GetWindow_uCmdFlags.GW_HWNDNEXT)) != IntPtr.Zero)
                {
                    IWindow? window = m_workspace.UnsafeGetWindow(hwnd);
                    if (window != null)
                    {
                        return window;
                    }
                }
                return null;
            }, null);
        }

        public IWindow? GetPreviousWindow()
        {
            return UseDefaults(() =>
            {
                IntPtr hwnd = m_hwnd;
                while ((hwnd = GetWindow(new(hwnd), GetWindow_uCmdFlags.GW_HWNDPREV)) != IntPtr.Zero)
                {
                    IWindow? window = m_workspace.UnsafeGetWindow(hwnd);
                    if (window != null)
                    {
                        return window;
                    }
                }
                return null;
            }, null);
        }

        public Process GetProcess()
        {
            CheckAlive();
            uint processId = 0;
            unsafe
            {
                _ = GetWindowThreadProcessId(new(m_hwnd), &processId);
            }
            return Process.GetProcessById((int)processId);
        }

        public void RaisePositionChangeStart()
        {
            PositionChangeStart?.Invoke(this, new WindowPositionChangedEventArgs(this, Position, Position));
            UpdateConfiguration();
        }

        public void RaisePositionChangeEnd()
        {
            UpdateConfiguration();
            PositionChangeEnd?.Invoke(this, new WindowPositionChangedEventArgs(this, Position, Position));
        }

        public override bool Equals(object? obj)
        {
            return obj is Win32Window win && win.m_hwnd == m_hwnd;
        }

        public bool Equals(IWindow? other)
        {
            return other is Win32Window win && win.m_hwnd == m_hwnd;
        }

        public override int GetHashCode()
        {
            return m_hwnd.GetHashCode();
        }

        public override string ToString()
        {
            string className;
            try
            {
                className = GetClassName();
            }
            catch (InvalidWindowReferenceException)
            {
                className = "Invalid handle";
            }
            catch (Win32Exception e)
            {
                className = $"Error code {e.ErrorCode:X8}";
            }
            return $"{{{m_hwnd:X8}: {className}}}";
        }

        internal void OnMoveSizeStart()
        {
            if (m_hwnd == IntPtr.Zero)
            {
                return;
            }

            Rectangle oldPosition, newPosition;
            lock (m_syncRoot)
            {
                oldPosition = m_initialPosition;
                newPosition = m_position;
                m_initialPosition = m_position;
            }

            PositionChangeStart?.Invoke(this, new WindowPositionChangedEventArgs(this, newPosition, oldPosition));
        }

        internal void OnMoveSizeEnd()
        {
            if (m_hwnd == IntPtr.Zero)
            {
                return;
            }
            UpdateConfiguration();

            Rectangle oldPosition, newPosition;
            lock (m_syncRoot)
            {
                oldPosition = m_initialPosition;
                newPosition = m_position;
                m_initialPosition = m_position;
            }

            PositionChangeEnd?.Invoke(this, new WindowPositionChangedEventArgs(this, newPosition, oldPosition));
        }

        internal void OnForeground()
        {
            if (m_hwnd == IntPtr.Zero)
            {
                return;
            }
            lock (m_syncRoot)
            {
                m_isFocused = true;
            }
            GotFocus?.Invoke(this, new WindowFocusChangedEventArgs(this, true));
        }

        internal void OnBackground()
        {
            if (m_hwnd == IntPtr.Zero)
            {
                return;
            }
            lock (m_syncRoot)
            {
                m_isFocused = false;
            }
            LostFocus?.Invoke(this, new WindowFocusChangedEventArgs(this, false));
        }

        internal void OnTitleChange()
        {
            string newTitle;
            try
            {
                newTitle = GetTitle();
            }
            catch (Win32Exception)
            {
                return;
            }
            catch (InvalidWindowReferenceException)
            {
                return;
            }

            Task<string> oldTitleTask;
            var newTitleTask = Task.FromResult(newTitle);

            lock (m_syncRoot)
            {
                oldTitleTask = m_title;
                m_title = newTitleTask;
            }

            var oldTitle = oldTitleTask.Result;
            if (oldTitle != newTitle)
            {
                TitleChanged?.Invoke(this, new WindowTitleChangedEventArgs(this, newTitle, oldTitle));
            }
        }

        internal void OnPositionChanged()
        {
            if (m_hwnd == IntPtr.Zero)
            {
                return;
            }
            if (IsTopLevelVisible)
            {
                UpdateConfiguration();
            }
        }

        internal void OnAdded()
        {
            UpdateConfiguration();
            Added?.Invoke(this, new WindowChangedEventArgs(this));
        }

        internal void OnRemoved()
        {
            Removed?.Invoke(this, new WindowChangedEventArgs(this));
        }

        internal void OnStateChanged()
        {
            if (m_hwnd == IntPtr.Zero)
            {
                return;
            }
            UpdateConfiguration();
        }

        internal void OnDestroyed()
        {
            try
            {
                m_isDead = true;
                Destroyed?.Invoke(this, new WindowChangedEventArgs(this));
            }
            finally
            {
                Destroyed = null;
                Added = null;
                Removed = null;
                StateChanged = null;
                PositionChanged = null;
                PositionChangeStart = null;
                PositionChangeEnd = null;
                TopmostChanged = null;
            }
        }

        private string GetTitleWithFallback()
        {
            try
            {
                return GetTitle();
            }
            catch (Win32Exception e) when (e.IsTimeoutException())
            {
                return GetClassName() + " (Not Responding)";
            }
            catch (Win32Exception)
            {
                return "";
            }
        }

        private string GetTitle()
        {
            unsafe
            {
                char[] buffer = new char[256];
                fixed (char* pBuffer = buffer)
                {
                    var flags = SendMessageTimeout_fuFlags.SMTO_NORMAL | SendMessageTimeout_fuFlags.SMTO_ABORTIFHUNG;
                    nuint result = 0;
                    if (new LRESULT() == SendMessageTimeout(new(m_hwnd), Constants.WM_GETTEXT, new((nuint)buffer.Length), new LPARAM((nint)pBuffer), flags, 3000, &result))
                    {
                        int err = Marshal.GetLastWin32Error();
                        err = err == 0 ? (int)Constants.ERROR_TIMEOUT : err;
                        CheckAlive();
                        throw new Win32Exception(err).WithMessage($"WM_GETTEXT failed for window {this}");
                    }
                    return new string(pBuffer);
                }
            }
        }

        private string GetClassName()
        {
            var t = BlockTimer.Create();
            var className = GetClassNameImpl();
            t.LogIfExceeded(15, className);
            return className;
        }

        private string GetClassNameImpl()
        {
            unsafe
            {
                char[] buffer = new char[256];
                fixed (char* pBuffer = buffer)
                {
                    if (0 == NativeMethods.GetClassName(new(m_hwnd), new(pBuffer), buffer.Length))
                    {
                        int err = Marshal.GetLastWin32Error();
                        CheckAlive();
                        throw new Win32Exception(err).WithMessage("GetClassName failed!");
                    }
                    return new string(pBuffer);
                }
            }
        }

        internal static string GetClassName(IntPtr hwnd)
        {
            unsafe
            {
                char[] buffer = new char[256];
                fixed (char* pBuffer = buffer)
                {
                    if (0 == NativeMethods.GetClassName(new(hwnd), new(pBuffer), buffer.Length))
                    {
                        int err = Marshal.GetLastWin32Error();
                        throw new Win32Exception(err).WithMessage("GetClassName failed!");
                    }
                    return new string(pBuffer);
                }
            }
        }

        private Rectangle GetFrameMargins()
        {
            RECT frame;
            unsafe
            {
                int hr = DwmGetWindowAttribute(new(m_hwnd), (uint)DWMWINDOWATTRIBUTE.DWMWA_EXTENDED_FRAME_BOUNDS, &frame, (uint)sizeof(RECT));
                if (hr != 0)
                {
                    CheckAlive();
                    throw new Win32Exception(hr).WithMessage($"DwmGetWindowAttribute(..., DWMWA_EXTENDED_FRAME_BOUNDS, ...) failed for {this}");
                }
            }

            var rect = GetWindowRectSafe();

            return new Rectangle(
                left: frame.left - rect.Left,
                top: frame.top - rect.Top,
                right: rect.Right - frame.right,
                bottom: rect.Bottom - frame.bottom);
        }

        private bool ProbeAccess()
        {
            uint processId = 0;
            unsafe
            {
                _ = GetWindowThreadProcessId(new(m_hwnd), &processId);
            }
            var hProcess = OpenProcess(PROCESS_ACCESS_RIGHTS.PROCESS_QUERY_INFORMATION, false, processId);
            if (hProcess == IntPtr.Zero)
            {
                return false;
            }
            CloseHandle(hProcess);
            return true;
        }

        private void InsertAfter(IntPtr hwnd, IntPtr hwndAfter)
        {
            var flags = SetWindowPos_uFlags.SWP_NOMOVE | SetWindowPos_uFlags.SWP_NOSIZE | SetWindowPos_uFlags.SWP_ASYNCWINDOWPOS | SetWindowPos_uFlags.SWP_NOACTIVATE;
            if (!SetWindowPos(new(hwnd), new(hwndAfter), 0, 0, 0, 0, flags))
            {
                CheckAlive();
                throw new Win32Exception().WithMessage($"Could not change the Z-order of {this}!");
            }
        }

        internal static bool GetIsLikelyToBeTopLevelVisibleSoon(IWorkspace workspace, IntPtr hwnd)
        {
            if (hwnd != GetAncestor(new(hwnd), GetAncestor_gaFlags.GA_ROOTOWNER))
            {
                return false;
            }

            WINDOWS_STYLE style = GetWINDOWS_STYLE(hwnd);
            if (style.HasFlag(WINDOWS_STYLE.WS_CHILD))
            {
                return false;
            }
            if (style.HasFlag(WINDOWS_STYLE.WS_POPUP))
            {
                return false;
            }
            if (style.HasFlag(WINDOWS_STYLE.WS_DISABLED))
            {
                return false;
            }

            return true;
        }

        internal static bool GetIsTopLevelVisible(IWorkspace workspace, IntPtr hwnd)
        {
            try
            {
                if (!IsWindowVisible(new(hwnd)))
                {
                    return false;
                }

                if (hwnd != GetAncestor(new(hwnd), GetAncestor_gaFlags.GA_ROOT))
                {
                    return false;
                }

                WINDOWS_STYLE style = GetWINDOWS_STYLE(hwnd);
                if (style.HasFlag(WINDOWS_STYLE.WS_CHILD))
                {
                    return false;
                }

                WINDOWS_EX_STYLE exStyle = GetWINDOWS_EX_STYLE(hwnd);

                bool isToolWindow = exStyle.HasFlag(WINDOWS_EX_STYLE.WS_EX_TOOLWINDOW);
                if (isToolWindow)
                {
                    return false;
                }

                bool CheckCloaked()
                {
                    uint cloaked = GetDwordDwmWindowAttribute(hwnd, DWMWINDOWATTRIBUTE.DWMWA_CLOAKED);
                    if (cloaked > 0)
                    {
                        // TODO: Windows 11
                        if (cloaked == Constants.DWM_CLOAKED_SHELL && workspace.VirtualDesktopManager is IWin32VirtualDesktopManagerInternal vdm)
                        {
                            return vdm.IsNotOnCurrentDesktop(hwnd);
                        }
                        else
                        {
                            return false;
                        }
                    }
                    return true;
                }

                bool isAppWindow = exStyle.HasFlag(WINDOWS_EX_STYLE.WS_EX_APPWINDOW);
                if (isAppWindow)
                {
                    return CheckCloaked() && GetWindowTextLength(new(hwnd)) != 0;
                }

                bool hasEdge = exStyle.HasFlag(WINDOWS_EX_STYLE.WS_EX_WINDOWEDGE);
                bool isTopmostOnly = exStyle == WINDOWS_EX_STYLE.WS_EX_TOPMOST;

                if (hasEdge || isTopmostOnly || exStyle == 0)
                {
                    return CheckCloaked() && GetWindowTextLength(new(hwnd)) != 0;
                }

                bool isAcceptFiles = exStyle.HasFlag(WINDOWS_EX_STYLE.WS_EX_ACCEPTFILES);
                if (isAcceptFiles /* && ShowStyle == ShowMaximized */)
                {
                    return CheckCloaked() && GetWindowTextLength(new(hwnd)) != 0;
                }

                return false;
            }
            catch (Win32Exception)
            {
                return false;
            }
        }

        internal void CheckChanges()
        {
            var t = BlockTimer.Create();
            UpdateConfiguration();
            t.LogIfExceeded(15, Handle.ToString());
        }

        internal static uint GetDwordDwmWindowAttribute(IntPtr hwnd, DWMWINDOWATTRIBUTE attr)
        {
            uint value;
            unsafe
            {
                DwmGetWindowAttribute(new(hwnd), (uint)attr, &value, sizeof(uint));
            }
            return value;
        }

        internal static WINDOWS_STYLE GetWINDOWS_STYLE(IntPtr hwnd)
        {
            return (WINDOWS_STYLE)GetWindowLong(new HWND(hwnd), GetWindowLongPtr_nIndex.GWL_STYLE);
        }

        internal static WINDOWS_EX_STYLE GetWINDOWS_EX_STYLE(IntPtr hwnd)
        {
            return (WINDOWS_EX_STYLE)GetWindowLong(new HWND(hwnd), GetWindowLongPtr_nIndex.GWL_EXSTYLE);
        }

        internal WINDOWPLACEMENT GetWindowPlacementSafe()
        {
            WINDOWPLACEMENT wplc = default;
            if (!GetWindowPlacement(new HWND(m_hwnd), ref wplc))
            {
                CheckAlive();
                throw new Win32Exception().WithMessage($"Could not read the placement of {this}");
            }

            return wplc;
        }

        private static WindowState FromSW(SHOW_WINDOW_CMD sw)
        {
            if (sw.HasFlag(SHOW_WINDOW_CMD.SW_MAXIMIZE))
            {
                return WindowState.Maximized;
            }
            else if (sw.HasFlag(SHOW_WINDOW_CMD.SW_SHOWMINIMIZED))
            {
                return WindowState.Minimized;
            }
            return WindowState.Restored;
        }

        private Rectangle GetWindowRectSafe()
        {
            if (!GetWindowRect(new HWND(m_hwnd), out RECT rc))
            {
                CheckAlive();
                throw new Win32Exception().WithMessage($"Could not get the bounding rectangle of {this}!");
            }

            return new Rectangle(rc.left, rc.top, rc.right, rc.bottom);
        }

        private void UpdateConfiguration()
        {
            Interlocked.Increment(ref m_revision);

            Rectangle placement;
            WindowState state;
            bool isTopmost;
            try
            {
                var wplc = GetWindowPlacementSafe();
                state = FromSW(wplc.showCmd);

                placement = GetWindowRectSafe();
                isTopmost = GetWINDOWS_EX_STYLE(m_hwnd).HasFlag(WINDOWS_EX_STYLE.WS_EX_TOPMOST);
                // E.g. Chrome fullscreen when switching desktops
                if (state == WindowState.Restored && m_workspace.DisplayManager.Displays.Any(x => CloseMatch(x.Bounds, placement, threshold: 2)))
                {
                    state = WindowState.Maximized;
                }
            }
            catch (InvalidWindowReferenceException)
            {
                m_isDead = true;
                return;
            }

            // Carry all of the updates, then notify!
            var raiseTopmostChanged = UpdateTopmostAndNotify(isTopmost);
            var raiseStateChanged = UpdateStateAndNotify(state);
            var raisePositionChanged = UpdatePositionAndNotify(placement);

            raiseTopmostChanged?.Invoke();
            raiseStateChanged?.Invoke();
            raisePositionChanged?.Invoke();
        }

        private bool CloseMatch(Rectangle rectA, Rectangle rectB, int threshold)
        {
            return Math.Abs(rectA.Left - rectB.Left) <= threshold
                && Math.Abs(rectA.Top - rectB.Top) <= threshold
                && Math.Abs(rectA.Right -  rectB.Right) <= threshold
                && Math.Abs(rectA.Bottom - rectB.Bottom) <= threshold;
        }

        private RaiseEvent? UpdatePositionAndNotify(Rectangle newPosition)
        {
            Rectangle oldPosition;
            lock (m_syncRoot)
            {
                oldPosition = m_position;
                m_position = newPosition;
            }

            if (oldPosition != newPosition)
            {
                return () => PositionChanged?.Invoke(this, new WindowPositionChangedEventArgs(this, newPosition, oldPosition));
            }
            return null;
        }

        private RaiseEvent? UpdateStateAndNotify(WindowState newState)
        {
            WindowState oldState;
            lock (m_syncRoot)
            {
                oldState = m_state;
                m_state = newState;
            }

            if (oldState != newState)
            {
                return () => StateChanged?.Invoke(this, new WindowStateChangedEventArgs(this, newState, oldState));
            }
            return null;
        }

        private RaiseEvent? UpdateTopmostAndNotify(bool newIsTopmost)
        {
            bool oldIsTopmost;
            lock (m_syncRoot)
            {
                oldIsTopmost = m_isTopmost;
                m_isTopmost = newIsTopmost;
            }

            if (oldIsTopmost != newIsTopmost)
            {
                return () => TopmostChanged?.Invoke(this, new WindowTopmostChangedEventArgs(this, newIsTopmost));
            }
            return null;
        }

        private (Point? minSize, Point? maxSize) GetMinMaxSize()
        {
            try
            {
                MINMAXINFO info = new MINMAXINFO
                {
                    ptMinTrackSize = new POINT { x = 0, y = 0 },
                    ptMaxTrackSize = new POINT { x = int.MaxValue, y = int.MaxValue },
                };

                var flags = SendMessageTimeout_fuFlags.SMTO_NORMAL | SendMessageTimeout_fuFlags.SMTO_ABORTIFHUNG;
                unsafe
                {
                    nuint result = 0;
                    if (new LRESULT() == SendMessageTimeout(new(Handle), Constants.WM_GETMINMAXINFO, new(), new(new IntPtr(&info)), flags, 3000, &result))
                    {
                        int err = Marshal.GetLastWin32Error();
                        err = err == 0 ? (int)Constants.ERROR_TIMEOUT : err;
                        throw new Win32Exception(err).WithMessage($"WM_GETMINMAXINFO failed for window {this}!");
                    }
                    // If an application processes WM_GETMINMAXINFO if should return 0.
                    if (0 != result)
                    {
                        return (null, null);
                    }
                }

                Point? minSize = null;
                if (info.ptMinTrackSize.x !=0 && info.ptMinTrackSize.y != 0)
                {
                    minSize = new Point(info.ptMinTrackSize.x, info.ptMinTrackSize.y);
                }

                Point? maxSize = null;
                if (info.ptMaxTrackSize.x != int.MaxValue && info.ptMaxTrackSize.y != int.MaxValue)
                {
                    maxSize = new Point(info.ptMaxTrackSize.x, info.ptMaxTrackSize.y);
                }

                return (minSize, maxSize);
            }
            catch (Win32Exception)
            {
                CheckAlive();
                throw;
            }
        }

        private void CheckAlive()
        {
            if (!IsAlive)
            {
                throw new InvalidWindowReferenceException(m_hwnd);
            }
        }

        private (Point? minSize, Point? maxSize) GetMinMaxSizeCached()
        {
            bool outdated;
            lock (m_syncRoot)
            {
                outdated = m_cachedMinMaxSize.Revision != m_revision;
                if (!outdated)
                {
                    return m_cachedMinMaxSize.Value;
                }
            }

            var fresh = GetMinMaxSize();
            lock (m_syncRoot)
            {
                m_cachedMinMaxSize.Revision = m_revision;
                m_cachedMinMaxSize.Value = fresh;
            }
            return fresh;
        }

        private WINDOWS_STYLE GetStyle()
        {
            lock (m_syncRoot)
            {
                if (m_cachedStyle.Revision != m_revision)
                {
                    m_cachedStyle.Revision = m_revision;
                    m_cachedStyle.Value = GetWINDOWS_STYLE(m_hwnd);
                }
                return m_cachedStyle.Value;
            }
        }

        private bool CanAccess()
        {
            lock (m_syncRoot)
            {
                if (m_cachedCanAccess.Revision != m_revision)
                {
                    m_cachedCanAccess.Revision = m_revision;
                    m_cachedCanAccess.Value = ProbeAccess();
                }
                return m_cachedCanAccess.Value;
            }
        }

        private T UseDefaults<T>(Func<T> func, T defaultValue)
        {
            if (!IsAlive)
            {
                return defaultValue;
            }

            try
            {
                return func();
            }
            catch (InvalidWindowReferenceException)
            {
                return defaultValue;
            }
            catch (Win32Exception e) when (e.IsInvalidWindowHandleException() || e.IsAccessDeniedException() || e.IsTimeoutException())
            {
                return defaultValue;
            }
            catch (Exception)
            {
                if (!IsAlive)
                {
                    return defaultValue;
                }
                throw;
            }
        }
    }
}
