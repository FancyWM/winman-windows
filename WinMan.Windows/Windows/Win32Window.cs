using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

using static WinMan.Windows.NativeMethods;

namespace WinMan.Windows
{
    [DebuggerDisplay("Handle = {Handle}, Title = {Title}")]
    public class Win32Window : IWindow
    {
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
        public string Title => m_title;

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

        public bool CanResize => UseDefaults(() => GetWindowStyles(m_hwnd).HasFlag(WindowStyles.WS_SIZEFRAME) && ProbeAccess(), false);

        public bool CanMove => UseDefaults(() => State == WindowState.Restored && ProbeAccess(), false);

        public bool CanReorder => UseDefaults(() => ProbeAccess(), false);

        public bool CanMinimize => UseDefaults(() => GetWindowStyles(m_hwnd).HasFlag(WindowStyles.WS_MINIMIZEBOX) && ProbeAccess(), false);

        public bool CanMaximize => UseDefaults(() => GetWindowStyles(m_hwnd).HasFlag(WindowStyles.WS_MAXIMIZEBOX) && ProbeAccess(), false);

        public bool CanCreateLiveThumbnail => true;

        public bool CanClose => UseDefaults(() => ProbeAccess(), false);

        public Point? MinSize => UseDefaults(() => !CanResize ? Position.Size : GetMinMaxSize().minSize, default);

        public Point? MaxSize => UseDefaults(() => !CanResize ? Position.Size : GetMinMaxSize().maxSize, default);

        public bool IsAlive
        {
            get
            {
                if (m_isDead)
                {
                    return false;
                }

                m_isDead = !IsWindow(m_hwnd);
                return !m_isDead;
            }
        }

        public Rectangle FrameMargins => UseDefaults(() => GetFrameMargins(), default);

        internal bool IsTopLevelVisible => GetIsTopLevelVisible(m_workspace, m_hwnd);

        private const string InvalidHandleTitle = "[Invalid handle]";

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
        private string m_title;
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

        internal Win32Window(Win32Workspace workspace, IntPtr hwnd)
        {
            m_workspace = workspace;
            m_hwnd = hwnd;
            m_isDead = false;
            m_oldHwndPrev = GetPreviousWindow()?.Handle ?? IntPtr.Zero;
            m_title = GetTitle(throwOnError: true);
        }

        public void Close()
        {
            if (!CloseWindow(m_hwnd))
            {
                CheckAlive();
                throw new Win32Exception();
            }
        }

        public void SetPosition(Rectangle position)
        {
            var flags = SetWindowPosFlags.IgnoreZOrder | SetWindowPosFlags.AsynchronousWindowPosition | SetWindowPosFlags.DoNotActivate;
            Rectangle currentPosition = Position;
            if (!CanResize)
            {
                flags |= SetWindowPosFlags.IgnoreResize;
                position = Rectangle.OffsetAndSize(
                    position.Left,
                    position.Top,
                    currentPosition.Width,
                    currentPosition.Height
                );
            }
            if (!CanMove)
            {
                flags |= SetWindowPosFlags.IgnoreMove;
            }

            if (!SetWindowPos(m_hwnd, IntPtr.Zero, position.Left, position.Top, position.Width, position.Height, flags))
            {
                CheckAlive();
            }
            else
            {
                UpdatePositionAndNotify(position);
            }
        }

        public void SetState(WindowState state)
        {
            CheckAlive();

            try
            {
                SW sw;
                switch (state)
                {
                    case WindowState.Minimized:
                        if (!CanMinimize)
                        {
                            throw new InvalidOperationException("The window does not support minimization");
                        }
                        sw = SW.Minimize;
                        break;
                    case WindowState.Maximized:
                        if (!CanMaximize)
                        {
                            throw new InvalidOperationException("The window does not support maximization");
                        }
                        sw = SW.Maximize;
                        break;
                    case WindowState.Restored:
                        sw = SW.Normal;
                        break;
                    default:
                        throw new InvalidProgramException();
                }

                if (!ShowWindow(m_hwnd, sw))
                {
                    CheckAlive();
                    throw new Win32Exception();
                }

                UpdateStateAndNotify(state);
            }
            catch (Win32Exception)
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
                IntPtr hwndAfter = topmost ? new IntPtr(-1) : new IntPtr(0);
                InsertAfter(Handle, hwndAfter);

                UpdateTopmostAndNotify(topmost);
            }
            catch (Win32Exception e) when (e.IsInvalidWindowHandleException())
            {
                m_isDead = true;
                CheckAlive();
            }
        }

        public void RequestFocus()
        {
            CheckAlive();

            try
            {
                if (!SetForegroundWindow(m_hwnd))
                {
                    throw new Win32Exception();
                }
            }
            catch (Win32Exception e) when (e.IsInvalidWindowHandleException())
            {
                m_isDead = true;
                CheckAlive();
            }
        }

        public void InsertAfter(IWindow other)
        {
            CheckAlive();
            InsertAfter(m_hwnd, other.Handle);
        }

        public void SendToBack()
        {
            CheckAlive();
            InsertAfter(m_hwnd, new IntPtr(1));
        }

        public void BringToFront()
        {
            CheckAlive();
            InsertAfter(m_hwnd, new IntPtr(0));
        }

        public IWindow? GetNextWindow()
        {
            return UseDefaults(() =>
            {
                IntPtr hwnd = m_hwnd;
                while ((hwnd = GetWindow(hwnd, GW.HWndNext)) != IntPtr.Zero)
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
                while ((hwnd = GetWindow(hwnd, GW.HWndPrev)) != IntPtr.Zero)
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
            _ = GetWindowThreadProcessId(m_hwnd, out int processId);
            return Process.GetProcessById(processId);
        }

        public override bool Equals(object obj)
        {
            return obj is Win32Window win && win.m_hwnd == m_hwnd;
        }

        public bool Equals(IWindow other)
        {
            return other is Win32Window win && win.m_hwnd == m_hwnd;
        }

        public override int GetHashCode()
        {
            return m_hwnd.GetHashCode();
        }

        public override string ToString()
        {
            return UseDefaults(() => $"[{m_hwnd}]: \"{Title}\"", "[Invalid window]");
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
            var newTitle = GetTitle(throwOnError: false);
            string oldTitle;
            lock (m_syncRoot)
            {
                oldTitle = m_title;
                m_title = newTitle;
            }

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

        private string GetTitle(bool throwOnError)
        {
            StringBuilder sb = new StringBuilder(128);
            if (IntPtr.Zero == SendMessageTimeoutText(m_hwnd, WM_GETTEXT, new IntPtr(sb.Capacity), sb, SendMessageTimeoutFlags.SMTO_NORMAL | SendMessageTimeoutFlags.SMTO_ABORTIFHUNG, 3000, out _))
            {
                if (throwOnError)
                {
                    CheckAlive();
                }
                return InvalidHandleTitle;
            }
            return sb.ToString();
        }

        private Rectangle GetFrameMargins()
        {
            if (DwmGetWindowAttribute(m_hwnd, DWMWINDOWATTRIBUTE.ExtendedFrameBounds, out RECT frame, Marshal.SizeOf<RECT>()) != 0)
            {
                CheckAlive();
                throw new Win32Exception();
            }

            var rect = GetWindowRectSafe();

            return new Rectangle(
                left: frame.LEFT - rect.Left,
                top: frame.TOP - rect.Top,
                right: rect.Right - frame.RIGHT,
                bottom: rect.Bottom - frame.BOTTOM);
        }

        private bool ProbeAccess()
        {
            _ = GetWindowThreadProcessId(m_hwnd, out int processId);
            IntPtr hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, false, processId);
            if (hProcess == IntPtr.Zero)
            {
                return false;
            }
            CloseHandle(hProcess);
            return true;
        }

        private void InsertAfter(IntPtr hwnd, IntPtr hwndAfter)
        {
            var flags = SetWindowPosFlags.IgnoreMove | SetWindowPosFlags.IgnoreResize | SetWindowPosFlags.AsynchronousWindowPosition | SetWindowPosFlags.DoNotActivate;
            if (!SetWindowPos(hwnd, hwndAfter, 0, 0, 0, 0, flags))
            {
                CheckAlive();
                throw new Win32Exception();
            }
        }

        internal static bool GetIsTopLevelVisible(IWorkspace workspace, IntPtr hwnd)
        {
            try
            {
                if (!IsWindowVisible(hwnd))
                {
                    return false;
                }

                if (hwnd != GetAncestor(hwnd, GA.GetRoot))
                {
                    return false;
                }

                WindowStyles style = GetWindowStyles(hwnd);
                if (style.HasFlag(WindowStyles.WS_CHILD))
                {
                    return false;
                }

                ExtendedWindowStyles exStyle = GetExtendedWindowStyles(hwnd);

                bool isToolWindow = exStyle.HasFlag(ExtendedWindowStyles.WS_EX_TOOLWINDOW);
                if (isToolWindow)
                {
                    return false;
                }

                bool CheckCloaked()
                {
                    uint cloaked = GetDwordDwmWindowAttribute(hwnd, DWMWINDOWATTRIBUTE.Cloaked);
                    if (cloaked > 0)
                    {
                        if (cloaked == DWM_CLOAKED_SHELL && workspace.VirtualDesktopManager is Win32VirtualDesktopManager vdm)
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

                bool isAppWindow = exStyle.HasFlag(ExtendedWindowStyles.WS_EX_APPWINDOW);
                if (isAppWindow)
                {
                    return CheckCloaked() && GetWindowTextLength(hwnd) != 0;
                }

                bool hasEdge = exStyle.HasFlag(ExtendedWindowStyles.WS_EX_WINDOWEDGE);
                bool isTopmostOnly = exStyle == ExtendedWindowStyles.WS_EX_TOPMOST;

                if (hasEdge || isTopmostOnly || exStyle == 0)
                {
                    return CheckCloaked() && GetWindowTextLength(hwnd) != 0;
                }

                bool isAcceptFiles = exStyle.HasFlag(ExtendedWindowStyles.WS_EX_ACCEPTFILES);
                if (isAcceptFiles /* && ShowStyle == ShowMaximized */)
                {
                    return CheckCloaked() && GetWindowTextLength(hwnd) != 0;
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
            UpdateConfiguration();
        }

        internal static uint GetDwordDwmWindowAttribute(IntPtr hwnd, DWMWINDOWATTRIBUTE attr)
        {
            uint value;
            int cbSize = Marshal.SizeOf(typeof(uint));
            DwmGetWindowAttribute(hwnd, attr, out value, cbSize);
            return value;
        }

        internal static WindowStyles GetWindowStyles(IntPtr hwnd)
        {
            return (WindowStyles)GetWindowLong(hwnd, WindowLongParam.GWL_STYLE).ToInt64();
        }

        internal static ExtendedWindowStyles GetExtendedWindowStyles(IntPtr hwnd)
        {
            return (ExtendedWindowStyles)GetWindowLong(hwnd, WindowLongParam.GWL_EXSTYLE).ToInt64();
        }

        internal WINDOWPLACEMENT GetWindowPlacementSafe()
        {
            WINDOWPLACEMENT wplc = default;
            if (!GetWindowPlacement(m_hwnd, ref wplc))
            {
                CheckAlive();
                throw new Win32Exception();
            }

            return wplc;
        }

        private static WindowState FromSW(SW sw)
        {
            if (sw.HasFlag(SW.Maximize))
            {
                return WindowState.Maximized;
            }
            else if (sw.HasFlag(SW.ShowMinimized))
            {
                return WindowState.Minimized;
            }
            return WindowState.Restored;
        }

        private Rectangle GetWindowRectSafe()
        {
            RECT rc = default;
            if (!GetWindowRect(m_hwnd, ref rc))
            {
                CheckAlive();
                throw new Win32Exception();
            }

            return new Rectangle(rc.LEFT, rc.TOP, rc.RIGHT, rc.BOTTOM);
        }

        private void UpdateConfiguration()
        {
            Rectangle placement;
            WindowState state;
            bool isTopmost;
            try
            {
                var wplc = GetWindowPlacementSafe();
                state = FromSW(wplc.showCmd);

                placement = GetWindowRectSafe();
                isTopmost = GetExtendedWindowStyles(m_hwnd).HasFlag(ExtendedWindowStyles.WS_EX_TOPMOST);
                if (state == WindowState.Restored && m_workspace.DisplayManager.Displays.Any(x => x.Bounds == placement))
                {
                    state = WindowState.Maximized;
                }
            }
            catch (InvalidWindowReferenceException)
            {
                m_isDead = true;
                return;
            }

            UpdatePositionAndNotify(placement);
            UpdateStateAndNotify(state);
            UpdateTopmostAndNotify(isTopmost);
        }

        private void UpdatePositionAndNotify(Rectangle newPosition)
        {
            Rectangle oldPosition;
            lock (m_syncRoot)
            {
                oldPosition = m_position;
                m_position = newPosition;
            }

            if (oldPosition != newPosition)
            {
                PositionChanged?.Invoke(this, new WindowPositionChangedEventArgs(this, newPosition, oldPosition));
            }
        }

        private void UpdateStateAndNotify(WindowState newState)
        {
            WindowState oldState;
            lock (m_syncRoot)
            {
                oldState = m_state;
                m_state = newState;
            }

            if (oldState != newState)
            {
                StateChanged?.Invoke(this, new WindowStateChangedEventArgs(this, newState, oldState));
            }
        }

        private void UpdateTopmostAndNotify(bool newIsTopmost)
        {
            bool oldIsTopmost;
            lock (m_syncRoot)
            {
                oldIsTopmost = m_isTopmost;
                m_isTopmost = newIsTopmost;
            }

            if (oldIsTopmost != newIsTopmost)
            {
                TopmostChanged?.Invoke(this, new WindowTopmostChangedEventArgs(this, newIsTopmost));
            }
        }

        private (Point? minSize, Point? maxSize) GetMinMaxSize()
        {
            try
            {
                MINMAXINFO info = new MINMAXINFO
                {
                    ptMinTrackSize = new POINT(0, 0),
                    ptMaxTrackSize = new POINT(int.MaxValue, int.MaxValue),
                };

                if (IntPtr.Zero == SendMessageTimeout(Handle, WM_GETMINMAXINFO, IntPtr.Zero, ref info, SendMessageTimeoutFlags.SMTO_NORMAL | SendMessageTimeoutFlags.SMTO_ABORTIFHUNG, 3000, out var result))
                {
                    throw new Win32Exception();
                }
                if (IntPtr.Zero != result)
                {
                    throw new Win32Exception();
                }

                Point? minSize = null;
                if (info.ptMinTrackSize.X !=0 && info.ptMinTrackSize.Y != 0)
                {
                    minSize = new Point(info.ptMinTrackSize.X, info.ptMinTrackSize.Y);
                }

                Point? maxSize = null;
                if (info.ptMaxTrackSize.X != int.MaxValue && info.ptMaxTrackSize.Y != int.MaxValue)
                {
                    maxSize = new Point(info.ptMaxTrackSize.X, info.ptMaxTrackSize.Y);
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
            catch (Win32Exception e) when (e.IsInvalidWindowHandleException())
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
