using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

using static WinMan.Implementation.Win32.NativeMethods;

namespace WinMan.Implementation.Win32
{
    [DebuggerDisplay("Handle = {Handle}, Title = {TitleInternal}")]
    internal class Win32Window : IWindow
    {
        public event WindowUpdatedEventHandler Added;
        public event WindowUpdatedEventHandler Removed;
        public event WindowUpdatedEventHandler Destroyed;

        public event WindowUpdatedEventHandler GotFocus;
        public event WindowUpdatedEventHandler LostFocus;

        public event WindowPositionChangedEventHandler PositionChangeStart;
        public event WindowPositionChangedEventHandler PositionChangeEnd;
        public event WindowPositionChangedEventHandler PositionChanged;
        public event WindowStateChangedEventHandler StateChanged;
        public event WindowUpdatedEventHandler TopmostChanged;

        public IntPtr Handle => m_hwnd;

        public object SyncRoot => m_syncRoot;

        public IWorkspace Workspace
        {
            get
            {
                return m_workspace;
            }
        }

        public string Title => UseDefaults(() => TitleInternal, "[DEAD]");

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
                WindowState GetValue()
                {
                    lock (m_syncRoot)
                    {
                        return m_state;
                    }
                }
                return UseDefaults(GetValue, WindowState.Minimized);
            }
        }

        public bool IsTopmost
        {
            get
            {
                bool GetValue()
                {
                    lock (m_syncRoot)
                    {
                        return m_isTopmost;
                    }
                }
                return UseDefaults(GetValue, false);
            }
        }

        public bool IsFocused => UseDefaults(() => m_isFocused, false);

        public bool CanResize => UseDefaults(() => GetWindowStyles().HasFlag(WindowStyles.WS_SIZEFRAME) && ProbeAccess(), false);

        public bool CanMove => UseDefaults(() => State == WindowState.Restored && ProbeAccess(), false);

        public bool CanReorder => UseDefaults(() => ProbeAccess(), false);

        public bool CanMinimize => UseDefaults(() => GetWindowStyles().HasFlag(WindowStyles.WS_MINIMIZEBOX) && ProbeAccess(), false);

        public bool CanMaximize => UseDefaults(() => GetWindowStyles().HasFlag(WindowStyles.WS_MAXIMIZEBOX) && ProbeAccess(), false);

        public bool CanCreateLiveThumbnail => true;

        public bool CanClose => UseDefaults(() => ProbeAccess(), false);

        public Point? MinSize => UseDefaults(() => GetMinMaxSize().minSize, default);

        public Point? MaxSize => UseDefaults(() => GetMinMaxSize().maxSize, default);

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

        internal bool IsTopLevelVisible => GetIsTopLevelVisible();

        internal string TitleInternal
        {
            get
            {
                try
                {
                    int len = GetWindowTextLengthSafe();
                    return GetWindowTextSafe(len);
                }
                catch
                {
                    return "";
                }
            }
        }

        public Process Process => GetProcess();

        public IWindow NextWindow
        {
            get
            {
                return UseDefaults(() =>
                {
                    IntPtr hwnd = m_hwnd;
                    while ((hwnd = GetWindow(hwnd, GW.HWndNext)) != IntPtr.Zero)
                    {
                        IWindow window = m_workspace.UnsafeGetWindow(hwnd);
                        if (window != null)
                        {
                            return window;
                        }
                    }
                    return null;
                }, null);
            }
        }

        public IWindow PreviousWindow
        {
            get
            {
                return UseDefaults(() =>
                {
                    IntPtr hwnd = m_hwnd;
                    while ((hwnd = GetWindow(hwnd, GW.HWndPrev)) != IntPtr.Zero)
                    {
                        IWindow window = m_workspace.UnsafeGetWindow(hwnd);
                        if (window != null)
                        {
                            return window;
                        }
                    }
                    return null;
                }, null);
            }
        }

        public Rectangle FrameMargins => UseDefaults(() => GetFrameMargins(), default);

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
            m_oldHwndPrev = PreviousWindow?.Handle ?? IntPtr.Zero;
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
            if (!CanResize)
            {
                flags |= SetWindowPosFlags.IgnoreResize;

                Rectangle currentPosition = Position;
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
            catch (Win32Exception e) when (e.IsInvalidHandleException())
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
            catch (Win32Exception e) when (e.IsInvalidHandleException())
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
            lock (m_syncRoot)
            {
                m_initialPosition = m_position;
            }
            PositionChangeStart?.Invoke(this, m_initialPosition);
        }

        internal void OnMoveSizeEnd()
        {
            if (m_hwnd == IntPtr.Zero)
            {
                return;
            }
            UpdateConfiguration();
            try
            {
                PositionChangeEnd?.Invoke(this, m_initialPosition);
            }
            finally
            {
                lock (m_syncRoot)
                {
                    m_initialPosition = m_position;
                }
            }
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
            GotFocus?.Invoke(this);
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
            LostFocus?.Invoke(this);
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
            Added?.Invoke(this);
        }

        internal void OnRemoved()
        {
            Removed?.Invoke(this);
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
                Destroyed?.Invoke(this);
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

        private bool GetIsTopLevelVisible()
        {
            try
            {
                if (!IsWindowVisible(m_hwnd))
                {
                    return false;
                }

                if (m_hwnd != GetAncestor(m_hwnd, GA.GetRoot))
                {
                    return false;
                }

                WindowStyles style = GetWindowStyles();
                if (style.HasFlag(WindowStyles.WS_CHILD))
                {
                    return false;
                }

                ExtendedWindowStyles exStyle = GetExtendedWindowStyles();

                bool isToolWindow = exStyle.HasFlag(ExtendedWindowStyles.WS_EX_TOOLWINDOW);
                if (isToolWindow)
                {
                    return false;
                }

                bool CheckCloaked()
                {
                    uint cloaked = GetDwordDwmWindowAttribute(m_hwnd, DWMWINDOWATTRIBUTE.Cloaked);
                    if (cloaked > 0)
                    {
                        if (cloaked == DWM_CLOAKED_SHELL && m_workspace.VirtualDesktopManager is Win32VirtualDesktopManager vdm)
                        {
                            return vdm.IsNotOnCurrentDesktop(this);
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
                    return CheckCloaked() && GetWindowTextLength(m_hwnd) != 0;
                }

                bool hasEdge = exStyle.HasFlag(ExtendedWindowStyles.WS_EX_WINDOWEDGE);
                bool isTopmostOnly = exStyle == ExtendedWindowStyles.WS_EX_TOPMOST;

                if (hasEdge || isTopmostOnly || exStyle == 0)
                {
                    return CheckCloaked() && GetWindowTextLength(m_hwnd) != 0;
                }

                bool isAcceptFiles = exStyle.HasFlag(ExtendedWindowStyles.WS_EX_ACCEPTFILES);
                if (isAcceptFiles /* && ShowStyle == ShowMaximized */)
                {
                    return CheckCloaked() && GetWindowTextLength(m_hwnd) != 0;
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

        internal WindowStyles GetWindowStyles()
        {
            return (WindowStyles)GetWindowLong(m_hwnd, WindowLongParam.GWL_STYLE).ToInt64();
        }

        internal ExtendedWindowStyles GetExtendedWindowStyles()
        {
            return (ExtendedWindowStyles)GetWindowLong(m_hwnd, WindowLongParam.GWL_EXSTYLE).ToInt64();
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

        private int GetWindowTextLengthSafe()
        {
            IntPtr length = new IntPtr(1);
            if (IntPtr.Zero == SendMessageTimeout(m_hwnd, WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero, SendMessageTimeoutFlags.SMTO_NORMAL | SendMessageTimeoutFlags.SMTO_ABORTIFHUNG, 3000, out length))
            {
                if (length != IntPtr.Zero)
                {
                    CheckAlive();
                    throw new Win32Exception();
                }
            }
            return length.ToInt32();
        }

        private string GetWindowTextSafe(int length)
        {
            StringBuilder sb = new StringBuilder(length + 1);

            IntPtr copied;
            if (IntPtr.Zero == SendMessageTimeoutText(m_hwnd, WM_GETTEXT, new IntPtr(length + 1), sb, SendMessageTimeoutFlags.SMTO_NORMAL | SendMessageTimeoutFlags.SMTO_ABORTIFHUNG, 3000, out copied))
            {
                CheckAlive();
                throw new Win32Exception();
            }

            if (copied.ToInt32() < length)
            {
                CheckAlive();
                throw new IndexOutOfRangeException("Insufficient buffer size");
            }
            return sb.ToString();
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

        private Process GetProcess()
        {
            CheckAlive();
            _ = GetWindowThreadProcessId(m_hwnd, out int processId);
            return Process.GetProcessById(processId);
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
                isTopmost = GetExtendedWindowStyles().HasFlag(ExtendedWindowStyles.WS_EX_TOPMOST);
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
                PositionChanged?.Invoke(this, oldPosition);
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
                StateChanged?.Invoke(this, oldState);
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
                TopmostChanged?.Invoke(this);
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
