using System;
using System.Runtime.InteropServices;
using System.Text;

namespace WinMan.Implementation.Win32
{
    internal static class NativeMethods
    {
        [StructLayout(LayoutKind.Sequential)]
        internal struct POINT
        {
            public int X;
            public int Y;

            public POINT(int x, int y)
            {
                X = x;
                Y = y;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct MSG
        {
            public IntPtr hwnd;
            public uint message;
            public UIntPtr wParam;
            public IntPtr lParam;
            public int time;
            public POINT pt;
            public int lPrivate;
        }


        [DllImport("user32.dll", SetLastError = true)]
        internal static extern sbyte GetMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);
        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool TranslateMessage([In] ref MSG lpMsg);
        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr DispatchMessage([In] ref MSG lpmsg);

        internal delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);


        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool GetCursorPos(out POINT lpPoint);

        /// <summary>
        ///
        /// </summary>
        /// <param name="hwnd"></param
        /// <param name="lpString"></param>
        /// <param name="maxCount"></param>
        /// <returns>A value of 0 might indicate failure, check GetLastError</returns>
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern int GetWindowText(IntPtr hwnd, StringBuilder lpString, int maxCount);

        /// <summary>
        ///
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns>A value of 0 might indicate failure, check GetLastError</returns>
        [DllImport("user32.dll", SetLastError = true)]
        internal static extern int GetWindowTextLength(IntPtr hwnd);

        internal const uint WINEVENT_OUTOFCONTEXT = 0x0000; // Events are ASYNC

        internal const uint WINEVENT_SKIPOWNTHREAD = 0x0001; // Don't call back for events on installer's thread

        internal const uint WINEVENT_SKIPOWNPROCESS = 0x0002; // Don't call back for events on installer's process

        internal const uint WINEVENT_INCONTEXT = 0x0004; // Events are SYNC, this causes your dll to be injected into every process

        internal const uint EVENT_MIN = 0x00000001;

        internal const uint EVENT_MAX = 0x7FFFFFFF;

        internal const uint EVENT_SYSTEM_SOUND = 0x0001;

        internal const uint EVENT_SYSTEM_ALERT = 0x0002;

        internal const uint EVENT_SYSTEM_FOREGROUND = 0x0003;

        internal const uint EVENT_SYSTEM_MENUSTART = 0x0004;

        internal const uint EVENT_SYSTEM_MENUEND = 0x0005;

        internal const uint EVENT_SYSTEM_MENUPOPUPSTART = 0x0006;

        internal const uint EVENT_SYSTEM_MENUPOPUPEND = 0x0007;

        internal const uint EVENT_SYSTEM_CAPTURESTART = 0x0008;

        internal const uint EVENT_SYSTEM_CAPTUREEND = 0x0009;

        internal const uint EVENT_SYSTEM_MOVESIZESTART = 0x000A;

        internal const uint EVENT_SYSTEM_MOVESIZEEND = 0x000B;

        internal const uint EVENT_SYSTEM_CONTEXTHELPSTART = 0x000C;

        internal const uint EVENT_SYSTEM_CONTEXTHELPEND = 0x000D;

        internal const uint EVENT_SYSTEM_DRAGDROPSTART = 0x000E;

        internal const uint EVENT_SYSTEM_DRAGDROPEND = 0x000F;

        internal const uint EVENT_SYSTEM_DIALOGSTART = 0x0010;

        internal const uint EVENT_SYSTEM_DIALOGEND = 0x0011;

        internal const uint EVENT_SYSTEM_SCROLLINGSTART = 0x0012;

        internal const uint EVENT_SYSTEM_SCROLLINGEND = 0x0013;

        internal const uint EVENT_SYSTEM_SWITCHSTART = 0x0014;

        internal const uint EVENT_SYSTEM_SWITCHEND = 0x0015;

        internal const uint EVENT_SYSTEM_MINIMIZESTART = 0x0016;

        internal const uint EVENT_SYSTEM_MINIMIZEEND = 0x0017;

        internal const uint EVENT_SYSTEM_DESKTOPSWITCH = 0x0020;

        internal const uint EVENT_SYSTEM_END = 0x00FF;

        internal const uint EVENT_OEM_DEFINED_START = 0x0101;

        internal const uint EVENT_OEM_DEFINED_END = 0x01FF;

        internal const uint EVENT_UIA_EVENTID_START = 0x4E00;

        internal const uint EVENT_UIA_EVENTID_END = 0x4EFF;

        internal const uint EVENT_UIA_PROPID_START = 0x7500;

        internal const uint EVENT_UIA_PROPID_END = 0x75FF;

        internal const uint EVENT_CONSOLE_CARET = 0x4001;

        internal const uint EVENT_CONSOLE_UPDATE_REGION = 0x4002;

        internal const uint EVENT_CONSOLE_UPDATE_SIMPLE = 0x4003;

        internal const uint EVENT_CONSOLE_UPDATE_SCROLL = 0x4004;

        internal const uint EVENT_CONSOLE_LAYOUT = 0x4005;

        internal const uint EVENT_CONSOLE_START_APPLICATION = 0x4006;

        internal const uint EVENT_CONSOLE_END_APPLICATION = 0x4007;

        internal const uint EVENT_CONSOLE_END = 0x40FF;

        internal const uint EVENT_OBJECT_CREATE = 0x8000; // hwnd ID idChild is created item

        internal const uint EVENT_OBJECT_DESTROY = 0x8001; // hwnd ID idChild is destroyed item

        internal const uint EVENT_OBJECT_SHOW = 0x8002; // hwnd ID idChild is shown item

        internal const uint EVENT_OBJECT_HIDE = 0x8003; // hwnd ID idChild is hidden item

        internal const uint EVENT_OBJECT_REORDER = 0x8004; // hwnd ID idChild is parent of zordering children

        internal const uint EVENT_OBJECT_FOCUS = 0x8005; // hwnd ID idChild is focused item

        internal const uint EVENT_OBJECT_SELECTION = 0x8006; // hwnd ID idChild is selected item (if only one), or idChild is OBJID_WINDOW if complex

        internal const uint EVENT_OBJECT_SELECTIONADD = 0x8007; // hwnd ID idChild is item added

        internal const uint EVENT_OBJECT_SELECTIONREMOVE = 0x8008; // hwnd ID idChild is item removed

        internal const uint EVENT_OBJECT_SELECTIONWITHIN = 0x8009; // hwnd ID idChild is parent of changed selected items

        internal const uint EVENT_OBJECT_STATECHANGE = 0x800A; // hwnd ID idChild is item w/ state change

        internal const uint EVENT_OBJECT_LOCATIONCHANGE = 0x800B; // hwnd ID idChild is moved/sized item

        internal const uint EVENT_OBJECT_NAMECHANGE = 0x800C; // hwnd ID idChild is item w/ name change

        internal const uint EVENT_OBJECT_DESCRIPTIONCHANGE = 0x800D; // hwnd ID idChild is item w/ desc change

        internal const uint EVENT_OBJECT_VALUECHANGE = 0x800E; // hwnd ID idChild is item w/ value change

        internal const uint EVENT_OBJECT_PARENTCHANGE = 0x800F; // hwnd ID idChild is item w/ new parent

        internal const uint EVENT_OBJECT_HELPCHANGE = 0x8010; // hwnd ID idChild is item w/ help change

        internal const uint EVENT_OBJECT_DEFACTIONCHANGE = 0x8011; // hwnd ID idChild is item w/ def action change

        internal const uint EVENT_OBJECT_ACCELERATORCHANGE = 0x8012; // hwnd ID idChild is item w/ keybd accel change

        internal const uint EVENT_OBJECT_INVOKED = 0x8013; // hwnd ID idChild is item invoked

        internal const uint EVENT_OBJECT_TEXTSELECTIONCHANGED = 0x8014; // hwnd ID idChild is item w? test selection change

        internal const uint EVENT_OBJECT_CONTENTSCROLLED = 0x8015;

        internal const uint EVENT_SYSTEM_ARRANGMENTPREVIEW = 0x8016;

        internal const uint EVENT_OBJECT_END = 0x80FF;

        internal const uint EVENT_AIA_START = 0xA000;

        internal const uint EVENT_AIA_END = 0xAFFF;

        internal const int OBJID_WINDOW = 0;
        internal const int OBJID_SYSMENU = -1;
        internal const int OBJID_TITLEBAR = -2;
        internal const int OBJID_MENU = -3;
        internal const int OBJID_CLIENT = -4;
        internal const int OBJID_VSCROLL = -5;
        internal const int OBJID_HSCROLL = -6;
        internal const int OBJID_SIZEGRIP = -7;
        internal const int OBJID_CARET = -8;
        internal const int OBJID_CURSOR = -9;
        internal const int OBJID_ALERT = -10;
        internal const int OBJID_SOUND = -11;
        internal const int OBJID_QUERYCLASSNAMEIDX = -12;
        internal const int OBJID_NATIVEOM = -13;

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        internal delegate void WinEventProc(IntPtr hWinEventHook, uint eventType,
            IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr
            hmodWinEventProc, WinEventProc lpfnWinEventProc, uint idProcess,
            uint idThread, uint dwFlags);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool UnhookWinEvent(IntPtr hHook);

        [Flags()]
        internal enum WindowStyles : uint
        {
            /// <summary>The window has a thin-line border.</summary>
            WS_BORDER = 0x800000,

            /// <summary>The window has a title bar (includes the WS_BORDER style).</summary>
            WS_CAPTION = 0xc00000,

            /// <summary>The window is a child window. A window with this style cannot have a menu bar. This style cannot be used with the WS_POPUP style.</summary>
            WS_CHILD = 0x40000000,

            /// <summary>Excludes the area occupied by child windows when drawing occurs within the parent window. This style is used when creating the parent window.</summary>
            WS_CLIPCHILDREN = 0x2000000,

            /// <summary>
            /// Clips child windows relative to each other; that is, when a particular child window receives a WM_PAINT message, the WS_CLIPSIBLINGS style clips all other overlapping child windows out of the region of the child window to be updated.
            /// If WS_CLIPSIBLINGS is not specified and child windows overlap, it is possible, when drawing within the client area of a child window, to draw within the client area of a neighboring child window.
            /// </summary>
            WS_CLIPSIBLINGS = 0x4000000,

            /// <summary>The window is initially disabled. A disabled window cannot receive input from the user. To change this after a window has been created, use the EnableWindow function.</summary>
            WS_DISABLED = 0x8000000,

            /// <summary>The window has a border of a style typically used with dialog boxes. A window with this style cannot have a title bar.</summary>
            WS_DLGFRAME = 0x400000,

            /// <summary>
            /// The window is the first control of a group of controls. The group consists of this first control and all controls defined after it, up to the next control with the WS_GROUP style.
            /// The first control in each group usually has the WS_TABSTOP style so that the user can move from group to group. The user can subsequently change the keyboard focus from one control in the group to the next control in the group by using the direction keys.
            /// You can turn this style on and off to change dialog box navigation. To change this style after a window has been created, use the SetWindowLong function.
            /// </summary>
            WS_GROUP = 0x20000,

            /// <summary>The window has a horizontal scroll bar.</summary>
            WS_HSCROLL = 0x100000,

            /// <summary>The window is initially maximized.</summary>
            WS_MAXIMIZE = 0x1000000,

            /// <summary>The window has a maximize button. Cannot be combined with the WS_EX_CONTEXTHELP style. The WS_SYSMENU style must also be specified.</summary>
            WS_MAXIMIZEBOX = 0x10000,

            /// <summary>The window is initially minimized.</summary>
            WS_MINIMIZE = 0x20000000,

            /// <summary>The window has a minimize button. Cannot be combined with the WS_EX_CONTEXTHELP style. The WS_SYSMENU style must also be specified.</summary>
            WS_MINIMIZEBOX = 0x20000,

            /// <summary>The window is an overlapped window. An overlapped window has a title bar and a border.</summary>
            WS_OVERLAPPED = 0x0,

            /// <summary>The window is an overlapped window.</summary>
            //WS_OVERLAPPEDWINDOW = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_SIZEFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX,

            /// <summary>The window is a pop-up window. This style cannot be used with the WS_CHILD style.</summary>
            WS_POPUP = 0x80000000u,

            /// <summary>The window is a pop-up window. The WS_CAPTION and WS_POPUPWINDOW styles must be combined to make the window menu visible.</summary>
            //WS_POPUPWINDOW = WS_POPUP | WS_BORDER | WS_SYSMENU,

            /// <summary>The window has a sizing border.</summary>
            WS_SIZEFRAME = 0x40000,

            /// <summary>The window has a window menu on its title bar. The WS_CAPTION style must also be specified.</summary>
            WS_SYSMENU = 0x80000,

            /// <summary>
            /// The window is a control that can receive the keyboard focus when the user presses the TAB key.
            /// Pressing the TAB key changes the keyboard focus to the next control with the WS_TABSTOP style.
            /// You can turn this style on and off to change dialog box navigation. To change this style after a window has been created, use the SetWindowLong function.
            /// For user-created windows and modeless dialogs to work with tab stops, alter the message loop to call the IsDialogMessage function.
            /// </summary>
            WS_TABSTOP = 0x10000,

            /// <summary>The window is initially visible. This style can be turned on and off by using the ShowWindow or SetWindowPos function.</summary>
            WS_VISIBLE = 0x10000000,

            /// <summary>The window has a vertical scroll bar.</summary>
            WS_VSCROLL = 0x200000
        }
        internal enum WindowLongParam
        {
            /// <summary>Sets a new address for the window procedure.</summary>
            /// <remarks>You cannot change this attribute if the window does not belong to the same process as the calling thread.</remarks>
            GWL_WNDPROC = -4,

            /// <summary>Sets a new application instance handle.</summary>
            GWLP_HINSTANCE = -6,

            GWLP_HWNDPARENT = -8,

            /// <summary>Sets a new identifier of the child window.</summary>
            /// <remarks>The window cannot be a top-level window.</remarks>
            GWL_ID = -12,

            /// <summary>Sets a new window style.</summary>
            GWL_STYLE = -16,

            /// <summary>Sets a new extended window style.</summary>
            /// <remarks>See <see cref="ExWindowStyles"/>.</remarks>
            GWL_EXSTYLE = -20,

            /// <summary>Sets the user data associated with the window.</summary>
            /// <remarks>This data is intended for use by the application that created the window. Its value is initially zero.</remarks>
            GWL_USERDATA = -21,

            /// <summary>Sets the return value of a message processed in the dialog box procedure.</summary>
            /// <remarks>Only applies to dialog boxes.</remarks>
            DWLP_MSGRESULT = 0,

            /// <summary>Sets new extra information that is private to the application, such as handles or pointers.</summary>
            /// <remarks>Only applies to dialog boxes.</remarks>
            DWLP_USER = 8,

            /// <summary>Sets the new address of the dialog box procedure.</summary>
            /// <remarks>Only applies to dialog boxes.</remarks>
            DWLP_DLGPROC = 4
        }


        [DllImport("user32.dll", EntryPoint = "GetWindowLong", SetLastError = true)]
        private static extern IntPtr GetWindowLong32(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "GetWindowLong", SetLastError = true)]
        private static extern IntPtr SetWindowLong32(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        internal static IntPtr GetWindowLong(IntPtr hWnd, WindowLongParam nIndex)
        {
            if (IntPtr.Size == 4)
            {
                return GetWindowLong32(hWnd, (int)nIndex);
            }
            return GetWindowLongPtr64(hWnd, (int)nIndex);
        }

        internal static IntPtr SetWindowLong(IntPtr hWnd, WindowLongParam nIndex, IntPtr dwNewLong)
        {
            if (IntPtr.Size == 4)
            {
                return SetWindowLong32(hWnd, (int)nIndex, dwNewLong);
            }
            return SetWindowLongPtr64(hWnd, (int)nIndex, dwNewLong);
        }

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr GetWindow(IntPtr hWnd, GW uCmd);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool IsWindow(IntPtr hWnd);

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern void SetLastError(uint dwErrorCode);


        internal enum DWMWINDOWATTRIBUTE : int
        {
            NCRenderingEnabled = 1,
            NCRenderingPolicy,
            TransitionsForceDisabled,
            AllowNCPaint,
            CaptionButtonBounds,
            NonClientRtlLayout,
            ForceIconicRepresentation,
            Flip3DPolicy,
            ExtendedFrameBounds,
            HasIconicBitmap,
            DisallowPeek,
            ExcludedFromPeek,
            Cloak,
            Cloaked,
            FreezeRepresentation
        }

        internal const uint DWM_CLOAKED_APP = 0x0000001;
        internal const uint DWM_CLOAKED_SHELL = 0x0000002;
        internal const uint DWM_CLOAKED_INHERITED = 0x0000004;

        internal enum GA : int
        {
            GetParent = 1,
            GetRoot = 2,
            GetRootOwner = 3
        }


        internal enum GW : int
        {
            HWndFirst = 0,
            HWndLast = 1,
            HWndNext = 2,
            HWndPrev = 3,
            Owner = 4,
            Child = 5,
            EnabledPopup = 6
        }

        public enum ExtendedWindowStyles : int
        {
            WS_EX_TOOLWINDOW = 0x00000080,
            WS_EX_APPWINDOW = 0x00040000,
            WS_EX_WINDOWEDGE = 0x100,
            WS_EX_TOPMOST = 0x8,
            WS_EX_ACCEPTFILES = 0x10,
            WS_EX_NOACTIVATE = 0x08000000,
            WS_EX_TRANSPARENT = 0x00000020,
            WS_EX_LAYERED = 0x00080000,
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct RECT
        {
            public int LEFT, TOP, RIGHT, BOTTOM;
        }


        [StructLayout(LayoutKind.Sequential)]
        internal struct WINDOWPLACEMENT
        {
            public int length;
            public int flags;
            public SW showCmd;
            public POINT ptMinPosition;
            public POINT ptMaxPosition;
            public RECT rcNormalPosition;
        }

        public enum SW : int
        {
            Hide = 0,
            ShowNormal = 1,
            Normal = 1,
            ShowMinimized = 2,
            ShowMaximized = 3,
            Maximize = 3,
            ShowNoActivate = 4,
            Show = 5,
            Minimize = 6,
            ShowMinNoActivate = 7,
            ShowNA = 8,
            Restore = 9,
            ShowDefault = 10,
            ForceMinimize = 11,
            Max = 11
        }


        [DllImport("USER32.DLL", SetLastError = true)]
        internal static extern IntPtr GetShellWindow();

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool SetForegroundWindow(IntPtr hWnd);


        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool ShowWindow(IntPtr hWnd, SW nCmdShow);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool CloseWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool DestroyWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetWindowRect(IntPtr hWnd, ref RECT lpRect);

        [DllImport("dwmapi.dll", SetLastError = true)]
        internal static extern int DwmGetWindowAttribute(IntPtr hwnd, DWMWINDOWATTRIBUTE dwAttribute, out uint pvAttribute, int cbAttribute);

        [DllImport("dwmapi.dll", SetLastError = true)]
        internal static extern int DwmGetWindowAttribute(IntPtr hwnd, DWMWINDOWATTRIBUTE dwAttribute, out RECT pvAttribute, int cbAttribute);

        public enum SPI : uint
        {
            /// <summary>
            /// Retrieves the size of the work area on the primary display monitor. The work area is the portion of the screen not obscured
            /// by the system taskbar or by application desktop toolbars. The pvParam parameter must point to a RECT structure that receives
            /// the coordinates of the work area, expressed in virtual screen coordinates.
            /// To get the work area of a monitor other than the primary display monitor, call the GetMonitorInfo function.
            /// </summary>
            SPI_GETWORKAREA = 0x0030
        }

        [Flags]
        internal enum SPIF
        {
            None = 0x00,
            /// <summary>Writes the new system-wide parameter setting to the user profile.</summary>
            SPIF_UPDATEINIFILE = 0x01,
            /// <summary>Broadcasts the WM_SETTINGCHANGE message after updating the user profile.</summary>
            SPIF_SENDCHANGE = 0x02,
            /// <summary>Same as SPIF_SENDCHANGE.</summary>
            SPIF_SENDWININICHANGE = 0x02
        }

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool SystemParametersInfo(SPI uiAction, uint uiParam, ref RECT pvParam, SPIF fWinIni);


        [Flags()]
        internal enum SetWindowPosFlags : uint
        {
            /// <summary>If the calling thread and the thread that owns the window are attached to different input queues,
            /// the system posts the request to the thread that owns the window. This prevents the calling thread from
            /// blocking its execution while other threads process the request.</summary>
            /// <remarks>SWP_ASYNCWINDOWPOS</remarks>
            AsynchronousWindowPosition = 0x4000,
            /// <summary>Prevents generation of the WM_SYNCPAINT message.</summary>
            /// <remarks>SWP_DEFERERASE</remarks>
            DeferErase = 0x2000,
            /// <summary>Draws a frame (defined in the window's class description) around the window.</summary>
            /// <remarks>SWP_DRAWFRAME</remarks>
            DrawFrame = 0x0020,
            /// <summary>Applies new frame styles set using the SetWindowLong function. Sends a WM_NCCALCSIZE message to
            /// the window, even if the window's size is not being changed. If this flag is not specified, WM_NCCALCSIZE
            /// is sent only when the window's size is being changed.</summary>
            /// <remarks>SWP_FRAMECHANGED</remarks>
            FrameChanged = 0x0020,
            /// <summary>Hides the window.</summary>
            /// <remarks>SWP_HIDEWINDOW</remarks>
            HideWindow = 0x0080,
            /// <summary>Does not activate the window. If this flag is not set, the window is activated and moved to the
            /// top of either the topmost or non-topmost group (depending on the setting of the hWndInsertAfter
            /// parameter).</summary>
            /// <remarks>SWP_NOACTIVATE</remarks>
            DoNotActivate = 0x0010,
            /// <summary>Discards the entire contents of the client area. If this flag is not specified, the valid
            /// contents of the client area are saved and copied back into the client area after the window is sized or
            /// repositioned.</summary>
            /// <remarks>SWP_NOCOPYBITS</remarks>
            DoNotCopyBits = 0x0100,
            /// <summary>Retains the current position (ignores X and Y parameters).</summary>
            /// <remarks>SWP_NOMOVE</remarks>
            IgnoreMove = 0x0002,
            /// <summary>Does not change the owner window's position in the Z order.</summary>
            /// <remarks>SWP_NOOWNERZORDER</remarks>
            DoNotChangeOwnerZOrder = 0x0200,
            /// <summary>Does not redraw changes. If this flag is set, no repainting of any kind occurs. This applies to
            /// the client area, the nonclient area (including the title bar and scroll bars), and any part of the parent
            /// window uncovered as a result of the window being moved. When this flag is set, the application must
            /// explicitly invalidate or redraw any parts of the window and parent window that need redrawing.</summary>
            /// <remarks>SWP_NOREDRAW</remarks>
            DoNotRedraw = 0x0008,
            /// <summary>Same as the SWP_NOOWNERZORDER flag.</summary>
            /// <remarks>SWP_NOREPOSITION</remarks>
            DoNotReposition = 0x0200,
            /// <summary>Prevents the window from receiving the WM_WINDOWPOSCHANGING message.</summary>
            /// <remarks>SWP_NOSENDCHANGING</remarks>
            DoNotSendChangingEvent = 0x0400,
            /// <summary>Retains the current size (ignores the cx and cy parameters).</summary>
            /// <remarks>SWP_NOSIZE</remarks>
            IgnoreResize = 0x0001,
            /// <summary>Retains the current Z order (ignores the hWndInsertAfter parameter).</summary>
            /// <remarks>SWP_NOZORDER</remarks>
            IgnoreZOrder = 0x0004,
            /// <summary>Displays the window.</summary>
            /// <remarks>SWP_SHOWWINDOW</remarks>
            ShowWindow = 0x0040,
        }

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, SetWindowPosFlags uFlags);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        internal static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr GetDesktopWindow();

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        internal delegate IntPtr WndProc(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr CreateWindowEx(
           ExtendedWindowStyles dwExStyle,
           string lpClassName,
           string lpWindowName,
           WindowStyles dwStyle,
           int x,
           int y,
           int nWidth,
           int nHeight,
           IntPtr hWndParent,
           IntPtr hMenu,
           IntPtr hInstance,
           IntPtr lpParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern IntPtr GetModuleHandle(string lpModuleName);

        [Flags]
        internal enum SendMessageTimeoutFlags : uint
        {
            SMTO_NORMAL = 0x0,
            SMTO_BLOCK = 0x1,
            SMTO_ABORTIFHUNG = 0x2,
            SMTO_NOTIMEOUTIFNOTHUNG = 0x8,
            SMTO_ERRORONEXIT = 0x20
        }

        internal const uint WM_GETMINMAXINFO = 0x0024;
        internal const uint WM_GETTEXT = 0x000D;
        internal const uint WM_GETTEXTLENGTH = 0x000E;
        internal const uint WM_DISPLAYCHANGE = 0x007E;
        internal const uint WM_TIMER = 0x0113;
        internal const uint WM_SETTINGCHANGE = 0x001A;

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        internal static extern IntPtr SendMessageTimeout(
            IntPtr hWnd,
            uint msg,
            IntPtr wParam,
            IntPtr lParam,
            SendMessageTimeoutFlags flags,
            uint timeout,
            out IntPtr result);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        internal static extern IntPtr SendMessageTimeout(
            IntPtr hWnd,
            uint msg,
            IntPtr wParam,
            ref MINMAXINFO lParam,
            SendMessageTimeoutFlags flags,
            uint timeout,
            out IntPtr result);

        [DllImport("user32.dll", EntryPoint = "SendMessageTimeout", SetLastError = true, CharSet = CharSet.Auto)]
        internal static extern IntPtr SendMessageTimeoutText(
            IntPtr hWnd,
            uint msg,              // Use WM_GETTEXT
            IntPtr nText,
            StringBuilder text,
            SendMessageTimeoutFlags flags,
            uint timeout,
            out IntPtr result);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        internal delegate void TimerProc(IntPtr hWnd, uint uMsg, IntPtr nIDEvent, uint dwTime);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr SetTimer(IntPtr hWnd, IntPtr nIDEvent, uint uElapse, TimerProc lpTimerFunc);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool KillTimer(IntPtr hWnd, IntPtr nIDEvent);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr GetAncestor(IntPtr hWnd, GA gaFlags);

        [StructLayout(LayoutKind.Sequential)]
        internal struct MINMAXINFO
        {
            public POINT ptReserved;
            public POINT ptMaxSize;
            public POINT ptMaxPosition;
            public POINT ptMinTrackSize;
            public POINT ptMaxTrackSize;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct DWM_THUMBNAIL_PROPERTIES
        {
            public int dwFlags;
            public RECT rcDestination;
            public RECT rcSource;
            public byte opacity;
            public int fVisible; //BOOL{TRUE,FALSE}
            public int fSourceClientAreaOnly; //BOOL{TRUE,FALSE}
        }

        [DllImport("dwmapi.dll", SetLastError = true)]
        internal static extern int DwmRegisterThumbnail(IntPtr dest, IntPtr src, out IntPtr thumb);

        [DllImport("dwmapi.dll", SetLastError = true)]
        internal static extern int DwmUnregisterThumbnail(IntPtr thumb);

        [DllImport("dwmapi.dll", PreserveSig = true)]
        internal static extern int DwmUpdateThumbnailProperties(IntPtr hThumbnail, ref DWM_THUMBNAIL_PROPERTIES props);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern int GetWindowThreadProcessId(IntPtr hWnd, out int processId);

        internal const uint PROCESS_QUERY_INFORMATION = 0x00000400;

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr OpenProcess(uint dwDesiredAccess, bool bInheritHandle, int dwProcessId);

        [DllImport("ntdll.dll", SetLastError = true)]
        internal static extern int NtQueryInformationProcess(IntPtr processHandle, int processInformationClass, ref PROCESS_EXTENDED_BASIC_INFORMATION processInformation, uint processInformationLength, IntPtr returnLength);

        internal const int ProcessBasicInformation = 0x00;

        internal const uint ProcessExtendedBasicInformationIsFrozen = 0x010;

        [StructLayout(LayoutKind.Sequential)]
        internal struct PROCESS_EXTENDED_BASIC_INFORMATION
        {
            public UIntPtr Size;
            public PROCESS_BASIC_INFORMATION BasicInfo;
            public uint Flags;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct PROCESS_BASIC_INFORMATION
        {
            public uint ExitStatus;
            public IntPtr PebBaseAddress;
            public UIntPtr AffinityMask;
            public int BasePriority;
            public UIntPtr UniqueProcessId;
            public UIntPtr InheritedFromUniqueProcessId;
        }


        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern bool CloseHandle(IntPtr hObject);

        internal enum SystemMetric : uint
        {
            SM_XVIRTUALSCREEN = 76,
            SM_YVIRTUALSCREEN = 77,
            SM_CXVIRTUALSCREEN = 78,
            SM_CYVIRTUALSCREEN = 79,
        }

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern int GetSystemMetrics(SystemMetric smIndex);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        internal delegate bool EnumMonitorsDelegate(IntPtr hMonitor, IntPtr hdcMonitor, ref RECT lprcMonitor, IntPtr dwData);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool EnumDisplayMonitors(IntPtr hdc, IntPtr lprcClip, EnumMonitorsDelegate lpfnEnum,  IntPtr dwData);

        [StructLayout(LayoutKind.Sequential)]
        internal struct MONITORINFO
        {
            public int cbSize;
            public RECT rcMonitor;
            public RECT rcWork;
            public uint dwFlags;
        }

        internal const uint DISPLAY_DEVICE_MIRRORING_DRIVER = 0x00000008;

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);

        internal static class NT_6_3
        {
            internal const int MDT_EFFECTIVE_DPI = 0;

            [DllImport("shcore.dll")]
            internal static extern uint GetDpiForMonitor(IntPtr hmonitor, int dpiType, out uint dpiX, out uint dpiY);
        }

        [DllImport("user32.dll", SetLastError=true)]
        internal static extern IntPtr CreateWindowEx(
           uint dwExStyle,
           string lpClassName,
           string lpWindowName,
           WindowStyles dwStyle,
           int x,
           int y,
           int nWidth,
           int nHeight,
           IntPtr hWndParent,
           IntPtr hMenu,
           IntPtr hInstance,
           IntPtr lpParam);
    }
}
