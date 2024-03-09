using System;
using System.Runtime.InteropServices;

namespace WinMan.Windows.Com.Build22631R3085
{
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("3F07F4BE-B107-441A-AF0F-39D82529072C")]
    internal interface IComVirtualDesktop
    {
        bool IsViewVisible(IComApplicationView view);
        Guid GetId();
        IntPtr GetName();
        [return: MarshalAs(UnmanagedType.HString)]
        string GetWallpaperPath();
        bool IsRemote();
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	[Guid("53F5CA0B-158F-4124-900C-057158060B27")]
	internal interface IComVirtualDesktopManagerInternal
    {
		int GetCount();
		void MoveViewToDesktop(IComApplicationView view, IComVirtualDesktop desktop);
		bool CanViewMoveDesktops(IComApplicationView view);
        IComVirtualDesktop GetCurrentDesktop();
		void GetDesktops(out IComObjectArray desktops);
		[PreserveSig]
		int GetAdjacentDesktop(IComVirtualDesktop from, int direction, out IComVirtualDesktop desktop);
		void SwitchDesktop(IComVirtualDesktop desktop);
        IComVirtualDesktop CreateDesktop();
        void MoveDesktop(IComVirtualDesktop desktop, int nIndex);
        void RemoveDesktop(IComVirtualDesktop desktop, IComVirtualDesktop fallback);
        IComVirtualDesktop FindDesktop(ref Guid desktopid);
        void GetDesktopSwitchIncludeExcludeViews(IComVirtualDesktop desktop, out IComObjectArray unknown1, out IComObjectArray unknown2);
        void SetDesktopName(IComVirtualDesktop desktop, [MarshalAs(UnmanagedType.HString)] string name);
        void SetDesktopWallpaper(IComVirtualDesktop desktop, [MarshalAs(UnmanagedType.HString)] string path);
        void UpdateWallpaperPathForAllDesktops([MarshalAs(UnmanagedType.HString)] string path);
        void CopyDesktopState(IComApplicationView pView0, IComApplicationView pView1);
        void CreateRemoteDesktop([MarshalAs(UnmanagedType.HString)] string path, out IComVirtualDesktop desktop);
        void SwitchRemoteDesktop(IComVirtualDesktop desktop);
        void SwitchDesktopWithAnimation(IComVirtualDesktop desktop);
        void GetLastActiveDesktop(out IComVirtualDesktop desktop);
        void WaitForAnimationToComplete();
    }
}
