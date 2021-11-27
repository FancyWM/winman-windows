/// Windows 11 or Windows 10 21H2
using System;
using System.Runtime.InteropServices;

namespace WinMan.Windows.Com.Build21H2
{
	[ComImport]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	[Guid("536D3495-B208-4CC9-AE26-DE8111275BF8")]
	internal interface IComVirtualDesktop
	{
		bool IsViewVisible(IComApplicationView view);
		Guid GetId();
		IntPtr Unknown1();
		IntPtr GetName();
		[return: MarshalAs(UnmanagedType.HString)]
		string GetWallpaperPath();
	}

	[ComImport]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	[Guid("B2F925B9-5A0F-4D2E-9F4D-2B1507593C10")]
	internal interface IComVirtualDesktopManagerInternal
	{
		int GetCount(IntPtr hWndOrMon);
		void MoveViewToDesktop(IComApplicationView view, IComVirtualDesktop desktop);
		bool CanViewMoveDesktops(IComApplicationView view);
		IComVirtualDesktop GetCurrentDesktop(IntPtr hWndOrMon);
		void GetDesktops(IntPtr hWndOrMon, out IComObjectArray desktops);
		[PreserveSig]
		int GetAdjacentDesktop(IComVirtualDesktop from, int direction, out IComVirtualDesktop desktop);
		void SwitchDesktop(IntPtr hWndOrMon, IComVirtualDesktop desktop);
		IComVirtualDesktop CreateDesktop(IntPtr hWndOrMon);
		void MoveDesktop(IComVirtualDesktop desktop, IntPtr hWndOrMon, int nIndex);
		void RemoveDesktop(IComVirtualDesktop desktop, IComVirtualDesktop fallback);
		IComVirtualDesktop FindDesktop(ref Guid desktopid);
		void GetDesktopSwitchIncludeExcludeViews(IComVirtualDesktop desktop, out IComObjectArray unknown1, out IComObjectArray unknown2);
		void SetDesktopName(IComVirtualDesktop desktop, [MarshalAs(UnmanagedType.HString)] string name);
		void SetDesktopWallpaper(IComVirtualDesktop desktop, [MarshalAs(UnmanagedType.HString)] string path);
		void UpdateWallpaperPathForAllDesktops([MarshalAs(UnmanagedType.HString)] string path);
		void CopyDesktopState(IComApplicationView pView0, IComApplicationView pView1);
		int GetDesktopIsPerMonitor();
		void SetDesktopIsPerMonitor(bool state);
	}
}
