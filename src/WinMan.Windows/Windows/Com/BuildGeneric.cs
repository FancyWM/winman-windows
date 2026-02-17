using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

using Microsoft.Win32;

namespace WinMan.Windows.Com.BuildGeneric
{
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("92CA9DCD-5622-4BBA-A805-5E9F541BD8C9")]
    internal interface IComObjectArrayPreserveSig
    {
        void GetCount(out int count);
        void GetAt(int index, ref Guid iid, out IntPtr obj);
    }

    internal class VirtualDesktopVTable
    {
        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int QueryInterfaceFn(IntPtr instance, ref Guid riid, out IntPtr ppvObject);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int GetIdFn(IntPtr instance, out Guid id);

        private readonly QueryInterfaceFn m_queryInterface;
        private readonly GetIdFn m_getId;

        public VirtualDesktopVTable(IntPtr pInterface)
        {
            if (pInterface == IntPtr.Zero)
                throw new ArgumentNullException(nameof(pInterface));

            IntPtr vTable = Marshal.ReadIntPtr(pInterface);
            T GetDelegate<T>(int index) where T : Delegate
            {
                IntPtr funcPtr = Marshal.ReadIntPtr(vTable, index * IntPtr.Size);
                return Marshal.GetDelegateForFunctionPointer<T>(funcPtr);
            }

            m_queryInterface = GetDelegate<QueryInterfaceFn>(0);
            m_getId = GetDelegate<GetIdFn>(4);
        }

        public int QueryInterface(IntPtr instance, ref Guid riid, out IntPtr ppvObject) => m_queryInterface(instance, ref  riid, out ppvObject);

        public int GetId(IntPtr instance, out Guid id) => m_getId(instance, out id);
    }

    internal class VirtualDesktopProxy
    {
        private readonly object m_managedHandle;
        private readonly VirtualDesktopVTable m_vtbl;
        private readonly IntPtr m_instance;

        public VirtualDesktopProxy(IntPtr pInterface)
        {
            if (pInterface == IntPtr.Zero)
                throw new ArgumentNullException(nameof(pInterface));

            m_managedHandle = Marshal.GetObjectForIUnknown(pInterface);
            m_vtbl = new(pInterface);
            m_instance = pInterface;
            Marshal.Release(pInterface);
        }

        internal IntPtr GetHandle() => m_instance;

        public bool IsInterfaceOf(Guid guid)
        {
            int hr = m_vtbl.QueryInterface(m_instance, ref guid, out IntPtr ppvObject);
            if (hr != 0 || ppvObject == IntPtr.Zero) return false;
            Marshal.Release(ppvObject);
            return m_instance == ppvObject;
        }

        public Guid GetId()
        {
            int hr = m_vtbl.GetId(m_instance, out Guid id);
            if (hr != 0) Marshal.ThrowExceptionForHR(hr);
            return id;
        }
    }

    /// <summary>
    /// Manually binds delegates to the vtable of <c>IVirtualDesktopManagerInternal</c>,
    /// exploiting two empirically observed invariants in Microsoft's implementation of this
    /// interface across all Windows 10 and Windows 11 builds to date.
    ///
    /// The first invariant is slot stability: although Microsoft changes the interface GUID
    /// with each major Windows release, the functions at vtable indices 0–9 have never been
    /// reordered or removed. New methods, when introduced, are always appended beyond index 9.
    /// This makes direct slot-index binding safe for the subset of methods this class wraps,
    /// regardless of which GUID variant is in use on the running system.
    ///
    /// The second invariant concerns the per-monitor API variant introduced temporarily in
    /// certain Windows 11 builds, which inserted an <c>hWndOrMon</c> parameter into several
    /// methods. Rather than detecting this variant by OS version, the constructor probes
    /// vtable slot 3 with a null monitor handle and inspects the returned HRESULT: a result
    /// of <c>RPC_S_NULL_REF_POINTER</c> (0x800706F4) indicates the classic single-monitor
    /// layout; any other result indicates the per-monitor layout. The correct overload for
    /// each affected method is selected once at construction time and stored, so callers are
    /// fully insulated from the distinction.
    ///
    /// The tradeoff is a reliance on undocumented and unguaranteed behaviour. Should Microsoft
    /// ever insert a new method before index 9, or change the error code returned by the
    /// per-monitor probe, this class will bind to the wrong functions and corrupt state or
    /// crash silently. These assumptions should be re-validated against the vtable layout
    /// whenever a new Windows feature update ships.
    ///
    /// This approach was conceived and originally implemented by
    /// <see href="https://github.com/veselink1">veselink1</see>;
    /// if you copy this code, a credit in your README or file header would be appreciated.
    /// </summary>
    internal class VirtualDesktopManagerInternalVTable
    {
        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int GetCountFn(IntPtr instance, out int count);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int GetCountPerMonFn(IntPtr instance, IntPtr hWndOrMon, out int count);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int MoveViewToDesktopFn(IntPtr instance, IntPtr view, IntPtr desktop);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int CanViewMoveDesktopsFn(IntPtr instance, IntPtr view, [MarshalAs(UnmanagedType.Bool)] out bool canMove);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int GetCurrentDesktopFn(IntPtr instance, out IntPtr desktop);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int GetCurrentDesktopPerMonFn(IntPtr instance, IntPtr hWndOrMon, out IntPtr desktop);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int GetDesktopsFn(IntPtr instance, out IComObjectArrayPreserveSig desktopsObjectArray);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int GetDesktopsPerMonFn(IntPtr instance, IntPtr hWndOrMon, out IComObjectArrayPreserveSig desktopsObjectArray);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int SwitchDesktopFn(IntPtr instance, IntPtr desktop);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int SwitchDesktopPerMonFn(IntPtr instance, IntPtr hWndOrMon, IntPtr desktop);

        private readonly GetCountFn? m_getCount;
        private readonly GetCountPerMonFn? m_getCountPerMon;
        private readonly MoveViewToDesktopFn m_moveViewToDesktop;
        private readonly CanViewMoveDesktopsFn m_canViewMoveDesktops;
        private readonly GetCurrentDesktopFn? m_getCurrentDesktop;
        private readonly GetCurrentDesktopPerMonFn? m_getCurrentDesktopPerMon;
        private readonly GetDesktopsFn? m_getDesktops;
        private readonly GetDesktopsPerMonFn? m_getDesktopsPerMon;
        private readonly SwitchDesktopFn? m_switchDesktop;
        private readonly SwitchDesktopPerMonFn? m_switchDesktopPerMon;

        public VirtualDesktopManagerInternalVTable(IntPtr pInterface)
        {
            if (pInterface == IntPtr.Zero)
            {
                throw new ArgumentNullException(nameof(pInterface));
            }

            IntPtr vTable = Marshal.ReadIntPtr(pInterface);
            T GetDelegate<T>(int index) where T : Delegate
            {
                IntPtr funcPtr = Marshal.ReadIntPtr(vTable, index * IntPtr.Size);
                return Marshal.GetDelegateForFunctionPointer<T>(funcPtr);
            }

            var getCountPerMon = GetDelegate<GetCountPerMonFn>(3);
            int hr = getCountPerMon(pInterface, IntPtr.Zero, out int count);
            bool isPerMonitor = unchecked((uint)hr) != unchecked(/* RPC_S_NULL_REF_POINTER */ 0x800706F4);

            if (isPerMonitor)
            {
                m_getCount = null;
                m_getCountPerMon = GetDelegate<GetCountPerMonFn>(3);
                m_moveViewToDesktop = GetDelegate<MoveViewToDesktopFn>(4);
                m_canViewMoveDesktops = GetDelegate<CanViewMoveDesktopsFn>(5);
                m_getCurrentDesktop = null;
                m_getCurrentDesktopPerMon = GetDelegate<GetCurrentDesktopPerMonFn>(6);
                m_getDesktops = null;
                m_getDesktopsPerMon = GetDelegate<GetDesktopsPerMonFn>(7);
                m_switchDesktop = GetDelegate<SwitchDesktopFn>(9);
            }
            else
            {
                m_getCount = GetDelegate<GetCountFn>(3);
                m_getCountPerMon = null;
                m_moveViewToDesktop = GetDelegate<MoveViewToDesktopFn>(4);
                m_canViewMoveDesktops = GetDelegate<CanViewMoveDesktopsFn>(5);
                m_getCurrentDesktop = GetDelegate<GetCurrentDesktopFn>(6);
                m_getCurrentDesktopPerMon = null;
                m_getDesktops = GetDelegate<GetDesktopsFn>(7);
                m_getDesktopsPerMon = null;
                m_switchDesktop = GetDelegate<SwitchDesktopFn>(9);
            }
        }

        public int GetCount(IntPtr instance, IntPtr hWndOrMon, out int count)
        {
            return m_getCount != default ? m_getCount(instance, out count) : m_getCountPerMon!(instance, hWndOrMon, out count);
        }

        public int MoveViewToDesktop(IntPtr instance, IntPtr view, IntPtr desktop) => m_moveViewToDesktop(instance, view, desktop);

        public int CanViewMoveDesktops(IntPtr instance, IntPtr view, out bool canMove) => m_canViewMoveDesktops(instance, view, out canMove);

        public int GetCurrentDesktop(IntPtr instance, IntPtr hWndOrMon, out IntPtr desktop)
        {
            return m_getCurrentDesktop != default ? m_getCurrentDesktop(instance, out desktop) : m_getCurrentDesktopPerMon!(instance, hWndOrMon, out desktop);
        }

        public int GetDesktops(IntPtr instance, IntPtr hWndOrMon, out IComObjectArrayPreserveSig desktopsObjectArray)
        {
            return m_getDesktops != default ? m_getDesktops(instance, out desktopsObjectArray) : m_getDesktopsPerMon!(instance, hWndOrMon, out desktopsObjectArray);
        }

        public int SwitchDesktop(IntPtr instance, IntPtr hWndOrMon, IntPtr desktop)
        {
            return m_switchDesktop != default ? m_switchDesktop(instance, desktop) : m_switchDesktopPerMon!(instance, hWndOrMon, desktop);
        }
    }

    /// <summary>
    /// A version-agnostic proxy for the undocumented <c>IVirtualDesktopManagerInternal</c>
    /// COM interface, designed to operate across Windows builds without hardcoded GUIDs or
    /// OS version checks.
    ///
    /// Rather than maintaining a lookup table of per-build GUIDs, this class uses two
    /// self-configuring strategies: <see cref="TryCreate"/> accepts caller-supplied GUIDs
    /// validated on a previous run, while <see cref="TryCreateSlow"/> discovers the correct
    /// GUIDs at runtime by scanning the registry for candidate interfaces (via
    /// <see cref="FindCandidateInterfaces"/>) and performing a round-trip validity check -
    /// creating a proxy for each candidate and confirming that the current desktop implements
    /// the corresponding virtual desktop interface. The per-monitor API variant introduced
    /// in Windows 11 is detected automatically by probing an expected HRESULT from vtable
    /// slot 3, avoiding any dependency on <see cref="System.Environment.OSVersion"/>.
    ///
    /// The tradeoff is that <see cref="TryCreateSlow"/> is substantially slower than
    /// <see cref="TryCreate"/> due to the parallel registry enumeration and COM round-trips.
    /// The intended pattern is to call <see cref="TryCreateSlow"/> once, persist the
    /// discovered <see cref="VirtualDesktopManagerInternalGuid"/> and
    /// <see cref="VirtualDesktopGuid"/> values, and use <see cref="TryCreate"/> on
    /// subsequent runs — falling back to <see cref="TryCreateSlow"/> only when the cached
    /// GUIDs no longer match (i.e. after a Windows feature update).
    ///
    /// This approach was conceived and originally implemented by
    /// <see href="https://github.com/veselink1">veselink1</see>;
    /// if you copy this code, a credit in your README or file header would be appreciated.
    /// </summary>
    internal class VirtualDesktopManagerInternalProxy
    {
        public Guid VirtualDesktopManagerInternalGuid { get; set; }
        public Guid VirtualDesktopGuid { get; set; }

        private readonly object m_managedHandle;
        private readonly VirtualDesktopManagerInternalVTable m_vtbl;
        private readonly IntPtr m_instance;

        public static VirtualDesktopManagerInternalProxy? TryCreateSlow()
        {
            var shell = (IComServiceProvider10PreserveSig?)Activator.CreateInstance(Type.GetTypeFromCLSID(ComGuids.CLSID_ImmersiveShell, true)!)
                ?? throw new COMException($"Failed to create instance of {ComGuids.CLSID_ImmersiveShell}");
            var comGuids = FindCandidateInterfaces();
            foreach (var virtualDesktopManagerGuid in comGuids)
            {
                int hr = shell.QueryService(ComGuids.CLSID_VirtualDesktopManagerInternal, virtualDesktopManagerGuid, out IntPtr pInstance);
                if (hr != 0 || pInstance == default)
                {
                    continue;
                }
                VirtualDesktopManagerInternalProxy manager = new(pInstance);
                VirtualDesktopProxy desktop = manager.GetCurrentDesktop(DllImports.NativeMethods.GetDesktopWindow());
                foreach (var virtualDesktopGuid in comGuids)
                {
                    if (desktop.IsInterfaceOf(virtualDesktopGuid))
                    {
                        manager.VirtualDesktopManagerInternalGuid = virtualDesktopManagerGuid;
                        manager.VirtualDesktopGuid = virtualDesktopGuid;
                        return manager;
                    }
                }
            }
            return null;
        }

        public static VirtualDesktopManagerInternalProxy? TryCreate(Guid virtualDesktopManagerInternalGuid, Guid virtualDesktopGuid)
        {
            var shell = (IComServiceProvider10PreserveSig?)Activator.CreateInstance(Type.GetTypeFromCLSID(ComGuids.CLSID_ImmersiveShell, true)!)
                ?? throw new COMException($"Failed to create instance of {ComGuids.CLSID_ImmersiveShell}");
            int hr = shell.QueryService(ComGuids.CLSID_VirtualDesktopManagerInternal, virtualDesktopManagerInternalGuid, out IntPtr pInstance);
            if (hr != 0 || pInstance == default)
            {
                return null;
            }
            VirtualDesktopManagerInternalProxy manager = new(pInstance);
            VirtualDesktopProxy desktop = manager.GetCurrentDesktop(DllImports.NativeMethods.GetDesktopWindow());
            if (desktop.IsInterfaceOf(virtualDesktopGuid))
            {
                manager.VirtualDesktopManagerInternalGuid = virtualDesktopManagerInternalGuid;
                manager.VirtualDesktopGuid = virtualDesktopGuid;
                return manager;
            }
            return null;
        }

        /// <summary>
        /// Scans the Windows registry to discover all COM interface GUIDs that share the Shell's
        /// internal free-threaded proxy stub CLSID (<c>{C90250F3-4D7D-4991-9B69-A5C5BC1C2AE6}</c>),
        /// producing the set of candidates from which the correct <c>IVirtualDesktopManagerInternal</c>
        /// and <c>IVirtualDesktop</c> GUIDs can be identified for the running Windows build.
        ///
        /// This approach replaces the conventional strategy of maintaining a hard-coded table of
        /// GUIDs indexed by Windows build number, which requires a manual update for every Windows
        /// feature release. Instead, the registry reflects whatever interfaces the currently
        /// installed Shell has registered, so the result is always up to date.
        ///
        /// The tradeoff is that the scan is imprecise: the proxy stub criterion is broad enough
        /// that multiple unrelated Shell interfaces will match, making a caller-side validity check
        /// (as performed by <see cref="TryCreateSlow"/>) necessary to identify the correct pair.
        /// The scan is parallelised to mitigate the cost, but it is still unsuitable for
        /// hot paths and should only be called once per session at startup.
        ///
        /// This approach was conceived and originally implemented by
        /// <see href="https://github.com/veselink1">veselink1</see>;
        /// if you copy this code, a credit in your README or file header would be appreciated.
        /// </summary>
        private static IEnumerable<Guid> FindCandidateInterfaces()
        {
            string basePath = @"SOFTWARE\Classes\Interface";
            string targetProxyStubClsid32 = "{C90250F3-4D7D-4991-9B69-A5C5BC1C2AE6}";

            var results = new ConcurrentBag<Guid>();
            using (var interfaceRoot = Registry.LocalMachine.OpenSubKey(basePath))
            {
                if (interfaceRoot == null)
                {
                    return results;
                }
                string[] subKeyNames = interfaceRoot.GetSubKeyNames();
                Parallel.ForEach(subKeyNames, iid =>
                {
                    try
                    {
                        using var iidKey = interfaceRoot.OpenSubKey(iid);
                        if (iidKey == null || iidKey.GetValue(null) != null)
                        {
                            return;
                        }
                        using var proxyKey = iidKey.OpenSubKey("ProxyStubClsid32");
                        if (proxyKey == null)
                        {
                            return;
                        }
                        var proxyValue = proxyKey.GetValue(null) as string;
                        if (string.Equals(proxyValue, targetProxyStubClsid32, StringComparison.OrdinalIgnoreCase))
                        {
                            results.Add(Guid.ParseExact(iid, "B"));
                        }
                    }
                    catch (System.Security.SecurityException)
                    {
                        // ignore
                    }
                });
            }
            return results;
        }

        public VirtualDesktopManagerInternalProxy(IntPtr pInterface)
        {
            if (pInterface == IntPtr.Zero)
            {
                throw new ArgumentNullException(nameof(pInterface));
            }

            m_managedHandle = Marshal.GetObjectForIUnknown(pInterface);
            m_instance = pInterface;
            m_vtbl = new(pInterface);
            Marshal.Release(pInterface);
        }

        public int GetCount(IntPtr hWndOrMon)
        {
            int hr = m_vtbl.GetCount(m_instance, hWndOrMon, out int count);
            if (hr != 0) Marshal.ThrowExceptionForHR(hr);
            return count;
        }

        public void MoveViewToDesktop(IComApplicationView view, VirtualDesktopProxy desktop)
        {
            IntPtr pView = Marshal.GetComInterfaceForObject(view, typeof(IComApplicationView));
            IntPtr pDesktop = desktop.GetHandle();
            try
            {
                int hr = m_vtbl.MoveViewToDesktop(m_instance, pView, pDesktop);
                GC.KeepAlive(desktop);
                if (hr != 0) Marshal.ThrowExceptionForHR(hr);
            }
            finally
            {
                Marshal.Release(pView);
            }
        }

        public bool CanViewMoveDesktops(IComApplicationView view)
        {
            IntPtr pView = Marshal.GetComInterfaceForObject(view, typeof(IComApplicationView));
            try
            {
                int hr = m_vtbl.CanViewMoveDesktops(m_instance, pView, out bool canMove);
                if (hr != 0) Marshal.ThrowExceptionForHR(hr);
                return canMove;
            }
            finally
            {
                Marshal.Release(pView);
            }
        }

        public VirtualDesktopProxy GetCurrentDesktop(IntPtr hWndOrMon)
        {
            int hr = m_vtbl.GetCurrentDesktop(m_instance, hWndOrMon, out IntPtr pDesktop);
            if (hr != 0) Marshal.ThrowExceptionForHR(hr);
            if (pDesktop == IntPtr.Zero) throw new COMException("Failed to get current desktop");
            return new VirtualDesktopProxy(pDesktop);
        }

        public List<VirtualDesktopProxy> GetDesktops(IntPtr hWndOrMon)
        {
            int hr = m_vtbl.GetDesktops(m_instance, hWndOrMon, out IComObjectArrayPreserveSig pArray);
            if (hr != 0) Marshal.ThrowExceptionForHR(hr);
            if (pArray == null) throw new COMException("Failed to get desktops array");

            pArray.GetCount(out int count);
            var list = new List<VirtualDesktopProxy>(count);
            for (int i = 0; i < count; i++)
            {
                pArray.GetAt(i, VirtualDesktopGuid, out IntPtr pDesktop);
                list.Add(new VirtualDesktopProxy(pDesktop));
            }
            return list;
        }

        public VirtualDesktopProxy GetDesktopAtIndex(IntPtr hWndOrMon, int index)
        {
            int hr = m_vtbl.GetDesktops(m_instance, hWndOrMon, out IComObjectArrayPreserveSig pArray);
            if (hr != 0) Marshal.ThrowExceptionForHR(hr);
            if (pArray == null) throw new COMException("Failed to get desktops array");

            pArray.GetAt(index, VirtualDesktopGuid, out IntPtr pDesktop);
            return new VirtualDesktopProxy(pDesktop);
        }

        public void SwitchDesktop(IntPtr hWndOrMon, VirtualDesktopProxy desktop)
        {
            IntPtr pDesktop = desktop.GetHandle();
            int hr = m_vtbl.SwitchDesktop(m_instance, hWndOrMon, pDesktop);
            GC.KeepAlive(desktop);
            if (hr != 0) Marshal.ThrowExceptionForHR(hr);
        }
    }
}
