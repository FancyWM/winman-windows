
using System;
using System.Runtime.InteropServices;

using WinMan.Windows.DllImports;

using static WinMan.Windows.DllImports.NativeMethods;

namespace WinMan.Windows
{
    internal static class MarshalExtensions
    {
        public static string MarshalIntoString(this HSTRING hStr)
        {
            unsafe
            {
                uint length = 0;
                PCWSTR pBuffer = WindowsGetStringRawBuffer(hStr, &length);
                string str = new((char*)pBuffer, 0, (int)length);
                WindowsDeleteString(hStr);
                return str;
            }
        }

        public static string MarshalIntoString(this Span<ushort> span)
        {
            int length = 0;
            while (length < span.Length && span[length] != 0)
            {
                length++;
            }
            return new(MemoryMarshal.Cast<ushort, char>(span.Slice(0, length)));
        }
    }
}
