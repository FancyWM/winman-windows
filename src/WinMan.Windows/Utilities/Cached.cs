using System;

namespace WinMan.Windows.Utilities
{
    internal struct Cached<T>
    {
        public T? Value = default;
        public long Revision = default;

        public Cached()
        {
        }
    }
}
