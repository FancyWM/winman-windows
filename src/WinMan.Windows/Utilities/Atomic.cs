using System;
using System.Runtime.CompilerServices;

namespace WinMan.Windows.Utilities
{
    internal class Atomic<T> where T : struct, IEquatable<T>
    {
        private T m_value;

        public Atomic()
        {
            m_value = default;
        }

        public Atomic(T value)
        {
            m_value = value;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public T Read()
        {
            lock (this)
            {
                return m_value;
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public T Exchange(in T newValue)
        {
            T original;
            lock (this)
            {
                original = m_value;
                m_value = newValue;
            }
            return original;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public T CompareExchange(in T value, in T comparand)
        {
            T original;
            lock (this)
            {
                original = m_value;
                if (original.Equals(comparand))
                {
                    m_value = value;
                }
            }
            return original;
        }
    }
}
