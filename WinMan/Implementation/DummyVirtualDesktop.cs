namespace WinMan.Implementation
{
    internal class DummyVirtualDesktop : IVirtualDesktop
    {
        public bool IsCurrent => true;

        public void Add(IWindow window)
        {
        }

        public bool Contains(IWindow window)
        {
            return window.IsValid;
        }

        public override string ToString()
        {
            return "DummyVirtualDesktop { IsCurrent = true }";
        }

        public override bool Equals(object obj)
        {
            return obj is DummyVirtualDesktop;
        }

        public override int GetHashCode()
        {
            return 0;
        }
    }
}
