using System;

namespace WinMan.Windows
{
    public class DummyVirtualDesktop : IVirtualDesktop
    {
        public bool IsCurrent => true;

        public int Index => 0;

        public string Name { get; private set; } = "Main Desktop";

        public IWorkspace Workspace { get; }

#pragma warning disable
        public event EventHandler<DesktopChangedEventArgs> Removed;
#pragma warning restore

        public DummyVirtualDesktop(IWorkspace workspace)
        {
            Workspace = workspace;
        }

        public void MoveWindow(IWindow window)
        {
        }

        public bool HasWindow(IWindow window)
        {
            return window.IsAlive;
        }

        public void SwitchTo()
        {
        }

        public void SetName(string newName)
        {
            Name = newName;
        }

        public void Remove()
        {
            throw new NotSupportedException();
        }

        public override int GetHashCode()
        {
            return 0;
        }

        public override string ToString()
        {
            return "DummyVirtualDesktop { IsCurrent = true }";
        }

        public override bool Equals(object obj)
        {
            return obj is DummyVirtualDesktop;
        }
    }
}
