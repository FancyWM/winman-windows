using System;
using System.Collections.Concurrent;
using System.Diagnostics;

namespace WinMan.Windows.Utilities
{
    internal class EventLoop
    {
        public event EventHandler<UnhandledExceptionEventArgs>? UnhandledException;

        private BlockingCollection<Action> m_actions = new BlockingCollection<Action>();

        public void InvokeAsync(Action action)
        {
            try
            {
                m_actions.Add(action);
            }
            catch (InvalidOperationException) when (m_actions.IsAddingCompleted)
            {
                return;
            }
        }

        public void Run()
        {
            foreach (var action in m_actions.GetConsumingEnumerable())
            {
                try
                {
                    action.Invoke();
                }
                catch (Exception e)
                {
                    UnhandledException?.Invoke(this, new UnhandledExceptionEventArgs(e, false));
                }
            }
        }

        public void Shutdown()
        {
            m_actions.CompleteAdding();
        }
    }
}
