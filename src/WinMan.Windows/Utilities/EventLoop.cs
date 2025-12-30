using System;
using System.Collections.Concurrent;
using System.Diagnostics;

namespace WinMan.Windows.Utilities
{
    internal class EventLoop
    {
        public event EventHandler<UnhandledExceptionEventArgs>? UnhandledException;

        private BlockingCollection<(Action action, TimeSpan queued)> m_actions = new();

        public void InvokeAsync(Action action)
        {
            try
            {
                m_actions.Add((action, SteadyClock.Now));
            }
            catch (InvalidOperationException) when (m_actions.IsAddingCompleted)
            {
                return;
            }
        }

        public void Run()
        {
            foreach (var (action, queued) in m_actions.GetConsumingEnumerable())
            {
                try
                {
#if DEBUG
                    BlockTimer.Create(queued).LogIfExceeded(15, $"Schedule of {action.Method.Name}");
#endif
                    var t = BlockTimer.Create();
                    action.Invoke();
                    t.LogIfExceeded(15, action.Method.Name);
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
