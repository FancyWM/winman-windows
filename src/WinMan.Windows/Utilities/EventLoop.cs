using System;
using System.Threading.Channels;

namespace WinMan.Windows.Utilities
{
    internal class EventLoop
    {
        public event EventHandler<UnhandledExceptionEventArgs>? UnhandledException;

        private readonly Channel<(Action action, TimeSpan queued)> m_channel =
            Channel.CreateUnbounded<(Action action, TimeSpan queued)>(new UnboundedChannelOptions
            {
                SingleReader = true,
                AllowSynchronousContinuations = false
            });

        public void InvokeAsync(Action action)
        {
            m_channel.Writer.TryWrite((action, SteadyClock.Now));
        }

        public void Run()
        {
            var reader = m_channel.Reader;
            while (reader.WaitToReadAsync().AsTask().GetAwaiter().GetResult())
            {
                while (reader.TryRead(out var item))
                {
                    var (action, queued) = item;
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
        }

        public void Shutdown()
        {
            m_channel.Writer.Complete();
        }
    }
}
