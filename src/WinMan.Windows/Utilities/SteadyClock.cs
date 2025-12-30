using System;
using System.Diagnostics;

namespace WinMan.Windows.Utilities
{
    internal static class SteadyClock
    {
        public static TimeSpan Now => s_stopatch.Elapsed;

        private static Stopwatch s_stopatch = CreateStopwatch();

        private static Stopwatch CreateStopwatch()
        {
            var s = new Stopwatch();
            s.Start();
            return s;
        }
    }
}
