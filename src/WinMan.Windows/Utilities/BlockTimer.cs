using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;

namespace WinMan.Windows.Utilities
{
    internal struct BlockTimer
    {
        private readonly TimeSpan m_startTime;

        private readonly string? m_callerName = null;
        private readonly string? m_callerFile = null;
        private readonly int m_callerLine = 0;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static BlockTimer Create([CallerMemberName] string callerName = "", [CallerFilePath] string callerFile = "", [CallerLineNumber] int callerLine = 0)
        {
#if DEBUG
            return new BlockTimer(SteadyClock.Now, callerName, callerFile, callerLine);
#else
            return new BlockTimer();
#endif
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static BlockTimer Create(TimeSpan startTime, [CallerMemberName] string callerName = "", [CallerFilePath] string callerFile = "", [CallerLineNumber] int callerLine = 0)
        {
#if DEBUG
            return new BlockTimer(startTime, callerName, callerFile, callerLine);
#else
            return new BlockTimer();
#endif
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private BlockTimer(TimeSpan startTime, string callerName, string callerFile, int callerLine)
        {
#if DEBUG
            m_startTime = startTime;
            m_callerName = callerName;
            m_callerFile = callerFile;
            m_callerLine = callerLine;
#endif
        }

        [Conditional("DEBUG")]
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public void LogIfExceeded(TimeSpan threshold, string? block = null, [CallerLineNumber] int endLine = 0)
        {
#if DEBUG
            var elapsed = SteadyClock.Now - m_startTime;
            if (elapsed > threshold)
            {
                Log(block, elapsed, threshold, endLine);
            }
#endif
        }

        [Conditional("DEBUG")]
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public void LogIfExceeded(double thresholdMs, string? block = null, [CallerLineNumber] int endLine = 0)
        {
#if DEBUG
            LogIfExceeded(TimeSpan.FromMilliseconds(thresholdMs), block, endLine);
#endif
        }

        private void Log(string? block, TimeSpan elapsed, TimeSpan threshold, int endLine)
        {
            StringBuilder sb = new();
            sb.Append("[WRN] Block ");
            if (!string.IsNullOrEmpty(block))
            {
                sb.Append('"');
                sb.Append(block);
                sb.Append('"');
            }

            if (!string.IsNullOrEmpty(m_callerFile) && !string.IsNullOrEmpty(m_callerName) && m_callerLine != 0 && endLine != 0)
            {
                sb.Append(' ');
                sb.Append('(');
                sb.Append(m_callerName);
                sb.Append('@');
                sb.Append(m_callerFile.AsSpan(m_callerFile.LastIndexOf(Path.DirectorySeparatorChar) + 1));
                sb.Append(':');
                sb.Append(m_callerLine);
                if (m_callerLine != endLine)
                {
                    sb.Append('-');
                    sb.Append(endLine);
                }
                sb.Append(')');
            }

            sb.Append(" took too long elapsed=");
            sb.Append(elapsed.TotalMilliseconds);
            sb.Append("ms threshold=");
            sb.Append(threshold.TotalMilliseconds);
            sb.Append("ms");

            Debug.WriteLine(sb.ToString());
        }
    }
}
