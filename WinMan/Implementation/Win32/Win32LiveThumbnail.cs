using System;
using System.ComponentModel;

using static WinMan.Implementation.Win32.NativeMethods;

namespace WinMan.Implementation.Win32
{
    internal class Win32LiveThumbnail : ILiveThumbnail
    {
        public IWindow SourceWindow => m_srcWindow;

        public IWindow DestinationWindow => m_destWindow;

        public Rectangle Location 
        {
            get => m_location;
            set => SetLocation(value);
        }

        private readonly IWindow m_srcWindow;
        private readonly IWindow m_destWindow;
        private readonly IntPtr m_thumb;
        private Rectangle m_location;
        private bool m_disposed;

        public Win32LiveThumbnail(IWindow destWindow, IWindow srcWindow)
        {
            m_srcWindow = srcWindow ?? throw new ArgumentNullException(nameof(srcWindow));
            m_destWindow = destWindow ?? throw new ArgumentNullException(nameof(srcWindow));

            if (srcWindow.Handle == destWindow.Handle)
            {
                throw new ArgumentException("Source and parent windows must differ");
            }

            if (DwmRegisterThumbnail(destWindow.Handle, srcWindow.Handle, out m_thumb) != 0)
            {
                throw new Win32Exception();
            }
        }

        private void SetLocation(Rectangle value)
        {
            var props = new DWM_THUMBNAIL_PROPERTIES
            {
                dwFlags = 0x00000008 | 0x00000004 | 0x00000001 | 0x00000010,
                fVisible = 1,
                fSourceClientAreaOnly = 0,
                opacity = 255,
                rcDestination = new RECT
                {
                    LEFT = value.Left,
                    TOP = value.Top,
                    RIGHT = value.Right,
                    BOTTOM = value.Bottom,
                },
            };

            if (DwmUpdateThumbnailProperties(m_thumb, ref props) != 0)
            {
                throw new Win32Exception();
            }

            m_location = value;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!m_disposed)
            {
                DwmUnregisterThumbnail(m_thumb);
                m_disposed = true;
            }
        }

        ~Win32LiveThumbnail()
        {
            Dispose(disposing: false);
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}