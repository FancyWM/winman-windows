using Microsoft.Win32;

using System;


namespace WinMan.Windows.Utilities
{
    internal static class VirtualDesktopRegistryHelper
    {
        private const string VIRTUALDESKTOP_REG_PATH = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops";
        private const string CURRENT_VIRTUALDESKTOP_VALUE = "CurrentVirtualDesktop";

        private static readonly RegistryHive _hive = RegistryHive.CurrentUser;
        private static readonly RegistryView _view = Environment.Is64BitOperatingSystem
            ? RegistryView.Registry64
            : RegistryView.Registry32;

        private static readonly RegistryKey? _registryKey = OpenRegistryKey();

        public static Guid GetCurrentDesktopGuid()
        {
            try
            {
                object? value = _registryKey?.GetValue(CURRENT_VIRTUALDESKTOP_VALUE);
                if (value is byte[] guidBytes && guidBytes.Length == 16)
                {
                    return new Guid(guidBytes);
                }

                return Guid.Empty;
            }
            catch
            {
                return Guid.Empty;
            }
        }

        private static RegistryKey? OpenRegistryKey()
        {
            try
            {
                return RegistryKey.OpenBaseKey(_hive, _view).OpenSubKey(VIRTUALDESKTOP_REG_PATH, false);
            }
            catch 
            {
                return null;
            }
        }
    }
}