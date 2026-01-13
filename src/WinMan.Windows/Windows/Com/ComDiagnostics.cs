using System.Collections.Generic;
using System.Linq;
using Microsoft.Win32;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text.Json;

namespace WinMan.Windows.Com
{
    public class ComDiagnostics
    {
        public static string GetReportJson(string clsid)
        {
            var report = new ComReport
            {
                Clsid = clsid,
                Process = GetProcessContext(),
                Registration = GetClsidRegistration(clsid),
                OleSettings = GetOleSettings()
            };

            JsonSerializerOptions options = new() { WriteIndented = true };
            return JsonSerializer.Serialize(report, options);
        }

        private static ProcessContext GetProcessContext()
        {
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return new ProcessContext
            {
                UserSid = identity.User?.ToString(),
                IsElevated = principal.IsInRole(WindowsBuiltInRole.Administrator),
                ImpersonationLevel = identity.ImpersonationLevel.ToString(),
                Groups = identity.Groups?.Select(g => g.ToString()).ToList()
            };
        }

        private static ClsidRegistration GetClsidRegistration(string clsid)
        {
            var reg = new ClsidRegistration();
            string clsidPath = $@"CLSID\{clsid}";

            using (var key = Registry.ClassesRoot.OpenSubKey(clsidPath))
            {
                if (key == null) return reg;
                reg.Exists = true;
                reg.AppId = key.GetValue("AppID")?.ToString();
            }

            if (!string.IsNullOrEmpty(reg.AppId))
            {
                using var appIdKey = Registry.LocalMachine.OpenSubKey($@"SOFTWARE\Classes\AppID\{reg.AppId}");
                reg.AppIdExists = appIdKey != null;
            }

            return reg;
        }

        private static OleSettings GetOleSettings()
        {
            var settings = new OleSettings();
            using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Ole"))
            {
                if (key == null) return settings;

                settings.EnableDCOM = key.GetValue("EnableDCOM")?.ToString();
                settings.LegacyAuthenticationLevel = (int?)key.GetValue("LegacyAuthenticationLevel");
                settings.LegacyImpersonationLevel = (int?)key.GetValue("LegacyImpersonationLevel");

                settings.MachineAccessRestriction = GetSddl(key, "MachineAccessRestriction");
                settings.MachineLaunchRestriction = GetSddl(key, "MachineLaunchRestriction");
                settings.DefaultAccessPermission = GetSddl(key, "DefaultAccessPermission");
                settings.DefaultLaunchPermission = GetSddl(key, "DefaultLaunchPermission");
            }
            return settings;
        }

        private static string GetSddl(RegistryKey key, string valueName)
        {
            if (key.GetValue(valueName) is byte[] bytes)
            {
                try { return new RawSecurityDescriptor(bytes, 0).GetSddlForm(AccessControlSections.All); }
                catch { return "Error parsing SDDL"; }
            }
            return "Not Set";
        }

        public class ComReport
        {
            public string? Clsid { get; set; }
            public ProcessContext? Process { get; set; }
            public ClsidRegistration? Registration { get; set; }
            public OleSettings? OleSettings { get; set; }
        }

        public class ProcessContext
        {
            public string? UserSid { get; set; }
            public bool IsElevated { get; set; }
            public string? ImpersonationLevel { get; set; }
            public List<string>? Groups { get; set; }
        }

        public class ClsidRegistration
        {
            public bool Exists { get; set; }
            public string? AppId { get; set; }
            public bool AppIdExists { get; set; }
        }

        public class OleSettings
        {
            public string? EnableDCOM { get; set; }
            public int? LegacyAuthenticationLevel { get; set; }
            public int? LegacyImpersonationLevel { get; set; }
            public string? MachineAccessRestriction { get; set; }
            public string? MachineLaunchRestriction { get; set; }
            public string? DefaultAccessPermission { get; set; }
            public string? DefaultLaunchPermission { get; set; }
        }
    }
}