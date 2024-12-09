using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dalton_Trapper.Utilities
{
    public static class PublishVersion
    {
        public static string GetVersion()
        {
            // If published, get apllication version number
            NameValueCollection nameValueTable = new NameValueCollection();
            if (Environment.GetEnvironmentVariable("ClickOnce_IsNetworkDeployed")?.ToLower() == "true")
            {
                string versionString = Environment.GetEnvironmentVariable("ClickOnce_CurrentVersion") ?? "0.0.0.0";
                return $"Version: {Version.Parse(versionString).ToString()}";
            }
            else
            {
                // Fallback in case the version attribute is not found
                return "Version: Not deployed";
            }
        }
    }
}
