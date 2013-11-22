using Microsoft.Web.Administration;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Text;

namespace fy24h.CustomActions.Core
{
    public static class ExtensionMethods
    {
        private const string IISEntry = "IIS://localhost/W3SVC/";
        private const string Root = "/root";
        private const string Path = "Path";

        public static string PhysicalPath(this Site site)
        {
            if (site == null) { throw new ArgumentNullException("site"); }

            var root = site.Applications.Where(a => a.Path == "/").Single();
            var vRoot = root.VirtualDirectories.Where(v => v.Path == "/")
                .Single();

            // Can get environment variables, so need to expand them
            return Environment.ExpandEnvironmentVariables(vRoot.PhysicalPath);
        }

        public static string PhysicalPath(this DirectoryEntry site)
        {
            if (site == null) { throw new ArgumentNullException("site"); }

            string path;

            using (DirectoryEntry de = new DirectoryEntry(IISEntry
                + site.Name + Root))
            {
                path = de.Properties[Path].Value.ToString();
            }

            return path;
        }
    }

}
