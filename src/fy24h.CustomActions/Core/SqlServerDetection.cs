using System;
using System.Collections.Generic;
using Microsoft.Win32;

namespace fy24h.CustomActions
{
    public class SqlServerDetection
    {
        private const string SQL2008REGKEY = @"SOFTWARE\Microsoft\Microsoft SQL Server\100\Tools\ClientSetup\CurrentVersion";
        private const string SQL2012REGKEY = @"SOFTWARE\Microsoft\Microsoft SQL Server\110\Tools\ClientSetup\CurrentVersion";
        private const string SqlRegKeyValue = "CurrentVersion";

        private static bool IsSql2008Installed()
        {
            bool foundFg = false;
            string regValue = "";
            if (Utility.GetRegistryValue(RegistryHive.LocalMachine, SQL2008REGKEY, SqlRegKeyValue, RegistryValueKind.String, out regValue))
            {
                if (regValue.StartsWith("10"))
                {
                    foundFg = true;
                }
            }

            return foundFg;
        }

        private static bool IsSql2012Installed()
        {
            bool foundFg = false;
            string regValue = "";
            if (Utility.GetRegistryValue(RegistryHive.LocalMachine, SQL2012REGKEY, SqlRegKeyValue, RegistryValueKind.String, out regValue))
            {
                if (regValue.StartsWith("11"))
                {
                    foundFg = true;
                }
            }

            return foundFg;
        }

        public static bool IsInstalled(ConstantEnum.SqlServerVersion sqlVersion)
        {
            bool installFg = false;

            switch (sqlVersion)
            {
                case ConstantEnum.SqlServerVersion.Sql2008:
                    {
                        installFg = IsSql2008Installed();
                        break;
                    }
                case ConstantEnum.SqlServerVersion.Sql2012:
                    {
                        installFg = IsSql2012Installed();
                        break;
                    }
            }

            return installFg;
        }
    }
}
