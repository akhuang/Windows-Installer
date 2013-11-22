using System;
using System.Collections.Generic;
using System.Text;
using System.Globalization;
using System.IO;
using Microsoft.Win32;

namespace fy24h.CustomActions
{
    public class InternetInformationServicesDetection
    {
        private const string IISRegKeyName = "Software\\Microsoft\\InetStp";
        private const string IISRegKeyValue = "MajorVersion";
        private const string IISRegKeyMinorVersionValue = "MinorVersion";

        #region GetRegistryValue

        #endregion

        private static bool IsIIS7Installed()
        {
            bool foundFg = false;
            int regValue = 0;
            if (Utility.GetRegistryValue(RegistryHive.LocalMachine, IISRegKeyName, IISRegKeyValue, RegistryValueKind.DWord, out regValue))
            {
                if (regValue >= 7)
                {
                    foundFg = true;
                }
            }

            return foundFg;
        }

        private static bool IsIIS6Installed()
        {
            bool foundFg = false;
            int regValue = 0;
            if (Utility.GetRegistryValue(RegistryHive.LocalMachine, IISRegKeyName, IISRegKeyValue, RegistryValueKind.DWord, out regValue))
            {
                if (regValue == 6)
                {
                    foundFg = true;
                }
            }

            return foundFg;
        }


        public static bool IsInstalled(ConstantEnum.InternetInformationServicesVersion iisVersion)
        {
            bool installFg = false;

            switch (iisVersion)
            {
                case ConstantEnum.InternetInformationServicesVersion.IIS6:
                    {
                        installFg = IsIIS6Installed();
                        break;
                    }
                case ConstantEnum.InternetInformationServicesVersion.IIS7:
                    {
                        installFg = IsIIS7Installed();
                        break;
                    }
            }

            return installFg;
        }
    }
}
