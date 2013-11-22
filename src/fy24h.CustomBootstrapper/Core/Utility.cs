using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using System.IO;
using System.Globalization;

namespace CustomBootstrapper.Core
{
    public class Utility
    {
        public static bool GetRegistryValue<T>(RegistryHive hive, string key, string value, RegistryValueKind kind, out T data)
        {
            bool success = false;
            data = default(T);

            using (RegistryKey baseKey = RegistryKey.OpenRemoteBaseKey(hive, String.Empty))
            {
                if (baseKey != null)
                {
                    using (RegistryKey registryKey = baseKey.OpenSubKey(key, RegistryKeyPermissionCheck.ReadSubTree))
                    {
                        if (registryKey != null)
                        {
                            try
                            {
                                // If the key was opened, try to retrieve the value.
                                RegistryValueKind kindFound = registryKey.GetValueKind(value);
                                if (kindFound == kind)
                                {
                                    object regValue = registryKey.GetValue(value, null);
                                    if (regValue != null)
                                    {
                                        data = (T)Convert.ChangeType(regValue, typeof(T), CultureInfo.InvariantCulture);
                                        success = true;
                                    }
                                }
                            }
                            catch (IOException)
                            {
                                // The registry value doesn't exist. Since the
                                // value doesn't exist we have to assume that
                                // the component isn't installed and return
                                // false and leave the data param as the
                                // default value.
                            }
                        }
                    }
                }
            }
            return success;
        }
    }
}
