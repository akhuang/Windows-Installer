using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;
using System.IO;
using System.Globalization;
using Microsoft.Deployment.WindowsInstaller;
using System.Windows.Forms;

namespace fy24h.CustomActions
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

        public static void ShowMessage(string message, MessageBoxIcon icon)
        {
            MessageBox.Show(message, "", MessageBoxButtons.OK, icon);
        }

        public static void ShowMessage(string message)
        {
            ShowMessage(message, MessageBoxIcon.Error);
        }

        public static void ShowInfoMessage(string message)
        {
            ShowMessage(message, MessageBoxIcon.Information);
        }

        private static void WriteLog(InstallMessage messageType, Session session, Exception ex)
        {
            session.Message(messageType, new Record { FormatString = ex.ToString() });
        }

        public static void Error(Session session, Exception ex)
        {
            WriteLog(InstallMessage.Error, session, ex);
        }

        public static void Error(Session session, string message)
        {
            WriteLog(InstallMessage.Error, session, new Exception(message));
        }

        public static void Info(Session session, string message)
        {
            WriteLog(InstallMessage.Info, session, new Exception(message));
        }
    }
}
