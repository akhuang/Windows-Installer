using Microsoft.Deployment.WindowsInstaller;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace fy24h.CustomActions
{
    public partial class CustomActions
    {
        [CustomAction]
        public static ActionResult CheckWebDlgInput(Session session)
        {
            //System.Diagnostics.Debugger.Launch();
            ActionResult result = ActionResult.Failure;
            string crmAppVirtualDirectoryName = "";
            string crmAppValidFg = "0";
            string wcfAttachmentDirectoryName = "";

            session["CRMAPPVALIDFG"] = crmAppValidFg;
            session["SQLSERVERCONNECTIONCHECK"] = "0";

            try
            {
                crmAppVirtualDirectoryName = session["VIRTUALDIRECTORYNAME"];
                wcfAttachmentDirectoryName = session["DIRATTACHMENTS"];

                if (crmAppVirtualDirectoryName == null || crmAppVirtualDirectoryName.Trim() == "")
                {
                    Utility.ShowMessage("安装目录不能为空");
                }
                else if (wcfAttachmentDirectoryName == null || wcfAttachmentDirectoryName.Trim() == "")
                {
                    Utility.ShowMessage("文件存放目录不能为空");
                }
                else
                {
                    crmAppValidFg = "1";
                }
                result = ActionResult.Success;
            }
            catch (Exception ex)
            {
                Utility.Error(session, ex);
            }

            session["CRMAPPVALIDFG"] = crmAppValidFg;

            return result;
        }

        [CustomAction]
        public static ActionResult CheckMSSqlDlgInput(Session session)
        {
            //System.Diagnostics.Debugger.Launch();
            ActionResult result = ActionResult.Failure;
            string dbHostName = "";
            string dbName = "";
            string dbUserName = "";
            string dbPassword = "";
            string databaseValidFg = "0";

            session["DATABASEVALIDFG"] = databaseValidFg;

            try
            {
                dbHostName = session["DBHOST"];
                dbName = session["SQLDATABASE"];
                dbUserName = session["SQLADMINUSERNAME"];
                dbPassword = session["SQLADMINPASSWORD"];

                if (dbHostName == null || dbHostName.Trim() == "")
                {
                    Utility.ShowMessage("服务器名称不能为空");


                }
                else if (dbName == null || dbName.Trim() == "")
                {
                    Utility.ShowMessage("数据库名称不能为空");
                }
                else if (dbUserName == null || dbUserName.Trim() == "")
                {
                    Utility.ShowMessage("用户名不能为空");

                }
                else if (dbPassword == null || dbPassword.Trim() == "")
                {
                    Utility.ShowMessage("密码不能为空");
                }
                else
                {
                    databaseValidFg = "1";
                }
                result = ActionResult.Success;
            }
            catch (Exception ex)
            {
                Utility.Error(session, ex);
            }

            session["DATABASEVALIDFG"] = databaseValidFg;

            return result;
        }
    }
}
