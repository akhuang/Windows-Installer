using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Deployment.WindowsInstaller;
using System.Data.SqlClient;
using System.Data;
using System.DirectoryServices;
using System.Globalization;
using System.Diagnostics;
using Microsoft.Web.Administration;
using Microsoft.Win32;
using fy24h.CustomActions.Core;

namespace fy24h.CustomActions
{
    public partial class CustomActions
    {
        [CustomAction]
        public static ActionResult CustomAction1(Session session)
        {
            session.Log("Begin CustomAction1");

            return ActionResult.Success;
        }

        [CustomAction]
        public static ActionResult CheckSqlServer(Session session)
        {
            session.Log("Begin CheckSqlServer...");

            session["SQLSERVERCHECK"] = "RedX";

            try
            {
                bool sqlServerCheckResult = SqlServerDetection.IsInstalled(ConstantEnum.SqlServerVersion.Sql2008) || SqlServerDetection.IsInstalled(ConstantEnum.SqlServerVersion.Sql2012);

                if (sqlServerCheckResult)
                {
                    session["SQLSERVERCHECK"] = "GreenCheck";
                }
            }
            catch (Exception ex)
            {
                session.Log("Error in custom action CheckSqlServer: {0} ", ex.ToString());

                return ActionResult.Failure;
            }

            session.Log("End CheckSqlServer...");

            return ActionResult.Success;
        }

        [CustomAction]
        public static ActionResult GetDatabaseSchemas(Session session)
        {
            //session.Log("Begin TestSqlServerConnection");

            //session["SQLSERVERCONNECTIONCHECK"] = "0";

            //session.Log("Opening view");
            //View lView = session.Database.OpenView("Delete from Combobox where Combobox.Property='' ");

            //try
            //{
            //    string connectionString = session["DATABASECONNECTIONSTRING"];
            //}
            //catch (Exception ex)
            //{
            //    session.Log("Error in custom action CheckSqlServer: {0} ", ex.ToString());

            //    return ActionResult.Failure;
            //}

            //session.Log("End CheckSqlServer...");

            return ActionResult.Success;
        }

        [CustomAction]
        public static ActionResult TestSqlServerConnection(Session session)
        {
            session.Log("Begin TestSqlServerConnection");

            session["SQLSERVERCONNECTIONCHECK"] = "0";

            ActionResult result = ActionResult.Failure;
            try
            {
                string connectionStringWithMaster = session["DATABASECONNECTIONSTRINGUSEMASTER"];
                string connectionString = session["DATABASECONNECTIONSTRING"];

                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    conn.Close();
                }

                session["SQLSERVERCONNECTIONCHECK"] = "1";
                result = ActionResult.Success;
            }
            catch (Exception ex)
            {
                session.Log("Error in custom action TestSqlServerConnection: {0} ", ex.ToString());
            }

            session.Log("End TestSqlServerConnection");


            if (result == ActionResult.Failure)
            {
                string errorMsg = "无法连接服务器";
                session.Message(InstallMessage.Error, new Record { FormatString = errorMsg });
                Utility.ShowMessage(errorMsg);
            }
            else
            {
                session.Message(InstallMessage.Info, new Record { FormatString = "连接成功" });
                Utility.ShowInfoMessage("连接成功");
            }

            return ActionResult.Success;
        }
    }
}