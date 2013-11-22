using Microsoft.Deployment.WindowsInstaller;
using Microsoft.Web.Administration;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Globalization;
using System.Linq;
using System.Text;
using fy24h.CustomActions.Core;

namespace fy24h.CustomActions
{
    public partial class CustomActions
    {
        private const string IISEntry = "IIS://localhost/W3SVC";
        //private const string SessionEntry = "ExistingWebsite";
        private const string SessionEntry = "SELECTEDWEBSITEID";
        private const string ServerComment = "ServerComment";
        private const string CustomActionException = "CustomActionException: ";
        private const string IISRegKey = @"Software\Microsoft\InetStp";
        private const string MajorVersion = "MajorVersion";
        private const string IISWebServer = "iiswebserver";
        private const string DeleteComoboxValue = " delete from ComboBox where ComboBox.Property='" + SessionEntry + "'";
        private const string GetComboContent = "select * from ComboBox where ComboBox.Property='" + SessionEntry + "'";
        private const string AvailableSites = "select * from AvailableWebSites";
        private const string DeleteDataFromAvailableSites = "delete from AvailableWebSites";
        private const string SpecificSite = "select * from AvailableWebSites where WebSiteId=";


        static Session ObjSession { get; set; }

        [CustomAction]
        public static ActionResult GetExistWebsites(Session session)
        {
            ObjSession = session;
            //System.Diagnostics.Debugger.Launch();
            session.Message(InstallMessage.Info, new Record { FormatString = "GetExistWebsites" });
            try
            {
                //System.Diagnostics.Debugger.Launch();
                //clear the exist value
                View needDeleteValueView = session.Database.OpenView(DeleteComoboxValue);
                needDeleteValueView.Execute();

                View needDeleteValueAvailableSitesView = session.Database.OpenView(DeleteDataFromAvailableSites);
                needDeleteValueAvailableSitesView.Execute();

                View comboBoxView = session.Database.OpenView(GetComboContent);
                View availableWSView = session.Database.OpenView(AvailableSites);

                DirectoryEntry iisRoot = new DirectoryEntry("IIS://localhost/W3SVC");

                if (IsIIS7)
                {
                    Utility.Info(session, "GetWebSitesViaWebAdministration method start..");
                    GetWebSitesViaWebAdministration(comboBoxView, availableWSView);
                }
                else
                {
                    Utility.Info(session, "GetWebSitesViaMetabase method start..");
                    GetWebSitesViaMetabase(comboBoxView, availableWSView);
                }
            }
            catch (Exception ex)
            {
                Utility.Info(session, "GetExistWebsites Exception..");
                Utility.Error(session, ex.ToString());
                //session.Message(InstallMessage.Error, new Record { FormatString = "错误:" + ex.ToString() });

                //session.Log("CustomActionException: " + ex.ToString());

                return ActionResult.Failure;

            }

            return ActionResult.Success;
        }

        private static void GetWebSitesViaWebAdministration(View comboView, View availableView)
        {
            using (ServerManager iisManager = new ServerManager())
            {
                int order = 1;

                foreach (Site webSite in iisManager.Sites)
                {
                    string id = webSite.Id.ToString(CultureInfo.InvariantCulture);
                    string name = webSite.Name;
                    //string path = webSite.PhysicalPath();

                    if (webSite.Bindings.Count <= 0)
                    {
                        ObjSession.Message(InstallMessage.Error, new Record { FormatString = "***webSite.Bindings.Count =0***" });
                    }

                    string ip = "";
                    string port = "";
                    string hostName = "";

                    string[] bindingInfo = webSite.Bindings[0].BindingInformation.Split(new string[] { ":" }, StringSplitOptions.None);
                    ip = bindingInfo[0];
                    port = bindingInfo[1];
                    hostName = bindingInfo[2];

                    WebsiteEntity websiteInfo = new WebsiteEntity()
                    {
                        Id = id,
                        Name = name,
                        IP = ip,
                        Port = port,
                        HostName = hostName
                    };
                    StoreSiteDataInComboBoxTable(websiteInfo, order++, comboView);
                    StoreSiteDataInAvailableSitesTable(websiteInfo, availableView);
                }
            }
        }

        private static void GetWebSitesViaMetabase(View comboView, View availableView)
        {
            using (DirectoryEntry iisRoot = new DirectoryEntry(IISEntry))
            {
                int order = 1;

                foreach (DirectoryEntry webSite in iisRoot.Children)
                {
                    if (webSite.SchemaClassName.ToLower(CultureInfo.InvariantCulture)
                        == IISWebServer)
                    {
                        string id = webSite.Name;
                        string name = webSite.Properties[ServerComment].Value.ToString();
                        string path = webSite.PhysicalPath();

                        string ip = "";
                        string port = "";
                        string hostName = "";
                        //(webSite.Properties["ServerBindings"].Value) 


                        PropertyValueCollection pvc = webSite.Properties["ServerBindings"];
                        Utility.Info(ObjSession, "pvc.Count::" + pvc.Count);
                        if (pvc.Count <= 1)
                        {
                            string[] bindingInfo = ((string)webSite.Properties["ServerBindings"].Value).Split(new string[] { ":" }, StringSplitOptions.None);

                            ip = bindingInfo[0];
                            port = bindingInfo[1];
                            hostName = bindingInfo[2];
                            Utility.Info(ObjSession, "ip-" + ip);
                            Utility.Info(ObjSession, "port-" + port);
                            Utility.Info(ObjSession, "hostName-" + hostName);
                        }
                        else
                        {
                            string[] bindingInfo = ((string)pvc[0]).Split(new string[] { ":" }, StringSplitOptions.None);
                            ip = bindingInfo[0];
                            port = bindingInfo[1];
                            hostName = bindingInfo[2];

                            Utility.Info(ObjSession, "pvc[0]-" + pvc[0].ToString());
                        }

                        WebsiteEntity websiteInfo = new WebsiteEntity()
                        {
                            Id = id,
                            Name = name,
                            IP = string.IsNullOrEmpty(ip) ? "*" : ip,
                            Port = port,
                            HostName = hostName
                        };

                        StoreSiteDataInComboBoxTable(websiteInfo, order++, comboView);
                        StoreSiteDataInAvailableSitesTable(websiteInfo, availableView);
                    }
                }
            }
        }

        private static void StoreSiteDataInComboBoxTable(WebsiteEntity websiteInfo, int order, View comboView)
        {
            Record newComboRecord = new Record(4);
            newComboRecord[1] = SessionEntry;
            newComboRecord[2] = order;
            newComboRecord[3] = websiteInfo.Id;
            newComboRecord[4] = websiteInfo.Name;
            //newComboRecord[5] = physicalPath;
            comboView.Modify(ViewModifyMode.InsertTemporary, newComboRecord);

            if (order == 1)
            {
                ObjSession[SessionEntry] = websiteInfo.Id;
            }

        }

        private static void StoreSiteDataInAvailableSitesTable(WebsiteEntity websiteInfo, View availableView)
        {
            Record newWebSiteRecord = new Record(5);
            newWebSiteRecord[1] = websiteInfo.Id;
            newWebSiteRecord[2] = websiteInfo.Name;
            newWebSiteRecord[3] = websiteInfo.IP;
            newWebSiteRecord[4] = websiteInfo.Port;
            newWebSiteRecord[5] = websiteInfo.HostName;

            availableView.Modify(ViewModifyMode.InsertTemporary, newWebSiteRecord);
        }

        [CustomAction]
        public static ActionResult UpdatePropertyWithSelectedWebSite(Session session)
        {
            ActionResult result = ActionResult.Failure;
            //System.Diagnostics.Debugger.Launch();
            try
            {
                if (session == null) { throw new ArgumentNullException("session"); }

                string selectedWebSiteId = session[SessionEntry];
                session.Log("CA:::Found web site id: " + selectedWebSiteId);

                using (View availableWebSitesView = session.Database.OpenView(SpecificSite + selectedWebSiteId))
                {
                    availableWebSitesView.Execute();

                    using (Record record = availableWebSitesView.Fetch())
                    {
                        if ((record[1].ToString()) == selectedWebSiteId)
                        {
                            session["WEBSITE_ID"] = selectedWebSiteId;
                            session["WEBSITEDESCRIPTION"] = (string)record[2];
                            session["WEBSITEIP"] = (string)record[3];
                            session["WEBSITEPORT"] = (string)record[4];
                            session["WEBSITEHOSTHEADER"] = (string)record[5];

                        }
                    }
                }

                result = ActionResult.Success;
            }
            catch (Exception ex)
            {
                session.Message(InstallMessage.Error, new Record { FormatString = "Exception::::" + ex.ToString() });
                if (session != null)
                {
                    session.Log(CustomActionException + ex);
                }
            }

            return result;
        }

        private static bool IsIIS7
        {
            get
            {
                return InternetInformationServicesDetection.IsInstalled(ConstantEnum.InternetInformationServicesVersion.IIS7);
            }
        }
    }
}
