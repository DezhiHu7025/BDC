using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;

namespace Kcis
{
    // 注意: 有关启用 IIS6 或 IIS7 经典模式的说明，
    // 请访问 http://go.microsoft.com/?LinkId=9394801
     
    public class MvcApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        { 
            AreaRegistration.RegisterAllAreas();

            WebApiConfig.Register(GlobalConfiguration.Configuration);
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);

            Kcis.Models.Config.HeadTitle = System.Web.Configuration.WebConfigurationManager.AppSettings["HeadTitle"];
            Kcis.Models.Config.GroupTitle = System.Web.Configuration.WebConfigurationManager.AppSettings["GroupTitle"];

            Kcis.Models.Config.FilePath = System.Web.Configuration.WebConfigurationManager.AppSettings["PhysicalStoragePath"];
            Kcis.Models.Config.WebURL = System.Web.Configuration.WebConfigurationManager.AppSettings["WebURL"];
            //Kcis.Models.Config.ExtWebURL = System.Web.Configuration.WebConfigurationManager.AppSettings["ExtWebURL"];
            Kcis.Models.Config.FMWebURL = System.Web.Configuration.WebConfigurationManager.AppSettings["FMWebURL"];
            Kcis.Models.Config.FMWeb_Port = System.Web.Configuration.WebConfigurationManager.AppSettings["FMWeb_Port"];


            Kcis.Models.Config.client_Host = System.Web.Configuration.WebConfigurationManager.AppSettings["client_Host"];
            Kcis.Models.Config.client_Port = Convert.ToInt16(System.Web.Configuration.WebConfigurationManager.AppSettings["client_Port"]);
            Kcis.Models.Config.client_Credentials_Account = System.Web.Configuration.WebConfigurationManager.AppSettings["client_Credentials_Account"];
            Kcis.Models.Config.client_Credentials_Password = System.Web.Configuration.WebConfigurationManager.AppSettings["client_Credentials_Password"];
            Kcis.Models.Config.client_FromAddr = System.Web.Configuration.WebConfigurationManager.AppSettings["client_FromAddr"];
            Kcis.Models.Config.client_FromName = System.Web.Configuration.WebConfigurationManager.AppSettings["client_FromName"];

            Kcis.Models.Config.IsEmail = Convert.ToBoolean(System.Web.Configuration.WebConfigurationManager.AppSettings["IsEmail"]);
            Kcis.Models.Config.IsSharing = Convert.ToBoolean(System.Web.Configuration.WebConfigurationManager.AppSettings["IsSharing"]);
            Kcis.Models.Config.iAutoCancelTime = Convert.ToInt32(System.Web.Configuration.WebConfigurationManager.AppSettings["AutoCancelTime"]);

       


            //AuthConfig.RegisterAuth();
        }
    }
}