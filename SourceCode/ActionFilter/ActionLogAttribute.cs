using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Security.Principal;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;
using System.Web.Security;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;
using System.IO;
using log4net;


namespace Common.ActionFilter
{
    public class ActionLogAttribute : ActionFilterAttribute
    {
        private static ILog log = LogManager.GetLogger(typeof(ActionLogAttribute));

        public string Description { get; set; }

        public override void OnActionExecuted(ActionExecutedContext filterContext)
        {
            //string strClientIP = filterContext.HttpContext.Request.UserHostAddress;

            Kcis.Models.UserModel user = filterContext.HttpContext.Session["UserProfile"] as Kcis.Models.UserModel;

            string userId = "";
            if (user != null && user.UserId != null && !user.UserId.Equals(""))
                userId = user.UserId;

            String message = String.Format(

                "Method=[{0}], Action=[{1}], Controller=[{2}], IPAddress=[{3}], UserId=[{4}], " +

                "TimeStamp=[{5}]",

                "OnActionExecuted",

                filterContext.RouteData.Values["action"] as String,

                filterContext.Controller.ToString(),

                filterContext.HttpContext.Request.UserHostAddress,
                
                userId,

                filterContext.HttpContext.Timestamp);

            log.Debug(message);
            //base.OnActionExecuted(filterContext);


           SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
           try
           {
               cn.Open();

               DataTable dt = new DataTable();

               string sql = @"Insert into SysLog(LogId, Method, Action, Controller, UserId, IP, ActionTime, SCode, SDescription) 
                                           values(@LogId, @Method, @Action, @Controller, @UserId, @IP, getdate(), @SCode, @SDescription)";

               SqlCommand cmd = new SqlCommand(sql, cn);
               cmd.Parameters.Add("@LogId", SqlDbType.NVarChar).Value = Convert.ToString(System.Guid.NewGuid());
               cmd.Parameters.Add("@Method", SqlDbType.NVarChar).Value = "OnActionExecuted";
               cmd.Parameters.Add("@Action", SqlDbType.NVarChar).Value = (string)filterContext.RouteData.Values["action"];
               cmd.Parameters.Add("@Controller", SqlDbType.NVarChar).Value = filterContext.Controller.ToString();
               cmd.Parameters.Add("@UserId", SqlDbType.NVarChar).Value = userId;
               cmd.Parameters.Add("@IP", SqlDbType.NVarChar).Value = filterContext.HttpContext.Request.UserHostAddress;

               string SCode = "";
               if (Description.IndexOf("-") >= 0)
                   SCode = Description.Substring(0, Description.IndexOf("-"));
 
               string SDescription = Description.Substring(Description.LastIndexOf('-') + 1, Description.Length - Description.LastIndexOf('-') - 1);

               cmd.Parameters.Add("@SCode", SqlDbType.NVarChar).Value = SCode;
               cmd.Parameters.Add("@SDescription", SqlDbType.NVarChar).Value = SDescription;


               cmd.ExecuteNonQuery();
               cmd.Dispose();
           }catch (Exception e){
                log.Error(e.Message);
           }finally{
                if (cn.State != ConnectionState.Closed)
                    cn.Close();
            }

        }// end of function

    }// end of class
}