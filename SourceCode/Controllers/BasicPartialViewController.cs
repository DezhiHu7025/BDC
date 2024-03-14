using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data;
using System.IO;
using log4net;
using System.Data.SqlClient;
using System.Configuration;

namespace Kcis.Controllers
{
    public class BasicPartialViewController : Controller
    {
        private static ILog log = LogManager.GetLogger(typeof(BasicPartialViewController));


        public ActionResult LoadCommonInclude()
        {
            return PartialView();  //這樣才能調用特定controll內的partial view(ascx)
        }


        public ActionResult LoadCommonIncludeAuto()
        {
            return PartialView();  //這樣才能調用特定controll內的partial view(ascx)
        }

        public ActionResult PageFooter()
        {

            Kcis.Models.UserModel user = (Kcis.Models.UserModel)Session["UserProfile"];
 
            return PartialView();  //這樣才能調用特定controll內的partial view(ascx)
        }


        //管理专区专用选单
        public ActionResult FrontPageHead()
        {
            return PartialView();
        }

        [Common.ActionFilter.CheckSessionFilter]
        public ActionResult MobilePageHead()
        {

            Kcis.Models.UserModel user = (Kcis.Models.UserModel)Session["UserProfile"];

            if (!Kcis.Models.Utility.UtilitySystem.CheckUserExist(user))
                ViewBag.UserName = user.UserName;
            else
                ViewBag.UserName = "未登入";

            return PartialView();
        }

        public ActionResult PageHead()
        {

            string title = "";
            Kcis.Models.UserModel user = (Kcis.Models.UserModel)Session["UserProfile"];
            if (!Kcis.Models.Utility.UtilitySystem.CheckUserExist(user))
            {
                if (Request.Url.AbsoluteUri.IndexOf("UploadPackForm") > 0)
                    title = "";
                else
                    title = "<a href='" + Url.Action("LogIn", "Account") + "' style='color: white;'>[登入]</a>";
            }
            else
            {
                System.Globalization.DateTimeFormatInfo fmt = (new System.Globalization.CultureInfo("zh-TW")).DateTimeFormat;
                string strDate = string.Format(fmt, "{0:yyyy/MM/dd mm:ss}", DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss"));


                title = "<span id='HeadTitle' title='" + user.UserName + "' style='color: white;'>" + user.UserName + " 您好 </span>";
                title = title + "<a href='" + Url.Action("LogIn", "Account") + "' style='color: white;'>[登出]</a>";
                title = title + "<input type='hidden' id='UserId' value='" + user.UserId + "' /> ";
                title = title + "<input type='hidden' id='UserName' value='" + user.UserName + "' /> ";
                title = title + "<input type='hidden' id='CurrentDate' value='" + strDate + "' /> ";
                title = title + "<input type='hidden' id='Department' value='" + user.DepName + "' /> ";
            }
            ViewData["Title"] = title;


            if (Kcis.Models.Utility.UtilitySystem.CheckUserExist(user))
                ViewData["HomePage"] = user.HomePage;
            else
                ViewData["HomePage"] = "";


            return PartialView();
        }

       [Common.ActionFilter.CheckSessionFilter]
        public ActionResult LoadUserMenu()
        {


            Kcis.Models.UserModel user = (Kcis.Models.UserModel)Session["UserProfile"];
            log.Debug("S1---LoadOAMenu()");

            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            try
            {
                //cn.Open();

                string strMenu = "";
                ViewData["strMenu"] = strMenu;

                return PartialView();

            }
            catch (Exception e)
            {
                log.Debug(" Error--" + e.ToString());
                return PartialView();
            }
            finally
            {
                if (cn.State != ConnectionState.Closed)
                    cn.Close();
            }


        }// end of method

    }
}
