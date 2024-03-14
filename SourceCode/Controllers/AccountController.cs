using System;
using System.Collections.Generic;
using System.Linq;
using System.Transactions;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using WebMatrix.WebData;
using log4net;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;
using Common.Models;
using Newtonsoft.Json;

namespace Kcis.Controllers
{

    public class AccountController : Controller
    {
        private static ILog log = LogManager.GetLogger(typeof(AccountController));

        public ActionResult Index()
        {

            return RedirectToAction("Login", "Account");

            //return View();
        }

        public ActionResult Login(string returnUrl)
        {

            //Session.Remove("UserProfile");
            //Response.Cookies[Kcis.Models.Config.WebURL].Expires = System.DateTime.Now.AddDays(-1);

            //ViewBag.ReturnUrl = returnUrl;
            //return View();

            return RedirectToAction("XSCC", "Manager");
        }

 
 

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult LogOff()
        {
            try
            {
                Session.Remove("UserProfile");
                Response.Cookies[Kcis.Models.Config.WebURL].Expires = System.DateTime.Now.AddDays(-1);

            }
            catch (Exception e)
            {
                log.Error(e.ToString());

            }
            finally
            {
                //if (cn.State != ConnectionState.Closed)
                //    cn.Close();

            }

            //return RedirectToAction("LoginSyncPage", "Account");
            return View();
        }


        /*
  * Creater:sam cheng 2015/5/20
  * 登入頁面帳號密碼接收
  */
        [HttpGet]
        public ActionResult LogInCheck(string UserId, string Password, string Lang, string returnUrl)
        {
            string strMessage = "";
            try
            {
                returnUrl = returnUrl.Replace("amp;", "");
                HttpBrowserCapabilitiesBase bc = Request.Browser;
                log.Debug("Login IP = " + HttpContext.Request.UserHostAddress);
                log.Debug("login userid   =" + UserId);
                log.Debug("login Password =" + Password);
                log.Debug("---瀏覽器種類  :" + bc.Browser);
                log.Debug("---瀏覽器版本  :" + bc.Version);
                log.Debug("---作業系統種類:" + bc.Platform);
                log.Debug("Browser lang   :" + Request.UserLanguages[0]);
                Session.Remove("UserProfile");

                //產生UserProfile
                Kcis.Models.LogOnModel lm = new Kcis.Models.LogOnModel();
                Kcis.Models.UserModel user = lm.BuildUserWithPassword(UserId, Password);


                if (!"Error".Equals(user.Status))
                {
                    // Login successfully
                    user.Lang = Lang;

                    Session.Add("UserProfile", user);
                    HttpCookie hcUserObject = new HttpCookie(Kcis.Models.Config.WebURL);
                    hcUserObject.Expires = System.DateTime.Now.AddDays(1);
                    hcUserObject.Value = user.UserId;
                    try
                    {
                        Response.Cookies.Add(hcUserObject);
                        log.Debug("---Cookie更新成功");
                    }
                    catch (Exception exCookie)
                    {
                        log.Debug("此瀏覽器不支持Cookie, 當Session timeout就必須登出! message:" + exCookie.Message);
                    }

                    //strMessage = "login_ok";
                    //log.Debug("---登入成功");


                    log.Debug("------------returnUrl="+ returnUrl);
                    if (returnUrl!=null && returnUrl.Length>3 && !returnUrl.Equals("/BDC"))  //优先导回
                        strMessage = returnUrl;
                    else if (!user.HomePage.Equals(""))
                        strMessage = Url.Content("~/" + user.HomePage);      //若群组有市志home则进入
                    else
                        strMessage = Url.Content("~/Manager/Index");         //都没有则进入首页

                    strMessage = "[ok]" + strMessage;

                }
                else
                {
                    strMessage = user.Remark;
                    log.Debug("---登入失敗");
                }

                log.Debug("strMessage=" + strMessage);

                log.Debug("~~~~~~~~~~~~~~~~~~2 accountid=" + UserId);

                return Content(strMessage);

            }
            catch (Exception e)
            {
                log.Error(e.ToString());
                return Content(e.ToString());

            }
            finally
            {
                log.Debug("---login finally!");

            }

        }// end of method

        [HttpPost]
        public ActionResult LogInCheck(FormCollection collection)
        {
            string strMessage = "";
            string UserId = collection["UserId"].ToString();
            string Password = collection["Password"].ToString();
            //string Lang = collection["sel_Orignal"].ToString();
            string returnUrl = collection["returnUrl"].ToString();
            string strReturnUrl = "";  //回传的URL

            DataTable dt_PageData = new DataTable();
            dt_PageData.Columns.Add("strStatus", typeof(string));
            dt_PageData.Columns.Add("strMessage", typeof(string));
            dt_PageData.Columns.Add("ToURL", typeof(string));
            System.Data.DataRow dRow = dt_PageData.NewRow();
     

            try
            {
                returnUrl = returnUrl.Replace("amp;", "");
                HttpBrowserCapabilitiesBase bc = Request.Browser;
                log.Debug("Login IP = " + HttpContext.Request.UserHostAddress);
                log.Debug("login userid   =" + UserId);
                log.Debug("login Password =" + Password);
                log.Debug("---瀏覽器種類  :" + bc.Browser);
                log.Debug("---瀏覽器版本  :" + bc.Version);
                log.Debug("---作業系統種類:" + bc.Platform);
                log.Debug("Browser lang   :" + Request.UserLanguages[0]);
                Session.Remove("UserProfile");

                //產生UserProfile
                Kcis.Models.LogOnModel lm = new Kcis.Models.LogOnModel();
                Kcis.Models.UserModel user = lm.BuildUserWithPassword(UserId, Password);


                if (!"Error".Equals(user.Status))
                {
                    log.Debug("[1].登入成功-写入帐号Cookie");
                    Session.Add("UserProfile", user);
                    HttpCookie hcUserObject = new HttpCookie(Kcis.Models.Config.WebURL);
                    hcUserObject.Expires = System.DateTime.Now.AddDays(1);
                    hcUserObject.Value = user.UserId;
                    try
                    {
                        Response.Cookies.Add(hcUserObject);
                        log.Debug("---Cookie更新成功");
                    }
                    catch (Exception exCookie)
                    {
                        log.Debug("此瀏覽器不支持Cookie, 當Session timeout就必須登出! message:" + exCookie.Message);
                    }

                    log.Debug("[2].登入成功-设置跳转目的网址");
                    if (returnUrl != null && returnUrl.Length > 3 && !returnUrl.Equals("/BDC"))  //优先导回
                        strReturnUrl = returnUrl;
                    else if (!user.HomePage.Equals(""))
                        strReturnUrl = Url.Content("~/" + user.HomePage);      //若群组有市志home则进入
                    else
                        strReturnUrl = Url.Content("~/Manager/Index");         //都没有则进入首页

                    dRow["ToURL"] = strReturnUrl;
                    dRow["strStatus"] = "{ok}";
                    strMessage = "登入成功！";

                }
                else
                {
                    log.Debug("[1].登入失败-设置页面错误讯息");
                    dRow["strStatus"] = "{error}";
                    strMessage = user.Remark;
                }

                dRow["strMessage"] = strMessage;
                dt_PageData.Rows.Add(dRow);
                strMessage = JsonConvert.SerializeObject(dt_PageData);


            }
            catch (Exception e)
            {
                log.Error(e.ToString());
                dRow["strStatus"] = "{error}";
                dRow["strMessage"] = e.Message;
                dt_PageData.Rows.Add(dRow);
                strMessage = JsonConvert.SerializeObject(dt_PageData);

            }
            finally
            {
                log.Debug("---login finally!");

            }
            return Content(strMessage);
        }// end of method


        public ActionResult LoginSyncPage(string CuserID)
        {

            Session.Remove("UserProfile");
            Response.Cookies[Kcis.Models.Config.WebURL].Expires = System.DateTime.Now.AddDays(-1);

            ViewBag.CuserID = CuserID;
            return View();
 
        }// end of LoginSyncPage

        [HttpPost]
        public ActionResult LogInSyncFunc(FormCollection collection)
        {
            string strMessage = "";
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);

            DataTable dt_Dates = new DataTable();
            dt_Dates.Columns.Add("strStatus", typeof(string));
            dt_Dates.Columns.Add("strMessage", typeof(string));
            dt_Dates.Columns.Add("ToURL", typeof(string));
            System.Data.DataRow dRow = dt_Dates.NewRow();


            try
            {
                string Cuserid = collection["Cuserid"].ToString();
                string UserId = "";
                string returnUrl = "";
                string strReturnUrl = "";

                cn.Open();
                string strSQL = @"Select Accountid, ReturnURL=isnull(ReturnURL,'') from Common.[dbo].[kcis_ChinaLogin] Where Enabled='Y' and Cuserid=@Cuserid";
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                cmd.Parameters.Add("@Cuserid", SqlDbType.NVarChar).Value = Cuserid;
                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_ChinaLogin = new DataTable();
                dt_ChinaLogin.Load(dr);
                dr.Dispose();
                cmd.Dispose();

                if (dt_ChinaLogin.Rows.Count > 0)
                {
                    UserId = dt_ChinaLogin.DefaultView[0]["Accountid"].ToString();
                    strReturnUrl = dt_ChinaLogin.DefaultView[0]["ReturnURL"].ToString();
                }
                else
                {
                    throw new Exception("登入已逾时，请重新登入！");
                }




            

                //產生UserProfile
                Session.Remove("UserProfile");
                Kcis.Models.LogOnModel lm = new Kcis.Models.LogOnModel();
                Kcis.Models.UserModel user = lm.BuildUserWithPassword(UserId, "KcisP@ss9020");


                if (!"Error".Equals(user.Status))
                {
                    // Login successfully
                    Session.Add("UserProfile", user);
                    HttpCookie hcUserObject = new HttpCookie(Kcis.Models.Config.WebURL);
                    hcUserObject.Expires = System.DateTime.Now.AddDays(1);
                    hcUserObject.Value = user.UserId;
                    try
                    {
                        Response.Cookies.Add(hcUserObject);
                        log.Debug("---Cookie更新成功");
                    }
                    catch (Exception exCookie)
                    {
                        log.Debug("此瀏覽器不支持Cookie, 當Session timeout就必須登出! message:" + exCookie.Message);
                    }

           
                    log.Debug("[2].登入成功-设置跳转目的网址");
                
                    if (returnUrl != null && returnUrl.Length > 3 && !returnUrl.ToUpper().Equals("/BDC"))  //优先导回
                        strReturnUrl = returnUrl;
                    else if (!user.HomePage.Equals(""))
                        strReturnUrl = Url.Content("~/" + user.HomePage);      //若群组有市志home则进入
                    else
                        strReturnUrl = Url.Content("~/Home/Index");         //都没有则进入首页
                   

                    //若有strReturnUrl 学生可能会直接跳过去管理页面可能有安全隐患，后续可以在checksession加上身分控管


                    log.Debug("======================strReturnUrl222:" + strReturnUrl);
                    dRow["ToURL"] = strReturnUrl;
                    dRow["strStatus"] = "{ok}";
                    strMessage = "登入成功！";

                }
                else
                {
                    strMessage = user.Remark;
                    dRow["strStatus"] = "{error}";
                    log.Debug("---登入失敗");
                }

                dRow["strMessage"] = strMessage;
                dt_Dates.Rows.Add(dRow);
                strMessage = JsonConvert.SerializeObject(dt_Dates);

            }
            catch (Exception e)
            {
                log.Error(e.ToString());
                dRow["strStatus"] = "{error}";
                dRow["ToURL"] = "";
                dRow["strMessage"] = e.Message;
                dt_Dates.Rows.Add(dRow);
                strMessage = JsonConvert.SerializeObject(dt_Dates);

            }
            finally
            {
                if (cn.State != ConnectionState.Closed)
                    cn.Close();

            }
            return Content(strMessage);
        }// end of method



    }//end of class

}
