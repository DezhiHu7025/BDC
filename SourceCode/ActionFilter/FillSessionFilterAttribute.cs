using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using System.Web;
using log4net;
using System.Web.Routing;

namespace Common.ActionFilter
{
    public class FillSessionFilterAttribute : ActionFilterAttribute
    {
        private static ILog log = LogManager.GetLogger(typeof(FillSessionFilterAttribute));

        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {

            HttpContext httpcontext = HttpContext.Current;

            try
            {

                log.Debug("判斷目前要求的HttpSessionState存在");
                if (httpcontext.Session != null)
                {


                    log.Debug("判斷Session是否為新建立");

                    Kcis.Models.UserModel user = httpcontext.Session["UserProfile"] as Kcis.Models.UserModel;
                    //檢查Session
                    log.Debug("--檢查Session");
                    if (user == null || user.UserId == null || user.UserId.Equals(""))
                    {

                        //檢查Cookie
                        if (httpcontext.Request.Cookies[Kcis.Models.Config.WebURL] != null)
                        {
                            log.Debug("Session已經消失, 但Cookie仍存活著!");
                            //向客户端浏览器加入Cookie  
                            log.Debug("準備再生Session並向客戶端瀏覽器refresh Cookie!");
                            string cookieAccountID = Convert.ToString(httpcontext.Request.Cookies[Kcis.Models.Config.WebURL].Value);
                            log.Debug("cookieAccountID=" + cookieAccountID);
                            Kcis.Models.LogOnModel lm = new Kcis.Models.LogOnModel();
                            Kcis.Models.UserModel newUser = lm.BuildUserWithoutPassword(cookieAccountID);
                            log.Debug("---新Session取得成功");
                            if (!"Error".Equals(newUser.Status))
                            {
                                httpcontext.Session.Add("UserProfile", newUser);
                                log.Debug("---Session更新成功@過濾器");
                                //將帳號刷新到Cookie 
                                HttpCookie hcUserObject = new HttpCookie(Kcis.Models.Config.WebURL);
                                //httpcontext.Response.Cookies["username"].Expires = System.DateTime.Now.AddSeconds;
                                hcUserObject.Expires = System.DateTime.Now.AddDays(1);
                                hcUserObject.Value = newUser.UserId;
                                httpcontext.Request.Cookies.Add(hcUserObject);
                                log.Debug("---Cookie更新成功@過濾器");
                            }
                            else
                                throw new Exception("在取用Cookie accountid後,數據庫仍比對錯誤!");

                        }

                    }
                    else
                        log.Debug("目前仍在登入成功狀態！");



                }



                base.OnActionExecuting(filterContext);
            }
            catch (Exception e)
            {
                log.Error(e.ToString());

            }

        }


    }
}// end of namespace