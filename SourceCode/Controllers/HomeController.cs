using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Kcis.Controllers
{
    public class HomeController : Controller
    {
        [Common.ActionFilter.CheckSessionFilter]
        public ActionResult Index()
        {
            Kcis.Models.UserModel user = (Kcis.Models.UserModel)Session["UserProfile"];

            //if (user.GroupIds == null || (user.GroupIds.IndexOf("manager") < 0 && user.GroupIds.IndexOf("supplier") < 0))
                return RedirectToAction("Index","Manager"); 

            ViewBag.UserGroup = user.GroupIds;
            ViewBag.HomePage = user.HomePage;
            ViewBag.ActionID = user.ActionID;
            ViewBag.GroupIds = user.GroupIds; 
  
            return View();
        }


        public ActionResult BackendError()
        {
            return View();
        }
 
    }//end of class
}
