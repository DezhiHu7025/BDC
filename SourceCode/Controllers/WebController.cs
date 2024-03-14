using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data.SqlClient;
using System.Configuration;
using log4net;
using System.IO;
using System.Data;
using System.Text;
using System.Collections;
using System.Globalization;

using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel; //ref NPOI.OOXML + OpenXml4
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;

using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Microsoft.Office.Core;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;

namespace Kcis.Controllers
{
    public class WebController : Controller
    {
        private static ILog log = LogManager.GetLogger(typeof(WebController));

        //外部引用
        public ActionResult ManageJSPartial()
        {
            return PartialView();   
        } 

        //管理专区专用选单
        public ActionResult ManageMenuPartial()
        {
            return PartialView();
        }

       
        public ActionResult Fill_MainPage()
        {
            Kcis.Models.UserModel user = (Kcis.Models.UserModel)Session["UserProfile"];
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["CarDB"].ConnectionString);
            SqlTransaction sTrans = null;
 
            try
            {

                //取申请人帐号,姓名

                cn.Open();

                string strSQL = @"Select ParentName, email, today=CONVERT(varchar(100), GETDATE(), 111) from Shuttle.dbo.Car_ApplyForm where StudentNO='" + user.UserId + "'  ";
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_Apply = new DataTable();
                dt_Apply.Load(dr);
                dr.Dispose();
                cmd.Dispose();

                string strEmail = "";
                string strParentName = "";
                string strToDay = DateTime.Now.ToString("yyyy/MM/dd");
                if (dt_Apply.Rows.Count > 0)
                {
                    strParentName = dt_Apply.DefaultView[0]["ParentName"].ToString();
                    strEmail = dt_Apply.DefaultView[0]["Email"].ToString();
                    strToDay = dt_Apply.DefaultView[0]["today"].ToString();
                }



                ViewBag.ParentName = strParentName;
                ViewBag.Email = strEmail;
                ViewBag.Today = strToDay;

            }
            catch (Kcis.Models.KcisException e)
            {

                //sTrans.Rollback();

                log.Error(e.ToString());
               
            }
            catch (Exception e)
            {

               //sTrans.Rollback();

                log.Error(e.ToString());
            
            }

            finally
            {
                if (cn!=null  && cn.State != ConnectionState.Closed)
                    cn.Close();

            }


            return View();
        }

    
    }//end of class
}
