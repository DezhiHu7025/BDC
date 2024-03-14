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
    public class ActionWebPageController : Controller
    {
        private static ILog log = LogManager.GetLogger(typeof(ActionWebPageController));

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

        public ActionResult KcisPage0A(string strCID)
        {


            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            try
            {
                log.Debug("-------------------strCID=[" + strCID + "]");
                cn.Open();
                string strSQL = "Select *  from webapp.dbo.OA_SchoolActivity_ActionWeb Where Status='Y' and wid='" + strCID + "'";
                log.Debug("---sql=" + strSQL);
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_Product = new DataTable();
                dt_Product.Load(dr);
                dr.Dispose();
                cmd.Dispose();

                string LoginUserID = "Guest";


                string strLogID = "";
                string strApplyTime = "";
                new Kcis.Models.App.ManagerCommon().Fill_OA_GenerateKey5(cn, cmd, null, "Web", out strLogID, out strApplyTime);


                strSQL = @"INSERT INTO OA_SchoolActivity_Log
                                        ([LogID]
                                        ,[CID]
                                        ,[Stype]
                                        ,[Remark]
                                        ,[ModifierID]
                                        ,[ModifiDate]
                                        ,[ModifiDescr]
                                        ,[IP])
                                    VALUES
                                        (@LogID
                                        ,@CID
                                        ,'网页阅读'
                                        ,''
                                        ,@ModifierID
                                        ,getdate()
                                        ,''
                                        ,@IP) ";
                cmd = new SqlCommand(strSQL, cn);

                cmd.Parameters.Add("@LogID", SqlDbType.NVarChar).Value = strLogID;
                cmd.Parameters.Add("@CID", SqlDbType.NVarChar).Value = strCID;

                cmd.Parameters.Add("@ModifierID", SqlDbType.NVarChar).Value = LoginUserID;
                cmd.Parameters.Add("@IP", SqlDbType.NVarChar).Value = HttpContext.Request.UserHostAddress;

                cmd.ExecuteNonQuery();
                cmd.Dispose();

                strSQL = @"Select count(*) from OA_SchoolActivity_Log Where Stype=N'网页阅读' and CID='" + strCID + "'";
                cmd = new SqlCommand(strSQL, cn);
                int iViewCount = Convert.ToInt32(cmd.ExecuteScalar());



                ViewBag.ClickRate = iViewCount;
                ViewBag.dt_Product = dt_Product;  //数据回传
                ViewBag.PageID = Convert.ToString(System.Guid.NewGuid());
                ViewBag.PageMode = "View";
                ViewBag.RowsCount = dt_Product.Rows.Count;
            }
            catch (Exception e)
            {
                log.Error(e.ToString());
                return Content(e.Message);
            }

            finally
            {
                if (cn.State != ConnectionState.Closed)
                    cn.Close();
            }

            return PartialView();
        }

    }//end of class
}
