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
    public class CommonServiceController : Controller
    {
        private static ILog log = LogManager.GetLogger(typeof(CommonServiceController));

        public ActionResult Index()
        {
            return View();
        }



        //File upload package1_1
        //档案储存于暂存区
        //需具备form tag, 与商业单元无相依性
        [Common.ActionFilter.CheckSessionFilter]
        public ActionResult File_SaveTemp(IEnumerable<HttpPostedFileBase> files, FormCollection collection)
        {
            Kcis.Models.UserModel user = (Kcis.Models.UserModel)Session["UserProfile"];
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            SqlTransaction sTrans = null;
            try
            {
                //处理档案
                HttpFileCollectionBase myfiles = Request.Files;

                cn.Open();
                sTrans = cn.BeginTransaction();

                for (int i = 0; i < myfiles.Count; i++)
                {
                    HttpPostedFileBase upfile = myfiles[i];
                    if (upfile != null && upfile.ContentLength > 0)
                    {
                        ////给新档名
                        string strOrignName = upfile.FileName;
                        string strFullName = Convert.ToString(System.Guid.NewGuid()) + Path.GetExtension(upfile.FileName);
                        string filePath = Path.Combine(HttpContext.Server.MapPath("~/uploadfile"), strFullName);
                        upfile.SaveAs(filePath);

                        //save to db
                        //处理数据库


                        string strSQL = @"INSERT INTO [webapp].[dbo].[OA_FilesTemp_List]
                                           ([FKey]
                                           ,[UUID]
                                           ,[FileName]
                                           ,[CreateTime]
                                           ,[CreateUser], PageID)
                                     VALUES
                                           (@FKey 
                                           ,@UUID 
                                           ,@FileName 
                                           ,getdate()
                                           ,@CreateUser, @PageID ) ";
                        SqlCommand cmd = new SqlCommand(strSQL, cn);
                        cmd.Transaction = sTrans;
                        cmd.Parameters.Add("@FKey", SqlDbType.NVarChar).Value = collection["inp_ArtID"];
                        cmd.Parameters.Add("@UUID", SqlDbType.NVarChar).Value = strFullName;
                        cmd.Parameters.Add("@FileName", SqlDbType.NVarChar).Value = strOrignName;
                        cmd.Parameters.Add("@CreateUser", SqlDbType.NVarChar).Value = user.UserId;
                        cmd.Parameters.Add("@PageID", SqlDbType.NVarChar).Value = collection["inp_PageID"];

                        cmd.ExecuteNonQuery();
                        cmd.Dispose();

                    }

                }


                sTrans.Commit();
                sTrans.Dispose();


                return Content("OK");


            }
            catch (Kcis.Models.KcisException e)
            {
                sTrans.Rollback();
                log.Error(e.ToString());
                return Content("{Error}" + e.Message);
            }
            catch (Exception e)
            {
                sTrans.Rollback();
                log.Error(e.ToString());
                return Content("[Error]");
            }

            finally
            {
                if (cn.State != ConnectionState.Closed)
                    cn.Close();

            }
        }//end of func


        //File upload package1_2
        //回传档案清单
        //需具备div tag, 与商业单元无相依性, 支持pageid
        //Mode example = view-D-A
        public ActionResult File_List(string strArtID, string strPageID, string strPageMode)
        {

            log.Debug("-----strPageMode=" + strPageMode);
            log.Debug("-----strPageID=" + strPageID);
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            try
            {

                cn.Open();
                //PageID不相等时 就无条件显示
                string strSQL = @"Select * from ( 
                                    Select uuid, filename, CreateTime from  webapp.dbo.OA_Files_List  where fkey=@fkey and (isnull(status,'')<>'暂时删除' or isnull(PageID,'')<>@pageid )
                                    union
                                    Select uuid, filename=filename +'[未储存]', CreateTime from webapp.dbo.OA_FilesTemp_List where fkey=@fkey and pageid=@pageid) aa
                                    order by aa.CreateTime desc";


                SqlCommand cmd = new SqlCommand(strSQL, cn);
                cmd.Parameters.Add("@fkey", SqlDbType.NVarChar).Value = strArtID;
                cmd.Parameters.Add("@pageid", SqlDbType.NVarChar).Value = strPageID;
                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_File = new DataTable();
                dt_File.Load(dr);
                dr.Dispose();
                cmd.Dispose();
                string strFileLinkList = "";

                for (int i = 0; i < dt_File.Rows.Count; i++)
                {
                    string strURL = Url.Content("~/uploadfile/");
                    strFileLinkList += "<a  target='kcis' style='text-decoration:none;' href='" + strURL + dt_File.DefaultView[i]["uuid"].ToString() + "'><span  style='font-size:18px;'>" + "文档" + (i + 1) + "：" + dt_File.DefaultView[i]["filename"].ToString() + "</span></a>&nbsp;&nbsp;&nbsp;&nbsp;";
                    if (strPageMode.ToLower().IndexOf("view")<0)
                        strFileLinkList += "<a target='kcis' href='#' name='FileDelete' id='" + dt_File.DefaultView[i]["uuid"].ToString() + "' ><img class='EditMode' alt='删除' src='" + Url.Content("~/Images/delete.gif") + "' style='border:0' /></a>";
                    strFileLinkList += "<p>";
                }

                log.Debug("strShortNameList=" + strFileLinkList);
                return Content(strFileLinkList);
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


        }

        //File upload package1_3
        //删除文档
        //需具备div tag, 与商业单元无相依性, 支持pageid
        public ActionResult File_DelTemp(string strArtID, string strUUID, string strPageID)
        {
            //Form修改储存时需加入以下D1, D2代码
            //检查是否有档案要删除D1-4
            //            strSQL = @"Select Count(*) from webapp.dbo.OA_Files_List where Status='暂时删除' and pageid='" + collection["inp_PageID"] + "'";
            //            cmd = new SqlCommand(strSQL, cn);
            //            cmd.Transaction = sTrans;
            //            int iDelFileCount = Convert.ToInt32(cmd.ExecuteScalar());

            //            //删除正式区标记要删除的档案D2-4
            //            if (iDelFileCount > 0)
            //            {
            //                strSQL = @"Delete from webapp.dbo.OA_Files_List where Status='暂时删除' and pageid='" + collection["inp_PageID"] + "'";
            //                cmd = new SqlCommand(strSQL, cn);
            //                cmd.Transaction = sTrans;
            //                cmd.ExecuteNonQuery();

            //                ar.Add("删除文档");
            //            }



            //            //将档案搬移到正式区D3-4
            //            strSQL = @"Insert into webapp.dbo.OA_Files_List(fkey, uuid, filename, CreateTime, CreateUser) 
            //                Select fkey, uuid, filename, CreateTime, CreateUser from webapp.dbo.OA_FilesTemp_List where pageid='" + collection["inp_PageID"] + "'";
            //            cmd = new SqlCommand(strSQL, cn);
            //            cmd.Transaction = sTrans;
            //            cmd.ExecuteNonQuery();
            //            cmd.Dispose();

            //            //将暂存区档案清除D4-4
            //            strSQL = @"Delete from webapp.dbo.OA_FilesTemp_List where pageid='" + collection["inp_PageID"] + "'";
            //            cmd = new SqlCommand(strSQL, cn);
            //            cmd.Transaction = sTrans;
            //            cmd.ExecuteNonQuery();
            //            cmd.Dispose();

            Kcis.Models.UserModel user = (Kcis.Models.UserModel)Session["UserProfile"];
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            try
            {

                cn.Open();
                //string strSQL = @"Delete from webapp.dbo.OA_Files_List where uuid='" + strUUID + "'";
                string strSQL = @"Update webapp.dbo.OA_Files_List set Status='暂时删除', UpdateTime=getdate(), PageID='" + strPageID + "' where uuid='" + strUUID + "'";
                log.Debug("-----------strSQL=" + strSQL);
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                cmd.ExecuteNonQuery();

                strSQL = @"Delete from webapp.dbo.OA_FilesTemp_List where uuid='" + strUUID + "'";
                cmd = new SqlCommand(strSQL, cn);
                cmd.ExecuteNonQuery();



                return Content("OK");
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


        }


        public ActionResult MobileFile_List(string strKey)
        {
 
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            try
            {

                cn.Open();
                //PageID不相等时 就无条件显示
                string strSQL = @"Select uuid, filename, bb.CID, SID, aa.CreateTime from  webapp.dbo.OA_Files_List aa inner join OA_SchoolActivity_ActionCourse bb on aa.FKey=bb.cid
                                    Where sid='" + strKey + "' and (isnull(status,'')<>'暂时删除' ) ";

                log.Debug("strSQL=" + strSQL);
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_File = new DataTable();
                dt_File.Load(dr);
                dr.Dispose();
                cmd.Dispose();
                string strFileLinkList = "";

                for (int i = 0; i < dt_File.Rows.Count; i++)
                {
                    string strURL = Url.Content("~/uploadfile/");
                    strFileLinkList += "<a  target='kcis' style='text-decoration:none;' href='" + strURL + dt_File.DefaultView[i]["uuid"].ToString() + "'><span  style='font-size:18px;'>" + "文档" + (i + 1) + "：" + dt_File.DefaultView[i]["filename"].ToString() + "</span></a>&nbsp;&nbsp;&nbsp;&nbsp;";
                    strFileLinkList += "<p>";
                }

                log.Debug("strShortNameList=" + strFileLinkList);
                return Content(strFileLinkList);
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


        }


    }//end of class
}
