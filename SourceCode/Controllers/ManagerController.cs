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
using Microsoft.Office.Interop.Word;
using System.Reflection;

namespace Kcis.Controllers
{
    public class ManagerController : Controller
    {
        private static ILog log = LogManager.GetLogger(typeof(ManagerController));

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

        //管理专区首页
        [Common.ActionFilter.CheckSessionFilter]
        public ActionResult Index()
        {

            return View();
        }

        //小学学生手册
        [Common.ActionFilter.CheckSessionFilter]
        public ActionResult XSCC()
        {


            return View();
        }


        /// <summary>
        /// Word转成Html
        /// </summary>
        /// <param name="path">要转换的文档的路径</param>
        /// <param name="savePath">转换成html的保存路径</param>
        /// <param name="wordFileName">转换成html的文件名字</param>
        public static void Word2Html(string path, string savePath, string wordFileName)
        {

            Microsoft.Office.Interop.Word.ApplicationClass word = new Microsoft.Office.Interop.Word.ApplicationClass();
            Type wordType = word.GetType();
            Microsoft.Office.Interop.Word.Documents docs = word.Documents;
            Type docsType = docs.GetType();
            Microsoft.Office.Interop.Word.Document doc = (Microsoft.Office.Interop.Word.Document)docsType.InvokeMember("Open", System.Reflection.BindingFlags.InvokeMethod, null, docs, new Object[] { (object)path, true, true });
            Type docType = doc.GetType();
            string strSaveFileName = savePath + wordFileName + ".html";
            object saveFileName = (object)strSaveFileName;
            docType.InvokeMember("SaveAs", System.Reflection.BindingFlags.InvokeMethod, null, doc, new object[] { saveFileName, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatFilteredHTML });
            docType.InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod, null, doc, null);
            wordType.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, word, null);

        }

        //
        [Common.ActionFilter.CheckSessionFilter]
        public ActionResult History_MainPage()
        {

            return View();
        }


        [Common.ActionFilter.CheckSessionFilter]
        public ActionResult History_MainList()
        {
            string strMsg = "";


            Kcis.Models.UserModel user = (Kcis.Models.UserModel)Session["UserProfile"];
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);



            System.Data.DataTable dt_Dates = new System.Data.DataTable();
            dt_Dates.Columns.Add("strStatus", typeof(string));
            dt_Dates.Columns.Add("strMessage", typeof(string));

            dt_Dates.Columns.Add("strTable", typeof(string));
            dt_Dates.Columns.Add("strList", typeof(string));


            System.Data.DataRow dRow = dt_Dates.NewRow();

            try
            {

                cn.Open();
                //取假单清单
                string strSQL = "Select * from OA_FormList xx inner join (";
                strSQL += @"Select bb.FormID, bb.Serialid, aa.SequenceID, ValidDate, Title='请假日期(off range)'+BeginDay1+'日'+BeginTime1+'~'+EndDay1+'日'+EndTime1,
                   status= CASE status WHEN '1' THEN '签核中' WHEN '100' THEN '已核准' WHEN '-100' THEN '已驳回' WHEN '-200' THEN '已作废' ELSE status END from (OA_LeaDay_Form aa inner join webapp.dbo.OA_Form bb
                    on aa.SequenceID=bb.SequenceID) inner join webapp.dbo.OA_Flow cc on bb.serialid=cc.serialid 
                    Where bb.Flag='Y' and (aa.ApplyID='" + user.UserId + "' or aa.FillerID='" + user.UserId + "'  ) ";

                strSQL += @"union Select bb.FormID, bb.Serialid, aa.SequenceID, ValidDate, aa.Title,
                   status = CASE status WHEN '1' THEN '签核中' WHEN '100' THEN '已核准' WHEN '-100' THEN '已驳回' WHEN '-200' THEN '已作废' ELSE status END from(OA_DormBI_Form aa inner join webapp.dbo.OA_Form bb
                     on aa.SequenceID= bb.SequenceID) inner join webapp.dbo.OA_Flow cc on bb.serialid = cc.serialid
                    Where bb.Flag = 'Y' and(aa.ApplyID ='" + user.UserId + "' or aa.FillerID ='" + user.UserId + "') ";

                strSQL += ") yy on xx.FormID=yy.formID";

                SqlCommand cmd = new SqlCommand(strSQL, cn);
                SqlDataReader dr = cmd.ExecuteReader();
                System.Data.DataTable dt_List = new System.Data.DataTable();
                dt_List.Load(dr);
                dr.Dispose();
                cmd.Dispose();

                string strTable = "";
                string strList = "";
                for (int i = 0; i < dt_List.Rows.Count; i++)
                {
                    //填大表隔
                    strTable += "<tr><td style = 'text-align:center'>" + dt_List.DefaultView[i]["SequenceID"] + "</td>";
                    strTable += "<td style = 'text -align:center'>" + dt_List.DefaultView[i]["FormName"] + "</td>";
                    strTable += "<td style = 'text -align:center'>" + dt_List.DefaultView[i]["ValidDate"] + "</td>";
                    strTable += "<td style = 'text -align:center'>" + dt_List.DefaultView[i]["status"] + "</td>";
                    strTable += "<td>" + dt_List.DefaultView[i]["Title"] + "</td>";
                    strTable += "<td><button type = 'button' class='btn btn-primary btn-sm' Title='" + dt_List.DefaultView[i]["Serialid"] + "' name='btn_View'>检视表单</button></td></tr>";


                    //填小清单
                    strList += "<div class='row' style='margin:10px;text-align:left'>";
                    strList += "<div class='col-12 bg-info' style='padding:5px;vertical-align:middle'><div style = 'font-weight:700;float:left' > 单号(Form number)：</div><div style = 'font-weight:700;float:left;' >" + dt_List.DefaultView[i]["SequenceID"] + "</div> &nbsp;&nbsp;&nbsp;&nbsp;<button type = 'button' class='btn btn-primary btn-sm' Title='" + dt_List.DefaultView[i]["Serialid"] + "'  name='btn_View'>检视假单(View)</button></div>";
                    strList += "<div class='col-xs-12 col-sm-6' style='padding:5px'><div style = 'font-weight:700;float:left;' > 单别(Form type)：</div>假单</div>";
                    strList += "<div class='col-xs-12 col-sm-6' style='padding:5px'><div style = 'font-weight:700;float:left;' > 填单日期(Apply Date)：</div>" + dt_List.DefaultView[i]["ValidDate"] + "</div>";
                    strList += "<div class='col-xs-12 col-sm-12' style='padding:5px;'><div style = 'font-weight:700;float:left;' > 状态(Form status)：</div>" + dt_List.DefaultView[i]["status"] + "</div>";
                    strList += "<div class='col-12' style='padding:5px;'><div style = 'font-weight:700;float:left;' > 说明：</div><div>" + dt_List.DefaultView[i]["Title"] + "</div></div>";
                    strList += "</div>";


                }//end of loop
                strTable += "";

                //sTrans.Commit();
                //sTrans.Dispose();
                dRow["strStatus"] = "[ok]";
                dRow["strMessage"] = "查询完毕！";
                dRow["strTable"] = strTable;
                dRow["strList"] = strList;

                dt_Dates.Rows.Add(dRow);
                strMsg = JsonConvert.SerializeObject(dt_Dates);
                return Content(strMsg);

            }
            catch (Common.Models.KcisException e)
            {

                //if (sTrans != null)
                //    sTrans.Rollback();
                log.Error(e.ToString());
                dRow["strStatus"] = "[alert]";
                dRow["strMessage"] = "查询失败！！" + e.Message;
                dt_Dates.Rows.Add(dRow);
                strMsg = JsonConvert.SerializeObject(dt_Dates);
                return Content(strMsg);


            }
            catch (Exception e)
            {

                //if (sTrans != null)
                //    sTrans.Rollback();
                log.Error(e.ToString());
                dRow["strStatus"] = "[error]";
                dRow["strMessage"] = "查询失败！！" + e.Message;
                dt_Dates.Rows.Add(dRow);
                strMsg = JsonConvert.SerializeObject(dt_Dates);
                return Content(strMsg);

            }

            finally
            {
                if (cn != null && cn.State != ConnectionState.Closed)
                    cn.Close();

            }//end of finally
        }


        [Common.ActionFilter.CheckSessionFilter]
        public ActionResult History_DetailAndHistorySignList(string strSerialID)
        {
            string strMsg = "";


            Kcis.Models.UserModel user = (Kcis.Models.UserModel)Session["UserProfile"];
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);



            System.Data.DataTable dt_Dates = new System.Data.DataTable();
            dt_Dates.Columns.Add("strStatus", typeof(string));
            dt_Dates.Columns.Add("strMessage", typeof(string));

            dt_Dates.Columns.Add("strTable", typeof(string));
            dt_Dates.Columns.Add("strDeatil", typeof(string));


            System.Data.DataRow dRow = dt_Dates.NewRow();

            try
            {

                cn.Open();
                //取假单清单
                string strSQL = @"Select aa.SequenceID, Tel, Carline, ValidDate, BeginTime=BeginDay1+'日'+BeginTime1+'分', EndTime=EndDay1+'日'+EndTime1+'分',
                                    VacationType= CASE VacationType WHEN 'sick' THEN '病假' WHEN 'personal' THEN '事假' WHEN 'affair' THEN '公假' WHEN 'noschool' THEN '停课' WHEN 'outsideschool' THEN '菁英课程' ELSE VacationType END,
                                    FormContent, Signer_Name, Signer_Title, sResult= CASE isnull(sResult,'') WHEN 'Y' THEN '同意' WHEN 'N' THEN '不同意'  WHEN '' THEN '' ELSE sResult END,
                                    sComment = isnull(sComment,'')
                                    from (OA_LeaDay_Form aa inner join webapp.dbo.OA_Form bb
                                    on aa.SequenceID=bb.SequenceID) LEFT join OA_LeaDay_Sign cc on bb.serialid=cc.serialid 
                                    Where bb.Flag='Y' and bb.Serialid='" + strSerialID + "' Order by arrivaldatetime ";
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                SqlDataReader dr = cmd.ExecuteReader();
                System.Data.DataTable dt_List = new System.Data.DataTable();
                dt_List.Load(dr);
                dr.Dispose();
                cmd.Dispose();

                string strTable = "";
                string strDeatil = "";
                for (int i = 0; i < dt_List.Rows.Count; i++)
                {
                    //填明细
                    if (i == 0)
                    {

                        strDeatil += "<div class='row' style='margin:10px;text-align:left'>";
                        strDeatil += "<div class='alert alert-success text-center  text-lg-center' role='alert'><strong>" + dt_List.DefaultView[i]["SequenceID"] + "</strong></div>";
                        strDeatil += "<div class='col-xs-12 col-sm-6' style='padding:5px;'><div class='' style='font-weight:700;float:left;'>单别(Form type)：</div>假单</div>";
                        strDeatil += "<div class='col-xs-12 col-sm-6' style='padding:5px;'><div class='' style='font-weight:700;float:left;'>填单日期(Apply Date)：</div>" + dt_List.DefaultView[i]["ValidDate"] + "</div>";
                        strDeatil += "<div class='col-xs-12 col-sm-6' style='padding:5px;'><div class='' style='font-weight:700;float:left;'>假别(Leave reason)：</div>" + dt_List.DefaultView[i]["VacationType"] + "</div>";

                        strDeatil += "<div class='col-xs-12 col-sm-6' style='padding:5px;'><div class='' style='font-weight:700;float:left;'>开始时间(Begin Time)：</div>" + dt_List.DefaultView[i]["BeginTime"] + "</div>";
                        strDeatil += "<div class='col-xs-12 col-sm-6' style='padding:5px;'><div class='' style='font-weight:700;float:left;'>结束时间(End Time)：</div>" + dt_List.DefaultView[i]["EndTime"] + "</div>";

                        strDeatil += "<div class='col-xs-12 col-sm-6' style='padding:5px;'><div class='' style='font-weight:700;float:left;'>联系电话(Telephone)：</div>" + dt_List.DefaultView[i]["TEL"] + "</div>";
                        strDeatil += "<div class='col-xs-12 col-sm-6' style='padding:5px;'><div class='' style='font-weight:700;float:left;'>乘车线别(Car Line)：</div>" + dt_List.DefaultView[i]["CarLine"] + "</div>";


                        strDeatil += "<div class='col-xs-12 col-sm-12' style='padding:5px;'><label>事由(Remark)</label>  <textarea class='form-control'  readonly='readonly' id='text_Content' name='text_Content' rows='3' style='min-width: 90%'>" + dt_List.DefaultView[i]["FormContent"] + "</textarea></div>";
                        strDeatil += "</div>";
                    }

                    //填签核记录
                    strTable += "<tr><td align='center'>" + (i + 1) + "</td>";
                    strTable += "<td align='left'>" + dt_List.DefaultView[i]["Signer_Name"] + "-" + dt_List.DefaultView[i]["Signer_Title"] + "</td>";
                    strTable += "<td align='left'>" + dt_List.DefaultView[i]["sResult"] + "</td>";
                    strTable += "<td align='left'>" + dt_List.DefaultView[i]["sComment"] + "</td></tr>";


                }//end of loop


                //sTrans.Commit();
                //sTrans.Dispose();
                dRow["strStatus"] = "[ok]";
                dRow["strMessage"] = "查询完毕！";
                dRow["strDeatil"] = strDeatil;
                dRow["strTable"] = strTable;


                dt_Dates.Rows.Add(dRow);
                strMsg = JsonConvert.SerializeObject(dt_Dates);
                return Content(strMsg);

            }
            catch (Common.Models.KcisException e)
            {

                //if (sTrans != null)
                //    sTrans.Rollback();
                log.Error(e.ToString());
                dRow["strStatus"] = "[alert]";
                dRow["strMessage"] = "查询失败！！" + e.Message;
                dt_Dates.Rows.Add(dRow);
                strMsg = JsonConvert.SerializeObject(dt_Dates);
                return Content(strMsg);


            }
            catch (Exception e)
            {

                //if (sTrans != null)
                //    sTrans.Rollback();
                log.Error(e.ToString());
                dRow["strStatus"] = "[error]";
                dRow["strMessage"] = "查询失败！！" + e.Message;
                dt_Dates.Rows.Add(dRow);
                strMsg = JsonConvert.SerializeObject(dt_Dates);
                return Content(strMsg);

            }

            finally
            {
                if (cn != null && cn.State != ConnectionState.Closed)
                    cn.Close();

            }//end of finally
        }


        [Common.ActionFilter.CheckSessionFilter]
        public ActionResult History_DormBIDetailAndHistorySignList(string strSerialID)
        {
            string strMsg = "";


            Kcis.Models.UserModel user = (Kcis.Models.UserModel)Session["UserProfile"];
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);



            System.Data.DataTable dt_Dates = new System.Data.DataTable();
            dt_Dates.Columns.Add("strStatus", typeof(string));
            dt_Dates.Columns.Add("strMessage", typeof(string));

            dt_Dates.Columns.Add("strTable", typeof(string));
            dt_Dates.Columns.Add("strDeatil", typeof(string));


            System.Data.DataRow dRow = dt_Dates.NewRow();

            try
            {

                cn.Open();
                //取假单清单
                string strSQL = @"Select aa.SequenceID, ApplyTime
                                    ,[Form_Year]
                                    ,Form_Section= CASE Form_Section WHEN 'F' THEN '秋' WHEN 'S' THEN '春'  ELSE Form_Section END
                                    ,[Form_Sex]= CASE Form_Sex WHEN '2' THEN '女' WHEN '1' THEN '男'  ELSE Form_Sex END
                                    ,[Form_Birthday]
                                    ,[Form_StudentCell]
                                    ,[Form_ParentCellA]
                                    ,[Form_ParentCellB]
                                    ,[Form_RelationA]
                                    ,[Form_RelationB]
                                    ,[Form_StudentHeight]
                                    ,[Form_Address]
                                    ,[Form_Remark]
                                    ,FormContent, Signer_Name, Signer_Title, sResult= CASE isnull(sResult,'') WHEN 'Y' THEN '同意' WHEN 'N' THEN '不同意'  WHEN '' THEN '' ELSE sResult END,
                                    sComment = isnull(sComment,'')
                                    from (OA_DormBI_Form aa inner join webapp.dbo.OA_Form bb
                                    on aa.SequenceID=bb.SequenceID) LEFT join OA_DormBI_Sign cc on bb.serialid=cc.serialid 
                                    Where bb.Flag='Y'  and bb.Serialid='" + strSerialID + "' Order by arrivaldatetime ";
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                SqlDataReader dr = cmd.ExecuteReader();
                System.Data.DataTable dt_List = new System.Data.DataTable();
                dt_List.Load(dr);
                dr.Dispose();
                cmd.Dispose();

                string strTable = "";
                string strDeatil = "";
                for (int i = 0; i < dt_List.Rows.Count; i++)
                {
                    //填明细
                    if (i == 0)
                    {

                        strDeatil += "<div class='row' style='margin:10px;text-align:left'>";
                        strDeatil += "<div class='alert alert-success text-center  text-lg-center' role='alert'><strong>" + dt_List.DefaultView[i]["SequenceID"] + "</strong></div>";
                        strDeatil += "<div class='col-xs-12 col-sm-6' style='padding:5px;'><div class='' style='font-weight:700;float:left;'>单别(Form type)：</div>小学住宿申请单</div>";
                        strDeatil += "<div class='col-xs-12 col-sm-6' style='padding:5px;'><div class='' style='font-weight:700;float:left;'>填单日期(Apply Date)：</div>" + dt_List.DefaultView[i]["ApplyTime"] + "</div>";
                        strDeatil += "<div class='col-xs-12 col-sm-6' style='padding:5px;'><div class='' style='font-weight:700;float:left;'>学年：</div>" + dt_List.DefaultView[i]["Form_Year"] + "</div>";

                        strDeatil += "<div class='col-xs-12 col-sm-6' style='padding:5px;'><div class='' style='font-weight:700;float:left;'>学期：</div>" + dt_List.DefaultView[i]["Form_Section"] + "</div>";
                        strDeatil += "<div class='col-xs-12 col-sm-6' style='padding:5px;'><div class='' style='font-weight:700;float:left;'>性别：</div>" + dt_List.DefaultView[i]["Form_Sex"] + "</div>";

                        strDeatil += "<div class='col-xs-12 col-sm-6' style='padding:5px;'><div class='' style='font-weight:700;float:left;'>学生电话：</div>" + dt_List.DefaultView[i]["Form_StudentCell"] + "</div>";
                        strDeatil += "<div class='col-xs-12 col-sm-6' style='padding:5px;'><div class='' style='font-weight:700;float:left;'>家长电话：</div>" + dt_List.DefaultView[i]["Form_ParentCellA"] + "</div>";
                        strDeatil += "<div class='col-xs-12 col-sm-6' style='padding:5px;'><div class='' style='font-weight:700;float:left;'>关系：</div>" + dt_List.DefaultView[i]["Form_RelationA"] + "</div>";
                        strDeatil += "<div class='col-xs-12 col-sm-6' style='padding:5px;'><div class='' style='font-weight:700;float:left;'>家长电话：</div>" + dt_List.DefaultView[i]["Form_ParentCellB"] + "</div>";
                        strDeatil += "<div class='col-xs-12 col-sm-6' style='padding:5px;'><div class='' style='font-weight:700;float:left;'>关系：</div>" + dt_List.DefaultView[i]["Form_RelationB"] + "</div>";
                        strDeatil += "<div class='col-xs-12 col-sm-6' style='padding:5px;'><div class='' style='font-weight:700;float:left;'>身高(cm)：</div>" + dt_List.DefaultView[i]["Form_StudentHeight"] + "</div>";
                        strDeatil += "<div class='col-xs-12 col-sm-6' style='padding:5px;'><div class='' style='font-weight:700;float:left;'>出生年月日：</div>" + dt_List.DefaultView[i]["Form_Birthday"] + "</div>";
                        strDeatil += "<div class='col-xs-12 col-sm-6' style='padding:5px;'><div class='' style='font-weight:700;float:left;'>家庭住址：</div>" + dt_List.DefaultView[i]["Form_Address"] + "</div>";


                        strDeatil += "<div class='col-xs-12 col-sm-12' style='padding:5px;'><label>生活习惯或特殊注记</label>  <textarea class='form-control'  readonly='readonly' id='text_Content' name='text_Content' rows='3' style='min-width: 90%'>" + dt_List.DefaultView[i]["FormContent"] + "</textarea></div>";
                        strDeatil += "</div>";
                    }

                    //填签核记录
                    strTable += "<tr><td align='center'>" + (i + 1) + "</td>";
                    strTable += "<td align='left'>" + dt_List.DefaultView[i]["Signer_Name"] + "-" + dt_List.DefaultView[i]["Signer_Title"] + "</td>";
                    strTable += "<td align='left'>" + dt_List.DefaultView[i]["sResult"] + "</td>";
                    strTable += "<td align='left'>" + dt_List.DefaultView[i]["sComment"] + "</td></tr>";


                }//end of loop


                //sTrans.Commit();
                //sTrans.Dispose();
                dRow["strStatus"] = "[ok]";
                dRow["strMessage"] = "查询完毕！";
                dRow["strDeatil"] = strDeatil;
                dRow["strTable"] = strTable;


                dt_Dates.Rows.Add(dRow);
                strMsg = JsonConvert.SerializeObject(dt_Dates);
                return Content(strMsg);

            }
            catch (Common.Models.KcisException e)
            {

                //if (sTrans != null)
                //    sTrans.Rollback();
                log.Error(e.ToString());
                dRow["strStatus"] = "[alert]";
                dRow["strMessage"] = "查询失败！！" + e.Message;
                dt_Dates.Rows.Add(dRow);
                strMsg = JsonConvert.SerializeObject(dt_Dates);
                return Content(strMsg);


            }
            catch (Exception e)
            {

                //if (sTrans != null)
                //    sTrans.Rollback();
                log.Error(e.ToString());
                dRow["strStatus"] = "[error]";
                dRow["strMessage"] = "查询失败！！" + e.Message;
                dt_Dates.Rows.Add(dRow);
                strMsg = JsonConvert.SerializeObject(dt_Dates);
                return Content(strMsg);

            }

            finally
            {
                if (cn != null && cn.State != ConnectionState.Closed)
                    cn.Close();

            }//end of finally
        }


        public static string WordToHtml(string path)
        {
            string root = AppDomain.CurrentDomain.BaseDirectory;
            //var htmlName = $"{Guid.NewGuid().ToString("N")}.html";
            //var htmlPath = root + $"Resource/Temporary/";

            var htmlName = "";
            var htmlPath = root + "";
            if (!Directory.Exists(htmlPath))
            {
                Directory.CreateDirectory(htmlPath);
            }

            ApplicationClass word = new ApplicationClass();
            Type wordType = word.GetType();
            Documents docs = word.Documents;
            Type docsType = docs.GetType();
            Document doc = (Document)docsType.InvokeMember("Open", BindingFlags.InvokeMethod, null, docs, new Object[] { (object)path, true, true });
            Type docType = doc.GetType();

            docType.InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, doc, new object[] { (htmlPath + htmlName), WdSaveFormat.wdFormatFilteredHTML });
            docType.InvokeMember("Close", BindingFlags.InvokeMethod, null, doc, null);
            wordType.InvokeMember("Quit", BindingFlags.InvokeMethod, null, word, null);

            return htmlName;
        }
    }//end of class
}
