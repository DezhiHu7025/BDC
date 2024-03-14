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
    public class OpenQuestionnaireController : Controller
    {
        private static ILog log = LogManager.GetLogger(typeof(OpenQuestionnaireController));

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
 
        public ActionResult F20190601G12()
        {

            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            SqlTransaction sTrans = null;
            string stractivityID = "F20190601G12";
            try
            {

                //取申请人帐号,姓名

                cn.Open();


                string strSQL = @"Select ActivityID, Qseq, Qtype, QtypeTitle, TitleCN, TitleEN, InpType=isnull(InpType,'TEXT') from webapp.dbo.OA_ActivityQuestion where ActivityID='" + stractivityID + "' Order by Qtype, Qseq ";
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_Question = new DataTable();
                dt_Question.Load(dr);
                dr.Dispose();
                cmd.Dispose();

                int iQtype_Old = -1;
                string strQuestions = "";
                Hashtable htColor = new Hashtable();
                htColor.Add("0", "panel-success");
                htColor.Add("1", "panel-info");
                htColor.Add("2", "panel-warning");
                int iColorCount = 0;
                for (int i = 0; i < dt_Question.Rows.Count; i++)
                {
                    string strTitleCN = dt_Question.DefaultView[i]["TitleCN"].ToString();
                    string strTitleEN = dt_Question.DefaultView[i]["TitleEN"].ToString();
                    string strQtypeTitle = dt_Question.DefaultView[i]["QtypeTitle"].ToString();
                    int iSEQ = Convert.ToInt32(dt_Question.DefaultView[i]["Qseq"]);
                    int iQtype = Convert.ToInt32(dt_Question.DefaultView[i]["Qtype"]);
                    string strInpType = dt_Question.DefaultView[i]["InpType"].ToString();

                    if (iQtype_Old != iQtype)
                    {
                        if(iQtype_Old!=-1)
                            strQuestions += @"</div></div>";

                        strQuestions += @" <div class='panel @panel-Color'>
                                            <div class='panel-heading'>
		                                        <h3 class='panel-title'>@strQtypeTitle</h3>
	                                        </div>
	                                        <div class='panel-body'>".Replace("@strQtypeTitle", strQtypeTitle).Replace("@panel-Color", Convert.ToString(htColor[Convert.ToString(iColorCount++ % 3)]));
                    }
                    strQuestions += @"
                    <div class='row' style='margin:12px;'>
                        <div class='col-xs-12 col-lg-12 custRadio'  style='width:100% '>
                        <label>@TitleCN<span class='text-danger'>*</span></label>
                        @inpType
                        </div>
                    </div>".Replace("@TitleCN", strTitleCN + "<p>" + strTitleEN);


                    if (strInpType.Equals("TEXT"))
                        strQuestions = strQuestions.Replace("@inpType", "<textarea class='form-control' name='@ItemID' rows='3' style='min-width:90% '></textarea>");
                    else if (strInpType.Equals("RADIO"))
                        strQuestions = strQuestions.Replace("@inpType", @"<p><label class='radio-inline'><input type = 'radio' name ='@ItemID'   value ='5' >非常同意 Strongly agree</label >
                        <p><label class='radio-inline'><input type = 'radio' name ='@ItemID'   value ='4' >同意 Agree</label >
                        <p><label class='radio-inline'><input type = 'radio' name ='@ItemID'   value ='3' >普通 Neutral</label >
                        <p><label class='radio-inline'><input type = 'radio' name ='@ItemID'   value ='2' >不同意 Disagree</label >
                        <p><label class='radio-inline'><input type = 'radio' name ='@ItemID'   value ='1' >非常不同意 Strongly disagree</label >");


                    strQuestions = strQuestions.Replace("@ItemID", iQtype + "_" + iSEQ);

                    iQtype_Old = iQtype;

                }//end of loop

                strQuestions += @"</div></div>";

                ViewBag.ActivityID = stractivityID;
                ViewBag.Questions = strQuestions;
                //log.Debug("ViewBag.Questions="+ ViewBag.Questions);
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

        public ActionResult F20190601G3456()
        {

            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            SqlTransaction sTrans = null;
            string stractivityID = "F20190601G3456";
            try
            {

                //取申请人帐号,姓名

                cn.Open();


                string strSQL = @"Select ActivityID, Qseq, Qtype, QtypeTitle, TitleCN, TitleEN, InpType=isnull(InpType,'TEXT') from webapp.dbo.OA_ActivityQuestion where ActivityID='" + stractivityID + "' Order by Qtype, Qseq ";
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_Question = new DataTable();
                dt_Question.Load(dr);
                dr.Dispose();
                cmd.Dispose();

                int iQtype_Old = -1;
                string strQuestions = "";
                Hashtable htColor = new Hashtable();
                htColor.Add("0", "panel-success");
                htColor.Add("1", "panel-info");
                htColor.Add("2", "panel-warning");
                int iColorCount = 0;
                for (int i = 0; i < dt_Question.Rows.Count; i++)
                {
                    string strTitleCN = dt_Question.DefaultView[i]["TitleCN"].ToString();
                    string strTitleEN = dt_Question.DefaultView[i]["TitleEN"].ToString();
                    string strQtypeTitle = dt_Question.DefaultView[i]["QtypeTitle"].ToString();
                    int iSEQ = Convert.ToInt32(dt_Question.DefaultView[i]["Qseq"]);
                    int iQtype = Convert.ToInt32(dt_Question.DefaultView[i]["Qtype"]);
                    string strInpType = dt_Question.DefaultView[i]["InpType"].ToString();

                    if (iQtype_Old != iQtype)
                    {
                        if (iQtype_Old != -1)
                            strQuestions += @"</div></div>";

                        strQuestions += @" <div class='panel @panel-Color'>
                                            <div class='panel-heading'>
		                                        <h3 class='panel-title'>@strQtypeTitle</h3>
	                                        </div>
	                                        <div class='panel-body'>".Replace("@strQtypeTitle", strQtypeTitle).Replace("@panel-Color", Convert.ToString(htColor[Convert.ToString(iColorCount++ % 3)]));
                    }
                    strQuestions += @"
                    <div class='row' style='margin:12px;'>
                        <div class='col-xs-12 col-lg-12 custRadio'  style='width:100% '>
                        <label>@TitleCN<span class='text-danger'>*</span></label>
                        @inpType
                        </div>
                    </div>".Replace("@TitleCN", strTitleCN + "<p>" + strTitleEN);


                    if (strInpType.Equals("TEXT"))
                        strQuestions = strQuestions.Replace("@inpType", "<textarea class='form-control' name='@ItemID' rows='3' style='min-width:90% '></textarea>");
                    else if (strInpType.Equals("RADIO"))
                        strQuestions = strQuestions.Replace("@inpType", @"<p><label class='radio-inline'><input type = 'radio' name ='@ItemID'   value ='5' >非常同意 Strongly agree</label >
                        <p><label class='radio-inline'><input type = 'radio' name ='@ItemID'   value ='4' >同意 Agree</label >
                        <p><label class='radio-inline'><input type = 'radio' name ='@ItemID'   value ='3' >普通 Neutral</label >
                        <p><label class='radio-inline'><input type = 'radio' name ='@ItemID'   value ='2' >不同意 Disagree</label >
                        <p><label class='radio-inline'><input type = 'radio' name ='@ItemID'   value ='1' >非常不同意 Strongly disagree</label >");


                    strQuestions = strQuestions.Replace("@ItemID", iQtype + "_" + iSEQ);

                    iQtype_Old = iQtype;

                }//end of loop

                strQuestions += @"</div></div>";

                ViewBag.ActivityID = stractivityID;
                ViewBag.Questions = strQuestions;
                //log.Debug("ViewBag.Questions="+ ViewBag.Questions);
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
                if (cn != null && cn.State != ConnectionState.Closed)
                    cn.Close();

            }


            return View();
        }

        public ActionResult F20190601E()
        {

            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            SqlTransaction sTrans = null;
            string stractivityID = "F20190601E";
            try
            {

                //取申请人帐号,姓名

                cn.Open();


                string strSQL = @"Select ActivityID, Qseq, Qtype, QtypeTitle, TitleCN, TitleEN, InpType=isnull(InpType,'TEXT') from webapp.dbo.OA_ActivityQuestion where ActivityID='" + stractivityID + "' Order by Qtype, Qseq ";
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_Question = new DataTable();
                dt_Question.Load(dr);
                dr.Dispose();
                cmd.Dispose();

                int iQtype_Old = -1;
                string strQuestions = "";
                Hashtable htColor = new Hashtable();
                htColor.Add("0", "panel-success");
                htColor.Add("1", "panel-info");
                htColor.Add("2", "panel-warning");
                int iColorCount = 0;
                for (int i = 0; i < dt_Question.Rows.Count; i++)
                {
                    string strTitleCN = dt_Question.DefaultView[i]["TitleCN"].ToString();
                    string strTitleEN = dt_Question.DefaultView[i]["TitleEN"].ToString();
                    string strQtypeTitle = dt_Question.DefaultView[i]["QtypeTitle"].ToString();
                    int iSEQ = Convert.ToInt32(dt_Question.DefaultView[i]["Qseq"]);
                    int iQtype = Convert.ToInt32(dt_Question.DefaultView[i]["Qtype"]);
                    string strInpType = dt_Question.DefaultView[i]["InpType"].ToString();

                    if (iQtype_Old != iQtype)
                    {
                        if (iQtype_Old != -1)
                            strQuestions += @"</div></div>";

                        strQuestions += @" <div class='panel @panel-Color'>
                                            <div class='panel-heading'>
		                                        <h3 class='panel-title'>@strQtypeTitle</h3>
	                                        </div>
	                                        <div class='panel-body'>".Replace("@strQtypeTitle", strQtypeTitle).Replace("@panel-Color", Convert.ToString( htColor[Convert.ToString(iColorCount++ % 3)]));
                    }
                    strQuestions += @"
                    <div class='row' style='margin:12px;'>
                        <div class='col-xs-12 col-lg-12 custRadio'  style='width:100% '>
                        <label>@TitleCN<span class='text-danger'>*</span></label>
                        @inpType
                        </div>
                    </div>".Replace("@TitleCN", strTitleCN + "<p>" + strTitleEN);


                    if (strInpType.Equals("TEXT"))
                        strQuestions = strQuestions.Replace("@inpType", "<textarea class='form-control' name='@ItemID' rows='3' style='min-width:90% '></textarea>");
                    else if (strInpType.Equals("RADIO"))
                        strQuestions = strQuestions.Replace("@inpType", @"<p><label class='radio-inline'><input type = 'radio' name ='@ItemID'   value ='5' >非常同意 Strongly agree</label >
                        <p><label class='radio-inline'><input type = 'radio' name ='@ItemID'   value ='4' >同意 Agree</label >
                        <p><label class='radio-inline'><input type = 'radio' name ='@ItemID'   value ='3' >普通 Neutral</label >
                        <p><label class='radio-inline'><input type = 'radio' name ='@ItemID'   value ='2' >不同意 Disagree</label >
                        <p><label class='radio-inline'><input type = 'radio' name ='@ItemID'   value ='1' >非常不同意 Strongly disagree</label >");


                    strQuestions = strQuestions.Replace("@ItemID", iQtype + "_" + iSEQ);

                    iQtype_Old = iQtype;

                }//end of loop

                strQuestions += @"</div></div>";

                ViewBag.ActivityID = stractivityID;
                ViewBag.Questions = strQuestions;
                //log.Debug("ViewBag.Questions="+ ViewBag.Questions);
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
                if (cn != null && cn.State != ConnectionState.Closed)
                    cn.Close();

            }


            return View();
        }

        public ActionResult QuestionnaireMessage()
        {

            return View();
        }

        //send form
        public ActionResult Fill_SendFormAjax(IEnumerable<HttpPostedFileBase> files, FormCollection collection)
        {
            //Kcis.Models.UserModel user = (Kcis.Models.UserModel)Session["UserProfile"];
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            SqlTransaction sTrans = null;
            string stractivityName = collection["inp_ActivityName"].ToString(); 
            try
            {
           
                //取申请人帐号,姓名

                cn.Open();
                sTrans = cn.BeginTransaction();    //get insert script from visual management



 
                //string strSQL = @"Select deptid, cname from common.dbo.kcis_account  Where Accountid='"+ collection["inp_ApplyID"] + "' ";
                //SqlCommand cmd = new SqlCommand(strSQL, cn);
                //cmd.Transaction = sTrans;
                //SqlDataReader dr = cmd.ExecuteReader();
                //DataTable dt_Man = new DataTable();
                //dt_Man.Load(dr);
                //dr.Dispose();
                //cmd.Dispose();

                //if (dt_Man.Rows.Count!=1)
                //    throw new Kcis.Models.KcisException("查无此学号讯息，请先检查学号有无输入正确！");

                //check activity exist
                //string strSQL = @"Select * from [WebApp].[dbo].[OA_ActivityAskStudentInfo] Where ActivityName='"+ stractivityName + "' and  StudentID='" + HttpContext.Request.UserHostAddress + "' ";
                //string strSQL = @"Select * from [WebApp].[dbo].[OA_ActivityAskStudentInfo] Where ActivityName='" + stractivityName + "'";
                string strSQL = @"Select ActivityID, Qseq, Qtype, QtypeTitle, TitleCN, TitleEN from webapp.dbo.OA_ActivityQuestion where ActivityID='" + stractivityName + "' Order by Qtype, Qseq ";
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_Question = new DataTable();
                dt_Question.Load(dr);
                dr.Dispose();
                cmd.Dispose();

                //if (dt_Activity.Rows.Count >0)
                //    throw new Kcis.Models.KcisException("这问卷您已经完成，无需重复提交！");

                for (int i = 0; i < dt_Question.Rows.Count; i++)
                {
                    string strTitleCN = dt_Question.DefaultView[i]["TitleCN"].ToString();
                    string strTitleEN = dt_Question.DefaultView[i]["TitleEN"].ToString();
                    string strQtypeTitle = dt_Question.DefaultView[i]["QtypeTitle"].ToString();
                    int iSEQ = Convert.ToInt32(dt_Question.DefaultView[i]["Qseq"]);
                    int iQtype = Convert.ToInt32(dt_Question.DefaultView[i]["Qtype"]);


                    strSQL = @"INSERT INTO [WebApp].[dbo].[OA_ActivityAskStudentInfo]
                                   ([ActivityName]
                                   ,[PhoneNumber]
                                   ,[StudentID], Sex
                                   ,[Name]
                                   ,[QuestionnaireID], Qtype
                                   ,[AnswerID]
                                   ,[ClassName]
                                   ,[ClassNameEN])
                             VALUES
                                   (@ActivityName 
                                   ,''
                                   ,@StudentID, @Sex
                                   ,''
                                   ,@QuestionnaireID, @Qtype
                                   ,@AnswerID
                                   ,@ClassName, @ClassNameEN )";

                    cmd = new SqlCommand(strSQL, cn);
                    cmd.Transaction = sTrans;
                    cmd.Parameters.Add("@ActivityName", SqlDbType.NVarChar).Value = stractivityName;                         
                    cmd.Parameters.Add("@StudentID", SqlDbType.NVarChar).Value = HttpContext.Request.UserHostAddress;
                    cmd.Parameters.Add("@ClassName", SqlDbType.NVarChar).Value = Convert.ToString(collection["inp_ClassCN"]).ToUpper();
                    cmd.Parameters.Add("@ClassNameEN", SqlDbType.NVarChar).Value =Convert.ToString( collection["inp_ClassEN"]).ToUpper();
                    cmd.Parameters.Add("@Sex", SqlDbType.NVarChar).Value = collection["inp_Sex"];
                    cmd.Parameters.Add("@Qtype", SqlDbType.NVarChar).Value = iQtype;
                    cmd.Parameters.Add("@QuestionnaireID", SqlDbType.NVarChar).Value = iSEQ;                
                    cmd.Parameters.Add("@AnswerID", SqlDbType.NVarChar).Value = collection[iQtype + "_" + iSEQ].ToString();
                    

                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                }
                sTrans.Commit();
                sTrans.Dispose();

 

                return Content("问卷已提交成功，感谢！(Submit form successfully, thank you!)");
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
     


            DataTable dt_Dates = new DataTable();
            dt_Dates.Columns.Add("strStatus", typeof(string));
            dt_Dates.Columns.Add("strMessage", typeof(string));

            dt_Dates.Columns.Add("strTable", typeof(string));
            dt_Dates.Columns.Add("strList", typeof(string));


            System.Data.DataRow dRow = dt_Dates.NewRow();

        try{

            cn.Open();
            //取假单清单
            string strSQL = @"Select bb.Serialid, aa.SequenceID, ValidDate, Title='请假日期(off range)'+BeginDay1+'日'+BeginTime1+'~'+EndDay1+'日'+EndTime1,
                   status= CASE status WHEN '1' THEN '签核中' WHEN '100' THEN '已核准' WHEN '-100' THEN '已驳回' WHEN '-200' THEN '已作废' ELSE status END from (OA_LeaDay_Form aa inner join webapp.dbo.OA_Form bb
                    on aa.SequenceID=bb.SequenceID) inner join webapp.dbo.OA_Flow cc on bb.serialid=cc.serialid 
                    Where bb.Flag='Y' and (aa.ApplyID='" + user.UserId+ "' or aa.FillerID='" + user.UserId + "'  ) ";
            SqlCommand cmd = new SqlCommand(strSQL, cn);
            SqlDataReader dr = cmd.ExecuteReader();
            DataTable dt_List = new DataTable();
            dt_List.Load(dr);
            dr.Dispose();
            cmd.Dispose();

            string strTable = "";
            string strList = "";
           for (int i = 0; i < dt_List.Rows.Count; i++)
            {
                    //填大表隔
                    strTable += "<tr><td style = 'text-align:center'>" + dt_List.DefaultView[i]["SequenceID"] + "</td>";
                    strTable += "<td style = 'text -align:center'>请假单</td>";
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
            strTable+="";

            //sTrans.Commit();
            //sTrans.Dispose();
            dRow["strStatus"] = "[ok]";
            dRow["strMessage"] = "查询完毕！";
            dRow["strTable"] = strTable;
            dRow["strList"] = strList;

            dt_Dates.Rows.Add(dRow);
            strMsg = JsonConvert.SerializeObject(dt_Dates);
            return Content(strMsg);

        }catch (Common.Models.KcisException e){

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



            DataTable dt_Dates = new DataTable();
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
                DataTable dt_List = new DataTable();
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
                    strTable += "<tr><td align='center'>"+(i+1)+"</td>";
                    strTable += "<td align='left'>" + dt_List.DefaultView[i]["Signer_Name"] +"-"+ dt_List.DefaultView[i]["Signer_Title"] + "</td>";
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

    }//end of class
}
