using System;
using System.Collections.Generic;
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

namespace SourceCode.Controllers
{
    public class LeaDayController : Controller 
    {
        //
        private static ILog log = LogManager.GetLogger(typeof(LeaDayController)); 
        private string strFormID = "LeaDay";
        public ActionResult Index()
        {
            return View();
        }

        [Common.ActionFilter.CheckSessionFilter]
        public ActionResult Fill_MainPage()
        {
            Kcis.Models.UserModel user = (Kcis.Models.UserModel)Session["UserProfile"];
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
       

            try
            {
 
                cn.Open();
 

                //取得申请人归属单位名稱
                string strSQL = "Select accountid, fullname, deptid from common.dbo.kcis_account where accountid='" + user.UserId + "'";
                log.Debug("----SQL="+strSQL);
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                SqlDataReader dataReader = cmd.ExecuteReader();


                string strAccountID = "";
                string strFullName = "";
                string strDeptID = "";
                while (dataReader.Read())
                {
                    strAccountID = dataReader["accountid"].ToString();
                    strFullName = dataReader["fullname"].ToString();
                    strDeptID = dataReader["deptid"].ToString();
                }
                dataReader.Dispose();
                cmd.Dispose();


                strSQL = "Select CONVERT(varchar(100), GETDATE(), 111)";
                cmd = new SqlCommand(strSQL, cn);
                ViewBag.ToDay = Convert.ToString(cmd.ExecuteScalar());
                cmd.Dispose();


                Kcis.Models.FormModel fm = new Kcis.Models.FormModel();
                string strSequenceID = "";
                string strSerialID = "";
                string strApplyTime = "";

                fm.Fill_OA_GenerateKey(cn, cmd, ref strFormID, out strSequenceID, out strSerialID, out strApplyTime);
            
                ViewBag.SequenceID = strSequenceID;
                ViewBag.AccountID = strAccountID;
                ViewBag.FullName = strFullName;
                ViewBag.DeptID = strDeptID;
                ViewBag.SourceType = user.SourceType;

            }
            catch (Exception e){
                log.Error(e.ToString());
          
            }

            finally
            {
                if (cn.State != ConnectionState.Closed)
                    cn.Close();

            }

            return View();

        }//end of func


        [Common.ActionFilter.CheckSessionFilter]
        public ActionResult Fill_GetFullNameAjax11(string strAccountID)
        {
            Kcis.Models.UserModel user = (Kcis.Models.UserModel)Session["UserProfile"];
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);

            string strFullname = "";
            try
            {

                cn.Open();
                //取得申请人归属单位名稱
                string strSQL = "Select fullname  from common.dbo.kcis_account where accountid='" + strAccountID + "'";
                log.Debug("----SQL=" + strSQL);
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                object obj = cmd.ExecuteScalar();
                cmd.Dispose();
                
                if (obj != null && !Convert.IsDBNull(obj))
                {
                    if (Convert.ToString(obj).Length > 0)
                        strFullname = Convert.ToString(obj);
                }

            }
            catch (Exception e)
            {
                log.Error(e.ToString());

            }

            finally
            {
                if (cn.State != ConnectionState.Closed)
                    cn.Close();

            }

            return Content(strFullname);

        }//end of func


        public ActionResult Fill_GetFullNameAjax(string strAccountID)
        {


            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            string strMessage = "";


            DataTable dt_Dates = new DataTable();
            dt_Dates.Columns.Add("strStatus", typeof(string));
            dt_Dates.Columns.Add("strMessage", typeof(string));

            dt_Dates.Columns.Add("FullName", typeof(string));
            dt_Dates.Columns.Add("DeptID", typeof(string));
 

            System.Data.DataRow dRow = dt_Dates.NewRow();
            dRow["FullName"] = "";
            dRow["DeptID"] = "";
  
            try
            {
                cn.Open();
 

                string strSQL = "Select fullname, deptid  from common.dbo.kcis_account where accountid='" + strAccountID + "'";
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_Man = new DataTable();
                dt_Man.Load(dr);
                dr.Dispose();
                cmd.Dispose();
 
                if (dt_Man.Rows.Count == 1){
                    dRow["FullName"] = dt_Man.DefaultView[0]["fullname"];
                    dRow["DeptID"] = dt_Man.DefaultView[0]["deptid"];
       
                }else
                    throw new Common.Models.KcisException("错误，您填的学号对应不到姓名数据，请检查学号是否正确！");
 

                dRow["strStatus"] = "[ok]";
                dRow["strMessage"] = "";
                dt_Dates.Rows.Add(dRow);
                strMessage = JsonConvert.SerializeObject(dt_Dates);

                log.Debug("strMsg=" + strMessage);



            }
            catch (Common.Models.KcisException e)
            {  //有作用的提醒或警告
                log.Error(e.ToString());
                dRow["strStatus"] = "{error}";
                dRow["strMessage"] = e.Message;
                dt_Dates.Rows.Add(dRow);
                strMessage = JsonConvert.SerializeObject(dt_Dates);

            }
            catch (Exception e)
            {  //无作用的提报错，需资讯介入
                log.Error(e.ToString());
                dRow["strStatus"] = "[error]";
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
        }//end of func 



        [Common.ActionFilter.CheckSessionFilter]
        public ActionResult Fill_SendFormAjax(IEnumerable<HttpPostedFileBase> files, FormCollection collection)
        {
            Kcis.Models.UserModel user = (Kcis.Models.UserModel)Session["UserProfile"];
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            SqlTransaction sTrans = null;
            bool hasCommit = false;
            try
            {
                //ocn.Open();
                //取申请人帐号,姓名
  
                cn.Open();
                sTrans = cn.BeginTransaction();    //get insert script from visual management


                if (collection["inp_DeptID"].ToString().ToLower().Equals("pre"))
                    throw new Kcis.Models.KcisException("送单失败，新生尚未编班无法使用假单！");

                //取得申请人归属单位名稱
                string strSQL = @" Select aa.cname, teacherID = aa.accountid from webapp.dbo.OA_HomeTeacher_List aa inner join common.dbo.kcis_account bb
                                    on aa.Deptid = bb.deptid where bb.AccountID = '" + collection["inp_AccountID"] + "'";
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                SqlDataReader dataReader = cmd.ExecuteReader();

          
                string strTeacherID ="";
                string strTeacherName = "";
                while (dataReader.Read())
                {
                    strTeacherID = dataReader["teacherID"].ToString();
                    strTeacherName = dataReader["cname"].ToString();
                }
                dataReader.Dispose();
                cmd.Dispose();
                log.Debug("--strTeacherID=" + strTeacherID);
                if (strTeacherID.Equals(""))
                    throw new Kcis.Models.KcisException("送单失败，班级在系统尚未设置班主任！");

                log.Debug("----sel_VacationType="+ collection["sel_VacationType"]);
                if (collection["sel_VacationType"].Equals("personal")) {

                    strSQL = @"Select datediff(dd, convert(datetime,'"+ collection["inp_BeginDay1"] + "'), convert(datetime,'"+ collection["inp_EndDay1"] + "')) ";
                    cmd = new SqlCommand(strSQL, cn);
                    cmd.Transaction = sTrans;
                    int iDays = Convert.ToInt32(cmd.ExecuteScalar());
                    log.Debug("iDays="+ iDays);
                    cmd.Dispose();
                    if(iDays>7)
                        throw new Kcis.Models.KcisException("送单失败，事假最多只能请7天！(Personal leave is limited to 7 days at most)");
                }


                //取出学生基本资料
                strSQL = @" Select AccountID, fullname, cname, status, deptid, SourceType=isnull(SourceType,'')  from [Common].[dbo].kcis_account aa where AccountID=@AccountID  ";
                cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                cmd.Parameters.Add("@AccountID", SqlDbType.NVarChar).Value = collection["inp_AccountID"]; 
                dataReader = cmd.ExecuteReader();
                string strAccountID = "";
                string strFullName = "";
                string strStatus = "N";
                string strDeptID = "";
                string strSourceType = "";
                if (dataReader.Read())
                {
                    strAccountID = dataReader["AccountID"].ToString();
                    strFullName = dataReader["fullname"].ToString();
                    strStatus = dataReader["status"].ToString();
                    strDeptID = dataReader["deptid"].ToString();
                    strSourceType = dataReader["SourceType"].ToString();
                }
                dataReader.Dispose();
                cmd.Dispose();
                 
                if (!strStatus.Equals("Y"))
                    throw new Kcis.Models.KcisException("送单失败，申请人帐号异常!");

                //if (!strDeptID.Equals(collection["inp_DeptID"]))
                //    throw new Kcis.Models.KcisException("送单失败，班级数据不一致!");

                if (!strSourceType.Equals("B"))
                    throw new Kcis.Models.KcisException("送单失败，不是小学部学生无法使用此假单!");
 

                strSQL = @" Select Count(*)  from [WebApp].[dbo].[OA_"+strFormID+"_Form] where SequenceID='"+ collection["inp_SequenceID"] + "'";
                cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                int iCount = Convert.ToInt32(cmd.ExecuteScalar());
                if(iCount>0)
                    throw new Kcis.Models.KcisException("单号重复，这张假单[" + collection["inp_SequenceID"] + "]您已经提交成功无需重复送单!");


                //strSQL = @" Select Accountid from OA_HomeTeacher_List where  DType='B' and Deptid='" + user.DepID + "'";
                //cmd = new SqlCommand(strSQL, cn);
                //cmd.Transaction = sTrans;
                //object obj = cmd.ExecuteScalar();
                //cmd.Dispose();
                //string strTeacherID ="";
                //if (obj != null && !Convert.IsDBNull(obj))
                //{
                //    if (Convert.ToString(obj).Length > 0)
                //        strTeacherID = Convert.ToString(obj);

                //}


                strSQL = @"INSERT INTO [WebApp].[dbo].[OA_LeaDay_Form]
                           ([SequenceID]
                           ,[FormID]
                           ,[ApplyID]
                           ,[ApplyName]
                           ,[ApplyDeptID]
                           ,[ApplyDeptName]
                           ,[ApplyTime],[titleid],[titlename]
                           ,[ValidDate]
                           ,[FillerID]
                           ,[FillerName]
                           ,[FillerDeptID]
                           ,[FillerDeptName]
                           ,[Title], PayType, WithoutCheckCard
                           ,[ParallelSigner],[ParallelSignerName]
                           ,[VacationType]
                           ,[BeginDay1]                           
                           ,[BeginTime1]
                            ,[EndDay1]
                           ,[EndTime1]
                           ,[CalTotalDay]
                           ,[FormContent], TeacherID, Tel, CarLine
                           )
                     VALUES
                           (@SequenceID, @FormID, @ApplyID, @ApplyName, @ApplyDeptID, @ApplyDeptName, Getdate(), @Titleid, @TitleName, CONVERT(varchar(100), GETDATE(), 111)
                           ,@FillerID, @FillerName, @FillerDeptID, @FillerDeptName, @Title, @PayType, @WithoutCheckCard, @ParallelSigner, @ParallelSignerName
                           ,@VacationType, @BeginDay1, @BeginTime1, @EndDay1, @EndTime1, @CalTotalDay, @FormContent, @TeacherID, @Tel, @CarLine
                           )";

                cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                //通用区begin
                cmd.Parameters.Add("@FormID", SqlDbType.NVarChar).Value = strFormID;                        //表单编号
                cmd.Parameters.Add("@SequenceID", SqlDbType.NVarChar).Value = collection["inp_SequenceID"]; //表单物件标号
                cmd.Parameters.Add("@ApplyID", SqlDbType.NVarChar).Value = strAccountID;               //申请人讯息
                cmd.Parameters.Add("@ApplyName", SqlDbType.NVarChar).Value = strFullName;
                cmd.Parameters.Add("@ApplyDeptID", SqlDbType.NVarChar).Value = strDeptID;
                cmd.Parameters.Add("@ApplyDeptName", SqlDbType.NVarChar).Value = strDeptID;
                cmd.Parameters.Add("@Titleid", SqlDbType.NVarChar).Value = strSourceType;
                cmd.Parameters.Add("@TitleName", SqlDbType.NVarChar).Value = strSourceType;

                //cmd.Parameters.Add("@ValidDate", SqlDbType.NVarChar).Value = collection["inp_ValidDate"];   //申请日期or生效日期
                cmd.Parameters.Add("@FillerID", SqlDbType.NVarChar).Value = user.UserId;                    //填单人讯息-可能是助理
                cmd.Parameters.Add("@FillerName", SqlDbType.NVarChar).Value = user.UserName;
                cmd.Parameters.Add("@FillerDeptID", SqlDbType.NVarChar).Value = user.DepID;
                cmd.Parameters.Add("@FillerDeptName", SqlDbType.NVarChar).Value = user.DepName;
                string strTitle = "请假日期：" + collection["inp_BeginDay1"] + " " + collection["inp_BeginTime1"] + " ~ " + collection["inp_EndDay1"] + collection["inp_EndTime1"];
                cmd.Parameters.Add("@Title", SqlDbType.NVarChar).Value = strTitle;


                cmd.Parameters.Add("@ParallelSigner", SqlDbType.NVarChar).Value = "";   // matrix_ParallelSigner[0]; //会签人员讯息
                cmd.Parameters.Add("@ParallelSignerName", SqlDbType.NVarChar).Value = "";   // matrix_ParallelSigner[1];
                //通用区end

                //----------------------------------------客制化区块
                cmd.Parameters.Add("@FormContent", SqlDbType.NVarChar).Value = collection["text_Content"];  //请假事由
                cmd.Parameters.Add("@PayType", SqlDbType.NVarChar).Value = "";
                cmd.Parameters.Add("@WithoutCheckCard", SqlDbType.NVarChar).Value = "";

                cmd.Parameters.Add("@VacationType", SqlDbType.NVarChar).Value = collection["sel_VacationType"];

                cmd.Parameters.Add("@BeginDay1", SqlDbType.NVarChar).Value = collection["inp_BeginDay1"];  //手填日期
                cmd.Parameters.Add("@BeginTime1", SqlDbType.NVarChar).Value = collection["inp_BeginTime1"];

                cmd.Parameters.Add("@EndDay1", SqlDbType.NVarChar).Value = collection["inp_EndDay1"];
                cmd.Parameters.Add("@EndTime1", SqlDbType.NVarChar).Value = collection["inp_EndTime1"];

                cmd.Parameters.Add("@CalTotalDay", SqlDbType.Float).Value = 0;
                cmd.Parameters.Add("@TeacherID", SqlDbType.NVarChar).Value = strTeacherID;
                cmd.Parameters.Add("@Tel", SqlDbType.NVarChar).Value = collection["inp_Tel"];
                cmd.Parameters.Add("@CarLine", SqlDbType.NVarChar).Value = collection["inp_CarLine"];
                cmd.ExecuteNonQuery();
                cmd.Dispose();

                //Insert into [WebApp].[dbo].OA_Health_Form

                strSQL = @"Select mi =(datediff(MI, CONVERT(datetime, beginDay1+' '+BeginTime1)  , CONVERT(datetime, EndDay1+' '+EndTime1))) from [WebApp].[dbo].OA_LeaDay_Form Where SequenceID = '" + collection["inp_SequenceID"] + "'";
                cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                object obj = cmd.ExecuteScalar();
                cmd.Dispose();
                if (obj != null && !Convert.IsDBNull(obj))
                {
                    if (Convert.ToInt32(obj) <= 0)
                        throw new Kcis.Models.KcisException("时间区间有误!");
                }else
                    throw new Kcis.Models.KcisException("时间格式有误!");



                //计算请假天数
                strSQL = @"Update [WebApp].[dbo].OA_LeaDay_Form  set CalTotalDay = (datediff(DAY, CONVERT(datetime, beginDay1), CONVERT(datetime, EndDay1)) + 1)  where SequenceID = '" + collection["inp_SequenceID"] + "'";
                cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                cmd.ExecuteNonQuery();
                cmd.Dispose();

                Kcis.Models.FormModel fm = new Kcis.Models.FormModel();
                Hashtable ht = new Hashtable();
                //填写OA_Form表格
                ht["@formID"] = strFormID;
                ht["@SequenceID"] = collection["inp_SequenceID"];
                ht["@serialid"] = collection["inp_SequenceID"].ToString()+"S001";
                ht["@title"] = strTitle;
                ht["@issueid"] = user.UserId;
                ht["@issue_name"] = user.UserName;
                ht["@applyid"] = strAccountID;
                ht["@apply_name"] = strFullName;
                ht["@ApplyDeptName"] = strDeptID;
                ht["@rule_id"] = "LeaDay01";
                fm.Fill_OA_FORM(cn, cmd, sTrans, ht);

                //填写Job表格
                string strUUID = fm.Fill_OA_Job(cn, cmd, sTrans, collection["inp_SequenceID"].ToString() + "S001", user.UserId, "-1", "C", "N");  //C：送审  S：签核

                sTrans.Commit();
                sTrans.Dispose();
                hasCommit = true;


                //---接口模组begin
                if (true)
                {
                    log.Debug("Call web service...");
                    string strProcData = "http://" + Kcis.Models.Config.FMWebURL + ":" + Kcis.Models.Config.FMWeb_Port + "/flowengin?servicekey=" + collection["inp_UUID"];
                    if (Kcis.Models.Config.FMWeb_Port.Equals("80"))
                        strProcData = "http://" + Kcis.Models.Config.FMWebURL + "/flow2018engin?servicekey=" + strUUID;


                    log.Debug("strProcData=" + strProcData);
                    string strContent = Kcis.Models.Utility.UtilityIO.RequestPage(strProcData);
                    log.Debug("strContent=" + strContent);

                    if (strContent.IndexOf("[Ok]") < 0)
                    {
                        fm.Fill_OA_KeyUsed(cn, cmd, sTrans, collection["inp_SequenceID"], "Delete"); //删除回收该SequenceID
                        throw new Exception(strContent);
                    }else
                        fm.Fill_OA_KeyUsed(cn, cmd, sTrans, collection["inp_SequenceID"],"S");    //最后更新[WebApp].[dbo].OA_KeyUsed状态为Y
                }
                //---接口模组end

              
                return Content("请假申请已提交成功，\n请到历史表单检视签核结果！(Submit form successfully!)");
            }
            catch (Kcis.Models.KcisException e)
            {
                if(!hasCommit)
                    sTrans.Rollback();

                log.Error(e.ToString());
                return Content("{Error}" + e.Message);
            }
            catch (Exception e)
            {
                if (!hasCommit)
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


    }
}
