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
    public class DormBIController : Controller 
    {
        //
        private static ILog log = LogManager.GetLogger(typeof(DormBIController)); 
        private string strFormID = "DormBI";
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

                strSQL = "SELECT YEAR(GETDATE())";
                cmd = new SqlCommand(strSQL, cn);
                int iYear = Convert.ToInt32(cmd.ExecuteScalar());
                cmd.Dispose();

                string strClassList = "<option value=''>请选择</option>";
                strClassList += "<option value='" + iYear + "'>" + iYear + "年</option>";
                strClassList += "<option value='" + ++iYear + "'>" + iYear + "年</option>";
                ViewBag.ClassList = strClassList;


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
                    throw new Kcis.Models.KcisException("送单失败，新生尚未编班无法使用此单！");





                //取出学生基本资料
                string strSQL = @" Select AccountID, fullname, cname, status, deptid, SourceType=isnull(SourceType,'')  from [Common].[dbo].kcis_account aa where AccountID=@AccountID  ";
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                cmd.Parameters.Add("@AccountID", SqlDbType.NVarChar).Value = collection["inp_AccountID"];
                SqlDataReader dataReader = cmd.ExecuteReader();
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
 

                if (!strSourceType.Equals("B"))
                    throw new Kcis.Models.KcisException("送单失败，不是小学部学生无法使用此假单!");
 

                strSQL = @" Select Count(*)  from [WebApp].[dbo].[OA_"+strFormID+"_Form] where SequenceID='"+ collection["inp_SequenceID"] + "'";
                cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                int iCount = Convert.ToInt32(cmd.ExecuteScalar());
                if(iCount>0)
                    throw new Kcis.Models.KcisException("此单[" + collection["inp_SequenceID"] + "]您已经提交成功无需重复送单!");


                //防止审批中重复提交
                strSQL = @" Select isnull(count(*),0) from Webapp.dbo.OA_DormBI_Form aa inner join Webapp.dbo.OA_Flow ff on aa.SequenceID=ff.SequenceID  where ff.Status='1' and applyID='" + user.UserId + "' and Form_Year='" + collection["sel_Form_Year"] + "' and Form_Section='" + collection["sel_Form_Section"] + "'";
                cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                iCount = Convert.ToInt32(cmd.ExecuteScalar());
                cmd.Dispose();

                if (iCount > 0)
                    throw new Kcis.Models.KcisException("送单失败，您本学期住宿申请正在审核，无需重复送单！");


                strSQL = @"INSERT INTO [WebApp].[dbo].[OA_DormBI_Form]
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
                           ,[Title] ,[FormContent]
                           ,Form_Year, Form_Section, Form_Sex
                           ,[Form_Birthday]
                           ,[Form_StudentCell]
                           ,[Form_ParentCellA]
                           ,[Form_ParentCellB]
                           ,[Form_RelationA]
                           ,[Form_RelationB]
                           ,[Form_StudentHeight]
                           ,[Form_Address]
                           )
                     VALUES
                           (@SequenceID, @FormID, @ApplyID, @ApplyName, @ApplyDeptID, @ApplyDeptName, Getdate(), @Titleid, @TitleName, CONVERT(varchar(100), GETDATE(), 111)
                           ,@FillerID, @FillerName, @FillerDeptID, @FillerDeptName, @Title, @FormContent
                           ,@Form_Year, @Form_Section, @Form_Sex
                           ,@Form_Birthday
                           ,@Form_StudentCell
                           ,@Form_ParentCellA
                           ,@Form_ParentCellB
                           ,@Form_RelationA
                           ,@Form_RelationB
                           ,@Form_StudentHeight
                           ,@Form_Address
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

                string strTitle = "申请学年:"+ collection["sel_Form_Year"];
                if (collection["sel_Form_Section"].ToString().Equals("S"))
                    strTitle += ",学期:春";
                else
                    strTitle += ",学期:秋";

                if (collection["sel_Form_Sex"].ToString().Equals("1"))
                    strTitle += ",性别:男";
                else
                    strTitle += ",性别:女"; ;

                cmd.Parameters.Add("@Title", SqlDbType.NVarChar).Value = strTitle;

                //通用区end
                //----------------------------------------客制化区块
                cmd.Parameters.Add("@FormContent", SqlDbType.NVarChar).Value = collection["text_Form_Remark"];  //请假事由
                cmd.Parameters.Add("@Form_Year", SqlDbType.NVarChar).Value = collection["sel_Form_Year"];
                cmd.Parameters.Add("@Form_Section", SqlDbType.NVarChar).Value = collection["sel_Form_Section"];
                log.Debug("sel_Form_Sex="+ collection["sel_Form_Sex"]);
                cmd.Parameters.Add("@Form_Sex", SqlDbType.NVarChar).Value = collection["sel_Form_Sex"];
                cmd.Parameters.Add("@Form_Birthday", SqlDbType.NVarChar).Value  = collection["inp_Form_Birthday"];

                cmd.Parameters.Add("@Form_StudentCell", SqlDbType.NVarChar).Value = collection["inp_Form_StudentCell"];

                cmd.Parameters.Add("@Form_ParentCellA", SqlDbType.NVarChar).Value = collection["inp_Form_ParentCellA"];  //手填日期
                cmd.Parameters.Add("@Form_ParentCellB", SqlDbType.NVarChar).Value = collection["inp_Form_ParentCellB"];

                cmd.Parameters.Add("@Form_RelationA", SqlDbType.NVarChar).Value = collection["inp_Form_RelationA"];
                cmd.Parameters.Add("@Form_RelationB", SqlDbType.NVarChar).Value = collection["inp_Form_RelationB"];
 
                cmd.Parameters.Add("@Form_StudentHeight", SqlDbType.NVarChar).Value = collection["inp_Form_StudentHeight"];
                cmd.Parameters.Add("@Form_Address", SqlDbType.NVarChar).Value = collection["text_Form_Address"];
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
                ht["@rule_id"] = "DormBI01";
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

              
                return Content("申请已提交成功，\n请到历史表单检视签核结果！(Submit form successfully!)");
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
