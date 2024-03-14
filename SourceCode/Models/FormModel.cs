using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data.SqlClient;
using System.Configuration;
using log4net;
using System.Data;
using System.Text;
using System.Collections;
using System.Globalization;
using System.Data.OracleClient;

namespace Kcis.Models
{

    /*送单时共用function
     * Create by sam
     * 2015/10/9
     * 
     */
    public class FormModel
    {
 

        private static ILog log = LogManager.GetLogger(typeof(FormModel));

        public void Fill_OA_GenerateKey(SqlConnection cn, SqlCommand cmd, ref string strFormID, out string strSequenceID, out string strSerialID, out string strApplyTime)
        {
            try
            {
                strSequenceID = "";
                strSerialID = "";
                strApplyTime = "";
                int iLoppCount = 0;
                while (iLoppCount++ < 3)
                {  //run 3 times

                    //抓取Sequenceid, 今天日期
                    string strSQL = "Select SequenceID=max(SKEY), ApplyDate=CONVERT(varchar(100), GETDATE(), 111) from  [WebApp].[dbo].OA_KeyUsed Where KeyKind='" + strFormID + "_Squence' and  CONVERT(varchar(100), GETDATE(), 111)=CONVERT(varchar(100), CreateTime, 111)";
                    cmd = new SqlCommand(strSQL, cn);
                    SqlDataReader userReader = cmd.ExecuteReader();
                    if (userReader.Read())
                    {
                        strApplyTime = userReader["ApplyDate"].ToString();
                        strSequenceID = Common.Models.Utility.UtilityString.TrimDBNull(userReader["SequenceID"]);
                        if (strSequenceID.Length == 0)
                            strSequenceID = strFormID + Common.Models.Utility.UtilityDate.trimDate(userReader["ApplyDate"].ToString()) + "001";
                        else
                        {
                            int iTailNum = Convert.ToInt32(strSequenceID.Substring(strSequenceID.Length - 3, 3));
                            strSequenceID = strFormID + Common.Models.Utility.UtilityDate.trimDate(userReader["ApplyDate"].ToString()) + string.Format("{0:000}", ++iTailNum); ;
                        }
                        //ViewData["strSequenceID"] = strSequenceID;
                        log.Debug("取得strSequenceID=" + strSequenceID);
                    }
                    userReader.Dispose();
                    cmd.Dispose();



                    //取用新的SerialID, 不考虑同时被占用问题
                    strSQL = "Select SerialID = MAX(serialid) from [WebApp].[dbo].OA_Form where formid='" + strSequenceID + "'";  //EX:HEAL20050915001S001
                    cmd = new SqlCommand(strSQL, cn);
                    userReader = cmd.ExecuteReader();
                    if (userReader.Read())
                    {

                        strSerialID = Common.Models.Utility.UtilityString.TrimDBNull(userReader["SerialID"]);
                        if (strSerialID.Length == 0)
                            strSerialID = strSequenceID + "S001";
                        else
                        {
                            int iTailNum = Convert.ToInt32(strSerialID.Substring(strSerialID.Length - 3, 3));
                            strSerialID = strSequenceID + "S" + string.Format("{0:000}", ++iTailNum);
                        }
                        //ViewData["strSerialID"] = strSerialID;
                    }
                    userReader.Dispose();
                    cmd.Dispose();


                    //这里要插入数据到OA_KeyUsed作保留
                    try
                    {
                        strSQL = "Insert into [WebApp].[dbo].OA_KeyUsed(KeyKind, SKey, Status, CreateTime) Select '" + strFormID + "_Squence', @SKey, 'U', getdate()  ";  //EX:HEAL20050915001S001
                        cmd = new SqlCommand(strSQL, cn);
                        cmd.Parameters.Add("@SKey", SqlDbType.VarChar).Value = strSequenceID;
                        cmd.ExecuteNonQuery();
                        iLoppCount = 3;
                    }
                    catch (SqlException se)
                    {
                        log.Error(se.ToString());
                        continue;  //SequenceID， SerialID重新产生
                    }

                }// end of loop



            }
            catch (Exception ex)
            {
                log.Error(ex.ToString());
                throw ex;
            }

        }//end of method


        public void Fill_OA_FORM(SqlConnection cn, SqlCommand cmd, SqlTransaction sTrans, Hashtable ht)
        {
            try
            {


                string strSQL = @"INSERT INTO [WebApp].[dbo].[OA_Form]
                           ([formID],[SequenceID]
                           ,[serialid]
                           ,[title],[title_en]
                           ,[issueid]
                           ,[issue_name]
                           ,[applyid]
                           ,[apply_name]
                           ,[rule_id]
                           ,[create_date])
                     VALUES
                           (@formID, @SequenceID, @serialid, @title, @title_en, @issueid, @issue_name, @applyid, @apply_name, @rule_id, getdate())";

                cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                cmd.Parameters.Add("@formID", SqlDbType.NVarChar).Value = ht["@formID"];
                cmd.Parameters.Add("@SequenceID", SqlDbType.NVarChar).Value = ht["@SequenceID"];
                cmd.Parameters.Add("@serialid", SqlDbType.NVarChar).Value = ht["@serialid"];
                cmd.Parameters.Add("@title", SqlDbType.NVarChar).Value = ht["@title"];
                if (ht["@title_en"] == null || ht["@title_en"].ToString().Equals(""))
                    ht["@title_en"] = ht["@title"];
                cmd.Parameters.Add("@title_en", SqlDbType.NVarChar).Value = ht["@title_en"];
                cmd.Parameters.Add("@issueid", SqlDbType.NVarChar).Value = ht["@issueid"];
                cmd.Parameters.Add("@issue_name", SqlDbType.NVarChar).Value = ht["@issue_name"];
                cmd.Parameters.Add("@applyid", SqlDbType.NVarChar).Value = ht["@applyid"];
                cmd.Parameters.Add("@apply_name", SqlDbType.NVarChar).Value = ht["@apply_name"];
                cmd.Parameters.Add("@ApplyDeptName", SqlDbType.NVarChar).Value = ht["@ApplyDeptName"];
                cmd.Parameters.Add("@rule_id", SqlDbType.NVarChar).Value = ht["@rule_id"];
                cmd.ExecuteNonQuery();
                cmd.Dispose();



            }
            catch (Exception ex)
            {
                log.Error(ex.ToString());
                throw ex;
            }

        }//end of method

        public void Fill_OA_FORM(SqlConnection cn, SqlCommand cmd, SqlTransaction sTrans, FormCollection collection, Kcis.Models.UserModel user, string strFormID, string strTitle, string strTitle_EN, string strApplyid, string strApply_name, string strApplyDeptName, string strRuleid)
        {
            try
            {


                string strSQL = @"INSERT INTO [WebApp].[dbo].[OA_Form]
                           ([formID],[SequenceID]
                           ,[serialid]
                           ,[title],[title_en]
                           ,[issueid]
                           ,[issue_name]
                           ,[applyid]
                           ,[apply_name]
                           ,[rule_id]
                           ,[create_date])
                     VALUES
                           (@formID, @SequenceID, @serialid, @title, @title_en, @issueid, @issue_name, @applyid, @apply_name, @rule_id, getdate())";

                cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                cmd.Parameters.Add("@formID", SqlDbType.NVarChar).Value = strFormID;
                cmd.Parameters.Add("@SequenceID", SqlDbType.NVarChar).Value = collection["inp_SequenceID"];
                cmd.Parameters.Add("@serialid", SqlDbType.NVarChar).Value = collection["inp_SerialID"];
                cmd.Parameters.Add("@title", SqlDbType.NVarChar).Value = strTitle;
                if (strTitle_EN == null || strTitle_EN.Equals(""))
                    strTitle_EN = strTitle;
                cmd.Parameters.Add("@title_en", SqlDbType.NVarChar).Value = strTitle_EN;
                cmd.Parameters.Add("@issueid", SqlDbType.NVarChar).Value = user.UserId;
                cmd.Parameters.Add("@issue_name", SqlDbType.NVarChar).Value = user.UserName;
                cmd.Parameters.Add("@applyid", SqlDbType.NVarChar).Value = strApplyid;
                cmd.Parameters.Add("@apply_name", SqlDbType.NVarChar).Value = strApply_name;
                cmd.Parameters.Add("@ApplyDeptName", SqlDbType.NVarChar).Value = strApplyDeptName;
                cmd.Parameters.Add("@rule_id", SqlDbType.NVarChar).Value = strRuleid;
                cmd.ExecuteNonQuery();
                cmd.Dispose();



            }
            catch (Exception ex)
            {
                log.Error(ex.ToString());
                throw ex;
            }

        }//end of method


        public string Fill_OA_Job(SqlConnection cn, SqlCommand cmd, SqlTransaction sTrans, string strSerialID, string IssueID, string strStep, string strAction, string strFlag)
        {
            string strUUID = Convert.ToString(System.Guid.NewGuid());
            try
            {


                string strSQL = @"INSERT INTO [WebApp].[dbo].[OA_Job]
                           ([uuid],[SerialID]
                           ,[HandleTime]
                           ,[Step]
                           ,[SenderID]
                           ,[Action]
                           ,[Flag])
                     VALUES
                           (@uuid, @SerialID, getdate(), @Step, @SenderID, @Action, @Flag)";
                cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                cmd.Parameters.Add("@uuid", SqlDbType.NVarChar).Value = strUUID;
                cmd.Parameters.Add("@SerialID", SqlDbType.NVarChar).Value = strSerialID;
                cmd.Parameters.Add("@Step", SqlDbType.NVarChar).Value = strStep;
                cmd.Parameters.Add("@SenderID", SqlDbType.NVarChar).Value = IssueID;
                cmd.Parameters.Add("@Action", SqlDbType.NVarChar).Value = strAction;
                cmd.Parameters.Add("@Flag", SqlDbType.NVarChar).Value = strFlag;
                cmd.ExecuteNonQuery();
                cmd.Dispose();



            }
            catch (Exception ex)
            {
                log.Error(ex.ToString());
                throw new Exception("表单已经送出,请避免重复送单！");
            }
            return strUUID;
        }//end of method


        public void Fill_OA_KeyUsed(SqlConnection cn, SqlCommand cmd, SqlTransaction sTrans, string strSequenceID, string strMode)
        {
            try
            {
                string strSQL = @"Delete from [WebApp].[dbo].[OA_KeyUsed] Where SKEY = @SKEY ";
                if (strMode.ToUpper().Equals("S"))
                    strSQL = @"Update [WebApp].[dbo].[OA_KeyUsed] set [Status]='Y' Where SKEY = @SKEY ";
           

                cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                cmd.Parameters.Add("@SKEY", SqlDbType.NVarChar).Value = strSequenceID;

                cmd.ExecuteNonQuery();
                cmd.Dispose();


            }
            catch (Exception ex)
            {
                log.Error(ex.ToString());
                throw ex;
            }

        }//end of method

        public void GetNewFormBaseParameters(SqlConnection cn, SqlTransaction sTrans, string strApplyID, string strApplyDeptID, out string strApplyDeptName, out string strTitleID, out string strTitleName, out int iMLevel)
        {

            try
            {

                string strSQL = "Select deptname from [Common].[dbo].kcis_dept where deptid='" + strApplyDeptID + "'";
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                strApplyDeptName = Convert.ToString(cmd.ExecuteScalar());



                //取出applyid deptid對應的titleid, 职级
                strTitleID = "";
                strTitleName = "";
                iMLevel = -1;
                strSQL = @" Select isnull(titleid,'') as titleid, titleName, gradelevel as mlevel  from [Common].[dbo].kcis_account aa where AccountID=@AccountID and deptid=@DEPTID
                                    union
                                   Select isnull(aa.titleid,'') as titleid, title_name as titleName, degree as mlevel from [Common].[dbo].kcis_member aa left join Common.dbo.kcis_title bb on aa.titleid = bb.titleid where AccountID=@AccountID and deptid=@DEPTID";

                cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                cmd.Parameters.Add("@AccountID", SqlDbType.NVarChar).Value = strApplyID;
                cmd.Parameters.Add("@DEPTID", SqlDbType.NVarChar).Value = strApplyDeptID;
                SqlDataReader dataReader = cmd.ExecuteReader();

                if (dataReader.Read())
                {
                    strTitleID = dataReader["titleid"].ToString();
                    strTitleName = dataReader["titleName"].ToString();
                    iMLevel = Convert.ToInt32(dataReader["mlevel"]);
                }
                dataReader.Dispose();
                cmd.Dispose();
                if (strTitleID.Equals(""))
                    throw new Exception("申請人沒有對應的職稱!");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //抓取假别代号
      
        //抓取假别代号
     
        public bool CheckCloseMonth(OracleConnection ocn, string strCheckDate)
        {
            //2017/02/01
            strCheckDate = strCheckDate.Replace("/", "-").Substring(0, 7);
            log.Debug("strCheckDate=" + strCheckDate);
            //检查系统是否关帐
            bool bCloseMonth = false;
            string strSQL = @"SELECT count(*) from HR_MONTHCLOSE A  WHERE to_char(A.YYMM, 'yyyy-mm') ='" + strCheckDate + "' and A.Seg_Segment_No = '4343'";


            OracleCommand ocmd = new OracleCommand(strSQL, ocn);
            if (Convert.ToInt32(ocmd.ExecuteScalar()) != 0)
            {

                strSQL = @"SELECT A.CLOSEFLAG from HR_MONTHCLOSE A WHERE to_char(A.YYMM, 'yyyy-mm') ='" + strCheckDate + "' and A.Seg_Segment_No = '4343'";
                log.Debug("检查系统是否关帐SQL:" + strSQL);
                string str = Convert.ToString(ocmd.ExecuteScalar());
                log.Debug("检查系统是否关帐返回值:" + str);
                if (str.Equals("Y"))
                    bCloseMonth = true;
            }

            //if (bCloseMonth)
            //    throw new Exception("该月份已经人事系统已经关帐无法再做异动!");
            return bCloseMonth;

        }

    }//end of class
}