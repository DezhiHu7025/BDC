using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel; //ref NPOI.OOXML + OpenXml4
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using System.Data.OracleClient;

using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Text;
using System.Collections;
using log4net;
using Common.Models.Utility;

namespace Kcis.Models
{
 

    public class CommonFuction
    {
        private static ILog log = LogManager.GetLogger(typeof(CommonFuction));

        public void SendMail(SqlConnection cn_EIP, SqlTransaction sTrans, UserModel Source_User, UserModel Target_User, StringBuilder strBody, string strSubject)
        {


            string strSQL = @"Insert into [Common].[dbo].[oa_emaillog](pid,emailid ,actiontype, fromaddr ,fromname, toaddr, toname, subject, body , attch , remark ,createdate ) 
                        values(@pid, @emailid , @actiontype , @fromaddr, @fromname, @toaddr, @toname, @subject, @body , @attch , @remark, getdate())";

            if (!Kcis.Models.Config.IsEmail)
            {
                Target_User.Email = "sam_cheng@kcisec.com";
                Target_User.UserName = "郑宇修";
                strSubject = "(测试邮件)" + strSubject;
            }

            if (Source_User == null)
            {
                log.Debug("------------------S1");
                Source_User = new UserModel();
                Source_User.Email = Kcis.Models.Config.client_FromAddr;
                Source_User.UserName =  Kcis.Models.Config.client_FromName;
            }

            SqlCommand cmd = new SqlCommand(strSQL, cn_EIP);
            if (sTrans != null)
                cmd.Transaction = sTrans;
            cmd.Parameters.Add("@pid", SqlDbType.NVarChar).Value = Kcis.Models.Config.WebURL;
            cmd.Parameters.Add("@emailid", SqlDbType.NVarChar).Value = Convert.ToString(System.Guid.NewGuid());
            cmd.Parameters.Add("@actiontype", SqlDbType.NVarChar).Value = "email";
            cmd.Parameters.Add("@fromaddr", SqlDbType.NVarChar).Value = Source_User.Email;
            cmd.Parameters.Add("@fromname", SqlDbType.NVarChar).Value = Source_User.UserName;
            cmd.Parameters.Add("@toaddr", SqlDbType.NVarChar).Value = Target_User.Email;
            cmd.Parameters.Add("@toname", SqlDbType.NVarChar).Value = Target_User.UserName;
            cmd.Parameters.Add("@subject", SqlDbType.NVarChar).Value = strSubject;
            cmd.Parameters.Add("@body", SqlDbType.NVarChar).Value = strBody.ToString();
            cmd.Parameters.Add("@attch", SqlDbType.NVarChar).Value = "";
            cmd.Parameters.Add("@remark", SqlDbType.NVarChar).Value = "";
            cmd.ExecuteNonQuery();
            cmd.Dispose();

        }//end of func



    }//end of class
         
}