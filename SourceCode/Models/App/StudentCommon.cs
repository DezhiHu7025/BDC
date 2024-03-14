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

namespace Kcis.Models.App
{
    public class StudentCommon
    {
        private static ILog log = LogManager.GetLogger(typeof(StudentCommon));

        //等待付款的超过30分钟自动作废
        // new Kcis.Models.App.StudentCommon().SetBillTimeOut(cn, cmd, null, "");
        public string SetBillTimeOut(SqlConnection cn, SqlCommand cmd, SqlTransaction sTrans, string strAID)
        {
            string strMsg = "";
            try
            {

                log.Debug("------------------------Kcis.Models.Config.iAutoCancelTime=" + Kcis.Models.Config.iAutoCancelTime);
                //作用于目前还在等待付款，且本质是订单后立即付款
                string strSQL = @"Update OA_SchoolActivity_OrderList set enabled='N' from OA_SchoolActivity_Order bb Where FlowType='P' and bb.Status='P'  and OA_SchoolActivity_OrderList.OrderNO=bb.OrderNO and DATEDIFF( MI , bb.CreateTime , GETDATE())>" + Kcis.Models.Config.iAutoCancelTime;
                if (!strAID.Equals(""))
                    strSQL += " and bb.AID='"+ strAID+"'";

                log.Debug("strSql1="+strSQL);
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;
                //cmd.ExecuteNonQuery();
                cmd.Dispose();

                //作用于目前还在等待付款，且本质是订单后立即付款
                strSQL = @"Update OA_SchoolActivity_Order set Status='-100', enabled='N', Remark='待缴费时间='+CONVERT(nvarchar(10), DATEDIFF( MI , CreateTime , GETDATE()))  Where FlowType='P' and Status='P' and Checked='N' and DATEDIFF( MI , CreateTime , GETDATE())>" + Kcis.Models.Config.iAutoCancelTime;
                if (!strAID.Equals(""))
                    strSQL += " and AID='" + strAID + "'";

                log.Debug("strSql2=" + strSQL);
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null) 
                    cmd.Transaction = sTrans;
                //cmd.ExecuteNonQuery();
                cmd.Dispose();

                strMsg = "[ok]";
            }
            catch (Exception e)
            {
                log.Error(e.ToString());
                strMsg = e.Message;
            }

           
            return strMsg;
        }



    }//end of class
}