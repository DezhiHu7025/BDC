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

namespace Kcis.Models
{

    public sealed class FireWall
    {
        private static ILog log = LogManager.GetLogger(typeof(FireWall));

        private static FireWall _instance = null;
        // Creates an syn object.
        private static readonly object SynObject = new object();
        private int iCount = 1;
        private Hashtable ht_en = new Hashtable();
        private Hashtable ht_cn = new Hashtable();
        private string strLocked = "N";
        private string strLock_Date = "";
        private string strReportStartDate = "";
        private string strReportEndDate = "";
        public int iProgress = 0;  //月结进度条  0~100
        public int iProgressStatus = 0;  //月结是否启动中  0:未启动

        FireWall()
        {
            //只有初始时执行一次
            log.Debug("---FireWall()初始化执行开始～, iCount=" + iCount);
 
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            try
            {
                log.Debug("---FireWall()初始化执行完成！");
            }
            catch (Exception e)
            {
                log.Error(e.ToString());
                 
            }
            finally
            {
                //if (cn.State != ConnectionState.Closed)
                //    cn.Close();
            }


        }//end of construtor

        //从数据库要中英翻译
        public string SystemSuspendCheck()
        {
            if (strLocked.ToUpper().Equals("Y"))
                return "系统暂停服务，目前系统财务正在结帐无法提供交易服务！";

            return "[ok]";
        }

        //关帐
        public void CloseMonthCaluateStatus()
        {
            strLocked = "Y";
        }
        //开帐
        public void OpenMonthCaluateStatus()
        {
            strLocked = "N";
        }

        public string SystemMonthCloseCheck()
        {
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            string strMsg = "";
            try
            {
                cn.Open();
                string strSQL = @"Select count(*) from OA_Payment_CalculateMaster where enddate>=CONVERT(varchar(100), GETDATE(), 111) ";
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                int icount = Convert.ToInt32( cmd.ExecuteScalar());
                if(icount==0)
                    strMsg= "[ok]";
                else
                    strMsg = "此帐务已经完成月结，无法再更新此数据！";
                
            }
            catch (Exception e)
            {
                log.Error(e.ToString());
                strMsg = e.Message;
            }
            finally
            {
                if (cn!=null && cn.State != ConnectionState.Closed)
                    cn.Close();
            }
            return strMsg;
        }

        public string SystemMonthCloseCheck(SqlConnection cn, SqlTransaction sTrans, string strFdate)
        {
            
            string strMsg = "";
            try
            {
                
                string strSQL = @"Select count(*) from OA_Payment_CalculateMaster where enddate>='" + strFdate + "' ";
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                if (sTrans!=null)
                    cmd.Transaction = sTrans;
                int icount = Convert.ToInt32(cmd.ExecuteScalar());
                if (icount == 0)
                    strMsg = "[ok]";
                else
                    strMsg = "此帐务已经完成月结，无法再更新此数据！";

            }
            catch (Exception e)
            {
                log.Error(e.ToString());
                strMsg = e.Message;
            }
            finally
            {}
            return strMsg;
        }

        public string SystemMonthCloseCheck(string strTransDate)
        {
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            string strMsg = "";
            try
            {
                cn.Open();
                string strSQL = @"Select count(*) from OA_Payment_CalculateMaster Where enddate>='"+ strTransDate +"'";
                log.Debug(">>>sql="+ strSQL);
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                int icount = Convert.ToInt32(cmd.ExecuteScalar());

                if (icount <= 0)
                    strMsg = "[ok]";
                else
                    strMsg = "此帐务已经完成月结，无法再更新此数据！";

            }
            catch (Exception e)
            {
                log.Error(e.ToString());
                strMsg = e.Message;
            }
            finally
            {
                if (cn.State != ConnectionState.Closed)
                    cn.Close();
            }

            return strMsg;
        }

        //不控管项目：国际部点心、国际部打印、国际部贴名牌、双语部打印
        public string ServiceAccountIDPromissionCheck(string strMenu1, string strAccountID)
        {
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            string strMsg = "";
            try
            {
                cn.Open();
                //string strSQL = @"Select count(*) from OA_ScanCard_Promission where accountID='" + strAccountID + "' and status='Y' ";
                string strSQL = @"Select status=isnull(status,'') from OA_ScanCard_Promission where accountID='" + strAccountID + "'  ";
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                //int icount = Convert.ToInt32(cmd.ExecuteScalar());
                object obj = cmd.ExecuteScalar();
                cmd.Dispose();
                //string strLatestMonthMessage = "";
                if (obj != null && !Convert.IsDBNull(obj))
                {
                    string strStatus = Convert.ToString(obj);
                    if (strStatus.ToUpper().Equals("Y"))
                        strMsg = "[ok]";
                    else
                        strMsg = "该学号由于退转学申请已经被财务停用消费权限！";
                }else
                    strMsg = "该学号无家长同意书！";

            }
            catch (Exception e)
            {
                log.Error(e.ToString());
                strMsg = e.Message;
            }
            finally
            {
                if (cn != null && cn.State != ConnectionState.Closed)
                    cn.Close();
            }
            return strMsg;
        }
 
        public static FireWall Instance
        {
            get
            {
                // Double-Checked Locking
                if (null == _instance)
                {
                    lock (SynObject)
                    {
                        if (null == _instance)
                        {
                            _instance = new FireWall();

                        }//end of if
                    }
                }
                return _instance;
            }
        }//end of sttaic method
    }//end of class
}