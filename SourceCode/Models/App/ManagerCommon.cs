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
    public class ManagerCommon
    {
        private static ILog log = LogManager.GetLogger(typeof(ManagerCommon));


        public void Fill_OA_GenerateKey5(SqlConnection cn, SqlCommand cmd, SqlTransaction sTrans, string strFormID, out string strSequenceID, out string strApplyTime)
        {
            try
            {
                strSequenceID = "";

                strApplyTime = "";
                int iLoppCount = 0;
                while (iLoppCount++ < 3)
                {  //run 3 times

                    //抓取Sequenceid, 今天日期
                    string strSQL = "Select SequenceID=max(SKEY), ApplyDate=CONVERT(varchar(100), GETDATE(), 111) from  [WebApp].[dbo].OA_Payment_KeyUsed Where KeyKind='" + strFormID + "_Squence' and  CONVERT(varchar(100), GETDATE(), 111)=CONVERT(varchar(100), CreateTime, 111)";
                    cmd = new SqlCommand(strSQL, cn);
                    if (sTrans != null)
                        cmd.Transaction = sTrans;
                    SqlDataReader userReader = cmd.ExecuteReader();
                    if (userReader.Read())
                    {
                        strApplyTime = userReader["ApplyDate"].ToString();
                        strSequenceID = Common.Models.Utility.UtilityString.TrimDBNull(userReader["SequenceID"]);
                        if (strSequenceID.Length == 0)
                            strSequenceID = strFormID + Common.Models.Utility.UtilityDate.trimDate(userReader["ApplyDate"].ToString()) + "00001";
                        else
                        {
                            int iTailNum = Convert.ToInt32(strSequenceID.Substring(strSequenceID.Length - 5, 5));
                            strSequenceID = strFormID + Common.Models.Utility.UtilityDate.trimDate(userReader["ApplyDate"].ToString()) + string.Format("{0:00000}", ++iTailNum); ;
                        }
                        //ViewData["strSequenceID"] = strSequenceID;
                        log.Debug("取得strSequenceID=" + strSequenceID);
                    }
                    userReader.Dispose();
                    cmd.Dispose();


                    //这里要插入数据到OA_KeyUsed作保留
                    try
                    {
                        strSQL = "Insert into [WebApp].[dbo].OA_Payment_KeyUsed(KeyKind, SKey, Status, CreateTime) Select '" + strFormID + "_Squence', @SKey, 'U', getdate()  ";  //EX:HEAL20050915001S001
                        cmd = new SqlCommand(strSQL, cn);
                        if (sTrans != null)
                            cmd.Transaction = sTrans;
                        cmd.Parameters.Add("@SKey", SqlDbType.NVarChar).Value = strSequenceID;
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

        //Just for MIS DB
        public void Fill_OA_GenerateKey3(SqlConnection cn, SqlCommand cmd, SqlTransaction sTrans, string strFormID, out string strSequenceID, out string strApplyTime)
        {
            try
            {
                strSequenceID = "";
              
                strApplyTime = "";
                int iLoppCount = 0;
                while (iLoppCount++ < 3)
                {  //run 3 times

                    //抓取Sequenceid, 今天日期
                    string strSQL = "Select SequenceID=max(SKEY), ApplyDate=CONVERT(varchar(100), GETDATE(), 111) from  [MIS].[dbo].OA_CS_KeyUsed Where KeyKind='" + strFormID + "_Squence' and  CONVERT(varchar(100), GETDATE(), 111)=CONVERT(varchar(100), CreateTime, 111)";
                    cmd = new SqlCommand(strSQL, cn);
                    if (sTrans != null)
                        cmd.Transaction = sTrans;
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

 
                    //这里要插入数据到OA_KeyUsed作保留
                    try
                    {
                        strSQL = "Insert into [MIS].[dbo].OA_CS_KeyUsed(KeyKind, SKey, Status, CreateTime) Select '" + strFormID + "_Squence', @SKey, 'U', getdate()  ";  //EX:HEAL20050915001S001
                        cmd = new SqlCommand(strSQL, cn);
                        if (sTrans != null)
                            cmd.Transaction = sTrans;
                        cmd.Parameters.Add("@SKey", SqlDbType.NVarChar).Value = strSequenceID;
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


        public void Fill_OA_GenerateKey8(SqlConnection cn, SqlCommand cmd, SqlTransaction sTrans, string strFormID, out string strSequenceID, out string strApplyTime)
        {
            try
            {
                strSequenceID = "";

                strApplyTime = "";
                int iLoppCount = 0;
                while (iLoppCount++ < 3)
                {  //run 3 times

                    //抓取Sequenceid, 今天日期
                    string strSQL = "Select SequenceID=max(SKEY), ApplyDate=CONVERT(varchar(100), GETDATE(), 111) from  [MIS].[dbo].OA_CS_KeyUsed Where KeyKind='" + strFormID + "_Squence' and  CONVERT(varchar(100), GETDATE(), 111)=CONVERT(varchar(100), CreateTime, 111)";
                    cmd = new SqlCommand(strSQL, cn);
                    if (sTrans != null)
                        cmd.Transaction = sTrans;
                    SqlDataReader userReader = cmd.ExecuteReader();
                    if (userReader.Read())
                    {
                        strApplyTime = userReader["ApplyDate"].ToString();
                        strSequenceID = Common.Models.Utility.UtilityString.TrimDBNull(userReader["SequenceID"]);
                        if (strSequenceID.Length == 0)
                            strSequenceID = strFormID + Common.Models.Utility.UtilityDate.trimDate(userReader["ApplyDate"].ToString()) + "001";
                        else
                        {
                            int iTailNum = Convert.ToInt32(strSequenceID.Substring(strSequenceID.Length -8, 8));
                            strSequenceID = strFormID + Common.Models.Utility.UtilityDate.trimDate(userReader["ApplyDate"].ToString()) + string.Format("{0:00000000}", ++iTailNum); ;
                        }
                        //ViewData["strSequenceID"] = strSequenceID;
                        log.Debug("取得strSequenceID=" + strSequenceID);
                    }
                    userReader.Dispose();
                    cmd.Dispose();


                    //这里要插入数据到OA_KeyUsed作保留
                    try
                    {
                        strSQL = "Insert into [MIS].[dbo].OA_CS_KeyUsed(KeyKind, SKey, Status, CreateTime) Select '" + strFormID + "_Squence', @SKey, 'U', getdate()  ";  //EX:HEAL20050915001S001
                        cmd = new SqlCommand(strSQL, cn);
                        if (sTrans != null)
                            cmd.Transaction = sTrans;
                        cmd.Parameters.Add("@SKey", SqlDbType.NVarChar).Value = strSequenceID;
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


        //S001取过就不再使用, 表单显示使用
        public void Fill_OA_SNO_GenerateKey(SqlConnection cn, SqlCommand cmd, SqlTransaction sTrans, string strFormID, out string strSequenceID)
        {
            try
            {
                strSequenceID = "";
       
                int iLoppCount = 0;
                while (iLoppCount++ < 3)
                {  //run 3 times

                    //抓取Sequenceid, 今天日期
                    string strSQL = "Select SequenceID=max(SKEY), ApplyDate=CONVERT(varchar(100), GETDATE(), 111) from  [WebApp].[dbo].OA_Payment_KeyUsed Where KeyKind='" + strFormID + "_Squence' and  CONVERT(varchar(100), GETDATE(), 111)=CONVERT(varchar(100), CreateTime, 111)";
                    cmd = new SqlCommand(strSQL, cn);
                    if (sTrans != null)
                        cmd.Transaction = sTrans;
                    SqlDataReader userReader = cmd.ExecuteReader();
                    if (userReader.Read())
                    {
                        string strApplyTime = userReader["ApplyDate"].ToString();
                        strSequenceID = Common.Models.Utility.UtilityString.TrimDBNull(userReader["SequenceID"]);
                        if (strSequenceID.Length == 0)
                            strSequenceID = strFormID   + "001";
                        else
                        {
                            int iTailNum = Convert.ToInt32(strSequenceID.Substring(1, 3));
                            strSequenceID = strFormID + string.Format("{0:000}", ++iTailNum); ;
                        }
 
                        log.Debug("取得strSequenceID=" + strSequenceID);
                    }
                    userReader.Dispose();
                    cmd.Dispose();

 
                    //这里要插入数据到OA_KeyUsed作保留
                    try
                    {
                        strSQL = "Insert into [WebApp].[dbo].OA_Payment_KeyUsed(KeyKind, SKey, Status, CreateTime) Select '" + strFormID + "_Squence', @SKey, 'U', getdate()  ";  //EX:HEAL20050915001S001
                        cmd = new SqlCommand(strSQL, cn);
                        if (sTrans != null)
                            cmd.Transaction = sTrans;
                        cmd.Parameters.Add("@SKey", SqlDbType.NVarChar).Value = strSequenceID;
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


        //S000001取最大值使用, insert之前取用
        public string Fill_OA_SNO6_MaxKey(SqlConnection cn, SqlCommand cmd, SqlTransaction sTrans, string strTable, string strField, string strFormID)
        {
            try
            {
                string strSequenceID = "";
 
                string strSQL = "Select SequenceID=max(" + strField + ") from  [WebApp].[dbo]." + strTable;
                cmd = new SqlCommand(strSQL, cn);
                if(sTrans!=null)
                cmd.Transaction = sTrans;
                strSequenceID = Common.Models.Utility.UtilityString.TrimDBNull(cmd.ExecuteScalar());
                cmd.Dispose();
 
                if (strSequenceID.Length == 0)
                    strSequenceID = strFormID + "000001";
                else
                {
                    int iTailNum = Convert.ToInt32(strSequenceID.Substring(strFormID.Length, 6));
                    strSequenceID = strFormID + string.Format("{0:000000}", ++iTailNum); ;
                }

               
                return strSequenceID;

            }
            catch (Exception ex)
            {
                log.Error(ex.ToString());
                throw ex;
            }

        }//end of method

        //N001取最大值使用, insert之前取用
        public string Fill_OA_SNO3_MaxKey(SqlConnection cn, SqlCommand cmd, SqlTransaction sTrans, string strTable, string strField, string strFormID)
        {
            try
            {
                string strSequenceID = "";

                string strSQL = "Select SequenceID=max(" + strField + ") from  [WebApp].[dbo]." + strTable;
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;

                strSequenceID = Common.Models.Utility.UtilityString.TrimDBNull(cmd.ExecuteScalar());
                cmd.Dispose();

                if (strSequenceID.Length == 0)
                    strSequenceID = strFormID + "001";
                else
                {
                    int iTailNum = Convert.ToInt32(strSequenceID.Substring(strFormID.Length, 3));
                    strSequenceID = strFormID + string.Format("{0:000}", ++iTailNum); ;
                }


                return strSequenceID;

            }
            catch (Exception ex)
            {
                log.Error(ex.ToString());
                throw ex;
            }

        }//end of method


        //取流水号最大值使用, insert之前取用
        public string Fill_OA_SNO3_MaxSubKey(SqlConnection cn, SqlCommand cmd, SqlTransaction sTrans, string strTable, string strMasterField, string strMasterValue, string strFormID, string strField)
        {
            try
            {
                string strSequenceID = "";

                string strSQL = "Select SequenceID=max(" + strField + ") from  [WebApp].[dbo]." + strTable + " Where " + strMasterField + "='" + strMasterValue + "'";
                log.Debug(">>>strSQL=" + strSQL);
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;

                strSequenceID = Common.Models.Utility.UtilityString.TrimDBNull(cmd.ExecuteScalar());
                cmd.Dispose();

                if (strSequenceID.Length == 0)
                    strSequenceID = strFormID + "001";
                else
                {
                    int iTailNum = Convert.ToInt32(strSequenceID.Substring(strFormID.Length, 3));
                    strSequenceID = strFormID + string.Format("{0:000}", ++iTailNum); ;
                }


                return strSequenceID;

            }
            catch (Exception ex)
            {
                log.Error(ex.ToString());
                throw ex;
            }

        }//end of method

        //显示下拉选单
        public string ShowAmountSelect(object objAmount)
        {
            
            return ShowAmountSelect(objAmount, 10);
        }
        //显示下拉选单
        public string ShowAmountSelect(object objAmount, int iMaxValue)
        {

            int iAmount = 0;
            if (objAmount != null)
                iAmount = Convert.ToInt32(objAmount);

            string strOption = "";
            for (int i = 0; i <= iMaxValue; i++)
            {
                string strCheck = "";
                if (i == iAmount)
                    strCheck = "selected='selected'";
                strOption += "<option value='"+i+"' " + strCheck + ">" + i + "</option>";
            }

            return strOption;
        }


        public string ShowAdjustSelect(object objValue, int iMaxValue)
        {

            int iAmount = 0;
            if (objValue != null)
                iAmount = Convert.ToInt32(objValue);

            string strOption = "";
            for (int i = 1; i < 4; i++)
            {
                string strCheck = "";
                if (i == iAmount)
                    strCheck = "selected='selected'";
                strOption += "<option value='" + i + "' " + strCheck + ">" + i + "</option>";
            }

            int index = 0;
            for (int i = 0; i <= iMaxValue; i++)
            {
                index = i * -1;
                string strCheck = "";
                if (index == iAmount)
                    strCheck = "selected='selected'";
                strOption += "<option value='" + index + "' " + strCheck + ">" + index + "</option>";
            }


            return strOption;
        }

        //同步更新,更新订单单头status=C/E与Purchase栏位、单身used数量。
        public void ReCaluateOrder(SqlConnection cn, SqlCommand cmd, SqlTransaction sTrans, string strOrderNO)
        {
            try
            {
                //避免采购量>订购量
                string strSQL = @"Select count(*) from OA_CST_OrderList where used>Amount and OrderNo=@OrderNo ";
                cmd = new SqlCommand(strSQL, cn);
                cmd.Parameters.Add("@OrderNo", SqlDbType.NVarChar).Value = strOrderNO;
                if (sTrans != null)
                    cmd.Transaction = sTrans;
                int iCount = Convert.ToInt32(cmd.ExecuteScalar());
                cmd.Dispose();
                if (iCount > 0)
                    throw new Kcis.Models.KcisException("该订单转出的采购量已经大于订购量，取消此次动作！");

                log.Debug("C01------------");
                //重计算OrderList各SNO已发采购数量
                strSQL = @"update OA_CST_OrderList set used = ( Select sum(pp.amount) from OA_CST_PurchaseList pp inner join dbo.OA_CST_Purchase aa on aa.purchaseno=pp.purchaseno 
                                     Where aa.status<>'N' and aa.OrderNo=@OrderNo and pp.OrderKey = OA_CST_OrderList.OrderNo+'-'+OA_CST_OrderList.SNO
                                     group by pp.OrderKey ) where OA_CST_OrderList.OrderNo=@OrderNo";
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;
                cmd.Parameters.Add("@OrderNo", SqlDbType.NVarChar).Value = strOrderNO;
                cmd.ExecuteNonQuery();
                cmd.Dispose();


                strSQL = @"update OA_CST_OrderList set used = 0 Where isnull(used,0)=0 and  OrderNo=@OrderNo";
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;
                cmd.Parameters.Add("@OrderNo", SqlDbType.NVarChar).Value = strOrderNO;
                cmd.ExecuteNonQuery();
                cmd.Dispose();

                //更新订单为E >>订单已确认
                strSQL = @"Update OA_CST_Order set status='E', ConfirmTime=getdate() , ConfirmAccount='sys'  from (Select inum=count(*) from OA_CST_OrderList where amount>used and OrderNo=@OrderNo ) C  
                            Where C.inum=0 and OA_CST_Order.OrderNo=@OrderNo  and status<>'N'";
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;
                cmd.Parameters.Add("@OrderNo", SqlDbType.NVarChar).Value = strOrderNO;
                cmd.ExecuteNonQuery();
                cmd.Dispose();

                //更新订单为F >>订单已完成
                strSQL = @"Update OA_CST_Order set status='F', ConfirmTime=getdate() , ConfirmAccount='sys'  from (Select inum=count(*) from OA_CST_Purchase where status not in('N','F') and OrderNo=@OrderNo ) C  
                            Where C.inum=0 and OA_CST_Order.OrderNo=@OrderNo  and status<>'N'";
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;
                cmd.Parameters.Add("@OrderNo", SqlDbType.NVarChar).Value = strOrderNO;
                cmd.ExecuteNonQuery();
                cmd.Dispose();

                //更新订单为C >>避免没采购单误判为完工
                strSQL = @"Update OA_CST_Order set status='C', ConfirmTime=getdate() , ConfirmAccount='sys'  from (Select inum=count(*) from OA_CST_Purchase where status not in('N') and OrderNo=@OrderNo ) C  
                            Where C.inum=0 and OA_CST_Order.OrderNo=@OrderNo and status<>'N'";
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;
                cmd.Parameters.Add("@OrderNo", SqlDbType.NVarChar).Value = strOrderNO;
                cmd.ExecuteNonQuery();
                cmd.Dispose();


                //更新订单为C >>处理中 (存在订单量>采购量>)
                strSQL = @"Update OA_CST_Order set status='C' from (Select inum=count(*) from OA_CST_OrderList where  isnull(amount,0) > isnull(used,0) and OrderNo=@OrderNo ) C  
                            Where C.inum>0 and OA_CST_Order.OrderNo=@OrderNo and status<>'N'";
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;
                cmd.Parameters.Add("@OrderNo", SqlDbType.NVarChar).Value = strOrderNO;
                cmd.ExecuteNonQuery();
                cmd.Dispose();

                //strSQL = @" select count(*) from  OA_CST_Purchase where OrderNo=@OrderNo ";
                //cmd = new SqlCommand(strSQL, cn);
                //if (sTrans != null)
                //    cmd.Transaction = sTrans;
                //cmd.Parameters.Add("@OrderNo", SqlDbType.NVarChar).Value = strOrderNO;
                //int iTotal = Convert.ToInt32(cmd.ExecuteScalar());

                //strSQL = @" select count(*) from  OA_CST_Purchase where OrderNo=@OrderNo and  status='N'";
                //cmd = new SqlCommand(strSQL, cn);
                //if (sTrans != null)
                //    cmd.Transaction = sTrans;
                //cmd.Parameters.Add("@OrderNo", SqlDbType.NVarChar).Value = strOrderNO;
                //int iTotalN = Convert.ToInt32(cmd.ExecuteScalar());

                ////当所有采购单都删除时, 订单回到可作废状态
                //if (iTotal == iTotalN) { 
                //    strSQL = @"Update OA_CST_Order set status='Y'  Where OrderNo=@OrderNo ";
                //    cmd = new SqlCommand(strSQL, cn);
                //    if (sTrans != null)
                //        cmd.Transaction = sTrans;
                //    cmd.Parameters.Add("@OrderNo", SqlDbType.NVarChar).Value = strOrderNO;
                //    cmd.ExecuteNonQuery();
                //    cmd.Dispose();

                //}


                //更新Order单头的PurchaseNO
                strSQL = @" select PurchaseNO from  OA_CST_Purchase where OrderNo=@OrderNo and status<>'N' ";
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;
                cmd.Parameters.Add("@OrderNo", SqlDbType.NVarChar).Value = strOrderNO;
                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_PurchaseNO = new DataTable();
                dt_PurchaseNO.Load(dr);
                dr.Dispose();
                cmd.Dispose();
                string strPurchaseList = "";
                for (int i = 0; i < dt_PurchaseNO.Rows.Count; i++)
                {
                    strPurchaseList += dt_PurchaseNO.DefaultView[i]["PurchaseNO"].ToString() + ",";
                }
                if (strPurchaseList.Length > 0)
                    strPurchaseList = strPurchaseList.Substring(0, strPurchaseList.Length - 1);

                strSQL = @"Update OA_CST_Order set PurchaseList=@PurchaseList  Where OrderNo=@OrderNo ";
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;
                cmd.Parameters.Add("@PurchaseList", SqlDbType.NVarChar).Value = strPurchaseList;
                cmd.Parameters.Add("@OrderNo", SqlDbType.NVarChar).Value = strOrderNO;
                cmd.ExecuteNonQuery();
                cmd.Dispose();

            }
            catch (Exception ex)
            {
                log.Error(ex.ToString());
                throw ex;
            }

        }//end of method



    }//end of class
}