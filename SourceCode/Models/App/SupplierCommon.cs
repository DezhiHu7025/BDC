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
    public class SupplierCommon
    {

        private static ILog log = LogManager.GetLogger(typeof(SupplierCommon));

        //同步更新,更新订单单头status=C/E与Purchase栏位、单身used数量。
        public void ReCaluatePurchase(SqlConnection cn, SqlCommand cmd, SqlTransaction sTrans, string strPurchase)
        {
            try
            {
                //避免出货量>采购量
                string strSQL = @"Select count(*) from OA_CST_PurchaseList where Delivered>Amount and PurchaseNo=@PurchaseNo ";
                cmd = new SqlCommand(strSQL, cn);
                cmd.Parameters.Add("@PurchaseNo", SqlDbType.NVarChar).Value = strPurchase;
                if (sTrans != null)
                    cmd.Transaction = sTrans;
                int iCount = Convert.ToInt32(cmd.ExecuteScalar());
                cmd.Dispose();
                if (iCount > 0)
                    throw new Kcis.Models.KcisException("该转出的出货量已经大于采购量，取消此次动作！");

                log.Debug("C01------------");
                //重计算OA_CST_ReceiptList各SNO已发出货数量
                strSQL = @"update OA_CST_PurchaseList set Delivered = isnull(( Select sum(pp.amount) from OA_CST_ReceiptList pp inner join dbo.OA_CST_Receipt aa on aa.ReceiptNO=pp.ReceiptNO 
                                     Where aa.status not in('Y','N') and aa.PurchaseNO=OA_CST_PurchaseList.PurchaseNo and pp.PurchaseKey = OA_CST_PurchaseList.PurchaseNO+'-'+OA_CST_PurchaseList.SNO
                                     group by pp.PurchaseKey ),0) where OA_CST_PurchaseList.PurchaseNo=@PurchaseNo";
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;
                cmd.Parameters.Add("@PurchaseNo", SqlDbType.NVarChar).Value = strPurchase;
                cmd.ExecuteNonQuery();
                cmd.Dispose();

                //计算未确认之出货量
                strSQL = @"update OA_CST_PurchaseList set UnConfirmNum = isnull(( Select sum(pp.amount) from OA_CST_ReceiptList pp inner join dbo.OA_CST_Receipt aa on aa.ReceiptNO=pp.ReceiptNO 
                                     Where aa.status ='Y' and aa.PurchaseNO=OA_CST_PurchaseList.PurchaseNo and pp.PurchaseKey = OA_CST_PurchaseList.PurchaseNO+'-'+OA_CST_PurchaseList.SNO
                                     group by pp.PurchaseKey ),0) where OA_CST_PurchaseList.PurchaseNo=@PurchaseNo";
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;
                cmd.Parameters.Add("@PurchaseNo", SqlDbType.NVarChar).Value = strPurchase;
                cmd.ExecuteNonQuery();
                cmd.Dispose();

                //计算未确认之退货量
                strSQL = @"update OA_CST_PurchaseList set UnConfirmCancelNum = isnull(( Select sum(pp.amount) from OA_CST_ReturnList pp inner join dbo.OA_CST_Return aa on aa.ReturnNO=pp.ReturnNO 
                                     Where aa.status ='Y' and Stype='D' and aa.PurchaseNO=OA_CST_PurchaseList.PurchaseNo and pp.PurchaseKey = OA_CST_PurchaseList.PurchaseNO+'-'+OA_CST_PurchaseList.SNO
                                     group by pp.PurchaseKey ),0) where OA_CST_PurchaseList.PurchaseNo=@PurchaseNo";
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;
                cmd.Parameters.Add("@PurchaseNo", SqlDbType.NVarChar).Value = strPurchase;
                cmd.ExecuteNonQuery();
                cmd.Dispose();


                //计算未确认之换货量
                strSQL = @"update OA_CST_PurchaseList set UnConfirmChangeNum = isnull(( Select sum(pp.amount) from OA_CST_ReturnList pp inner join dbo.OA_CST_Return aa on aa.ReturnNO=pp.ReturnNO 
                                     Where aa.status ='Y' and Stype='F' and aa.PurchaseNO=OA_CST_PurchaseList.PurchaseNo and pp.PurchaseKey = OA_CST_PurchaseList.PurchaseNO+'-'+OA_CST_PurchaseList.SNO
                                     group by pp.PurchaseKey ),0) where OA_CST_PurchaseList.PurchaseNo=@PurchaseNo";
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;
                cmd.Parameters.Add("@PurchaseNo", SqlDbType.NVarChar).Value = strPurchase;
                cmd.ExecuteNonQuery();
                cmd.Dispose();


                //计算已确认之退货量
                strSQL = @"update OA_CST_PurchaseList set CancelAmount = isnull(( Select sum(pp.amount) from OA_CST_ReturnList pp inner join dbo.OA_CST_Return aa on aa.ReturnNO=pp.ReturnNO 
                                     Where aa.status ='C' and Stype='D' and aa.PurchaseNO=OA_CST_PurchaseList.PurchaseNo and pp.PurchaseKey = OA_CST_PurchaseList.PurchaseNO+'-'+OA_CST_PurchaseList.SNO
                                     group by pp.PurchaseKey ),0) where OA_CST_PurchaseList.PurchaseNo=@PurchaseNo";
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;
                cmd.Parameters.Add("@PurchaseNo", SqlDbType.NVarChar).Value = strPurchase;
                cmd.ExecuteNonQuery();
                cmd.Dispose();

                //更新出货量Delivered (换货＝退回出货并将发生新的出货单)
                strSQL = @"update OA_CST_PurchaseList set Delivered = 
	                         isnull(( Select sum(pp.amount) from OA_CST_ReceiptList pp inner join dbo.OA_CST_Receipt aa on aa.ReceiptNO=pp.ReceiptNO 
                             Where aa.status ='C' and aa.PurchaseNO=OA_CST_PurchaseList.PurchaseNo and pp.PurchaseKey = OA_CST_PurchaseList.PurchaseNO+'-'+OA_CST_PurchaseList.SNO
                             group by pp.PurchaseKey ),0)
                             - isnull(( Select sum(pp.amount) from OA_CST_ReturnList pp inner join dbo.OA_CST_Return aa on aa.ReturnNO=pp.ReturnNO 
                             Where aa.status ='C' and Stype='D' and aa.PurchaseNO=OA_CST_PurchaseList.PurchaseNo and pp.PurchaseKey = OA_CST_PurchaseList.PurchaseNO+'-'+OA_CST_PurchaseList.SNO
                             group by pp.PurchaseKey ),0)
                             - isnull(( Select sum(pp.amount) from OA_CST_ReturnList pp inner join dbo.OA_CST_Return aa on aa.ReturnNO=pp.ReturnNO 
                             Where aa.status ='C' and Stype='F' and aa.PurchaseNO=OA_CST_PurchaseList.PurchaseNo and pp.PurchaseKey = OA_CST_PurchaseList.PurchaseNO+'-'+OA_CST_PurchaseList.SNO
                             group by pp.PurchaseKey ),0)
                             where OA_CST_PurchaseList.PurchaseNo=@PurchaseNo";
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;
                cmd.Parameters.Add("@PurchaseNo", SqlDbType.NVarChar).Value = strPurchase;
                cmd.ExecuteNonQuery();
                cmd.Dispose();

                strSQL = @"Select count(*) from OA_CST_PurchaseList Where PurchaseNo=@PurchaseNo and Delivered<0";
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;
                cmd.Parameters.Add("@PurchaseNo", SqlDbType.NVarChar).Value = strPurchase;
                if(Convert.ToInt32(cmd.ExecuteScalar())>0)
                    //throw new Exception("采购单号：" + strPurchase + "发生出货量<0, 作业失败！");
                    throw new Exception("作业失败, 因为已经打了退换货单，无法再作废此单据！");
                cmd.Dispose();


                //防止总计栏位出现null
                strSQL = @"update OA_CST_PurchaseList set Delivered=isnull(Delivered,0),CancelAmount=isnull(CancelAmount,0),UnConfirmNum=isnull(UnConfirmNum,0),UnConfirmCancelNum=isnull(UnConfirmCancelNum,0),UnConfirmChangeNum=isnull(UnConfirmChangeNum,0) Where PurchaseNo=@PurchaseNo";
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;
                cmd.Parameters.Add("@PurchaseNo", SqlDbType.NVarChar).Value = strPurchase;
                cmd.ExecuteNonQuery();
                cmd.Dispose();

               
                //更新采购单为F >>采购单结案 （未确认出货单视同未生效）
                strSQL = @"Update OA_CST_Purchase set status='F' from (Select inum=count(*) from OA_CST_PurchaseList where (amount-CancelAmount)>Delivered and PurchaseNo=@PurchaseNo ) C  
                            Where C.inum=0 and OA_CST_Purchase.PurchaseNO=@PurchaseNO  and  OA_CST_Purchase.status not in('N','Y')";
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;
                cmd.Parameters.Add("@PurchaseNo", SqlDbType.NVarChar).Value = strPurchase;
                cmd.ExecuteNonQuery();
                cmd.Dispose();

                //有作废动作时, 复原为C
                strSQL = @"Update OA_CST_Purchase set status='C' from (Select inum=count(*) from OA_CST_PurchaseList where (amount-CancelAmount)>Delivered and PurchaseNo=@PurchaseNo ) C  
                            Where C.inum>0 and OA_CST_Purchase.PurchaseNO=@PurchaseNO  and  OA_CST_Purchase.status not in('N','Y')";
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;
                cmd.Parameters.Add("@PurchaseNo", SqlDbType.NVarChar).Value = strPurchase;
                cmd.ExecuteNonQuery();
                cmd.Dispose();


                #region 取所有出货／退货清单单号到 OA_CST_Purchase 采购主档
                strSQL = @" select ReceiptNO from  OA_CST_Receipt where status<>'N' and PurchaseNO=@PurchaseNO ";
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;

                cmd.Parameters.Add("@PurchaseNO", SqlDbType.NVarChar).Value = strPurchase;
                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_Receipt = new DataTable();
                dt_Receipt.Load(dr);
                dr.Dispose();
                cmd.Dispose();
                string strReceiptList = "";
                for (int i = 0; i < dt_Receipt.Rows.Count; i++)
                {
                    strReceiptList += dt_Receipt.DefaultView[i]["ReceiptNO"].ToString() + ",";
                }
                if (strReceiptList.Length > 0)
                    strReceiptList = strReceiptList.Substring(0, strReceiptList.Length - 1);



                strSQL = @" select ReturnNO from  OA_CST_Return where PurchaseNO=@PurchaseNO and  status<>'N'";
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;

                cmd.Parameters.Add("@PurchaseNO", SqlDbType.NVarChar).Value = strPurchase;
                dr = cmd.ExecuteReader();
                DataTable dt_Return = new DataTable();
                dt_Return.Load(dr);
                dr.Dispose();
                cmd.Dispose();

                string strReturnList = "";
                for (int i = 0; i < dt_Return.Rows.Count; i++)
                {
                    strReturnList += dt_Return.DefaultView[i]["ReturnNO"].ToString() + ",";
                }
                if (strReturnList.Length > 0)
                    strReturnList = strReturnList.Substring(0, strReturnList.Length - 1);


                //当所有退货单/出货单都取消确认/删除时, 采购单回到可作废状态
                strSQL = @"Update OA_CST_Purchase set ReceiptList=@ReceiptList, ReturnList=@ReturnList  Where PurchaseNO=@PurchaseNO ";
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;
                cmd.Parameters.Add("@ReceiptList", SqlDbType.NVarChar).Value = strReceiptList;
                cmd.Parameters.Add("@ReturnList", SqlDbType.NVarChar).Value = strReturnList;
                cmd.Parameters.Add("@PurchaseNO", SqlDbType.NVarChar).Value = strPurchase;
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                #endregion


                strSQL = @"Select OrderNO from OA_CST_Purchase Where PurchaseNO=@PurchaseNO";
                cmd = new SqlCommand(strSQL, cn);
                if (sTrans != null)
                    cmd.Transaction = sTrans;
                cmd.Parameters.Add("@PurchaseNo", SqlDbType.NVarChar).Value = strPurchase;
                object obj = cmd.ExecuteScalar();
                cmd.Dispose();
                if (obj != null)
                {
                    strSQL = @"Update OA_CST_Order set status='F' from (Select inum=count(*) from OA_CST_Purchase where  status not in('N','F') and OrderNo=@OrderNo ) C  
                            Where C.inum=0 and OA_CST_Order.OrderNo=@OrderNo  and  OA_CST_Order.status not in('N')";
                    cmd = new SqlCommand(strSQL, cn);
                    if (sTrans != null)
                        cmd.Transaction = sTrans;
                    cmd.Parameters.Add("@OrderNo", SqlDbType.NVarChar).Value = Convert.ToString(obj);
                    cmd.ExecuteNonQuery();
                    cmd.Dispose();

                }



              

            }
            catch (Exception ex)
            {
                log.Error(ex.ToString());
                throw ex;
            }

        }//end of method


    }//end of class
}