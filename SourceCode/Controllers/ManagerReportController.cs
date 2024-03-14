using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using log4net;
using Microsoft.Office.Core;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel; //ref NPOI.OOXML + OpenXml4

namespace SourceCode.Controllers
{
    public class ManagerReportController : Controller
    {
        //
        private static ILog log = LogManager.GetLogger(typeof(ManagerReportController));


        /**--------------------------------------------------------------------------------------------------------**/
        //Prj月报导出-主页
        [Common.ActionFilter.CheckSessionFilter]
        public ActionResult SalesMonthReport_MainPage()
        {
            Kcis.Models.UserModel user = (Kcis.Models.UserModel)Session["UserProfile"];
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);

            try
            {


                cn.Open();
                string strSQL = @"select CONVERT(varchar(100), GETDATE(), 111) ";
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                string strEndDay = Convert.ToString(cmd.ExecuteScalar());
                cmd.Dispose();

                strSQL = @"select CONVERT(varchar(100), dateadd(month,-1, GETDATE()), 111) ";
                cmd = new SqlCommand(strSQL, cn);
                string strBeginDay = Convert.ToString(cmd.ExecuteScalar());
                cmd.Dispose();


                ViewBag.BeginDay = strBeginDay;
                ViewBag.EndDay = strEndDay;


                log.Debug("user.GroupIds=" + Common.Models.Utility.UtilityString.Dot2SQL(user.GroupIds));
         
                strSQL = @"Select distinct aa.PrjID  from OA_ScanCard_ItemGroup aa inner join OA_ScanCard_ItemPrice bb on aa.prjid=bb.prjid where  aa.enabled='Y'  and bb.enabled='Y'  ";
                if (user.GroupIds.IndexOf("manager") < 0)
                    strSQL += " and aa.groupid in(" + Common.Models.Utility.UtilityString.Dot2SQL(user.GroupIds) + ") ";
                
                log.Debug("strSQL=" + strSQL);
                cmd = new SqlCommand(strSQL, cn);
                //cmd.Parameters.Add("@groupid", SqlDbType.NVarChar).Value = ;
                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_Menu = new DataTable();
                dt_Menu.Load(dr);
                dr.Dispose();
                cmd.Dispose();


                string strMenu = "<select name=\"Sel_Menu1\">";
                if (dt_Menu.Rows.Count == 0)
                {
                    strMenu += "<option value=''>你的群组 " + Common.Models.Utility.UtilityString.Dot2SQL(user.GroupIds) + " 尚未配置消费项目</option>";
                }
                else
                    for (int i = 0; i < dt_Menu.Rows.Count; i++)
                    {
                        strMenu += "<option value='" + dt_Menu.DefaultView[i]["prjid"] + "'>" + dt_Menu.DefaultView[i]["prjid"] + "</option>";
                    }

                strMenu += "</select>";

                log.Debug("menu=" + strMenu);
                ViewBag.strMenu = strMenu;
                

            }
            catch (Exception e)
            {
                log.Error(e.ToString());
                return Content(e.Message);
            }

            finally
            {
                if (cn.State != ConnectionState.Closed)
                    cn.Close();
            }

            return View();
        }

        public ActionResult SalesMonthReport_LoadReport(string strBeginDay, string strEndDay, string strMenu1)
        {




            // unis.Models.UserModel user = (unis.Models.UserModel)Session["UserProfile"];
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            try
            {


                cn.Open();

                string strSdate = "", strEdate = "";
                //选择月费的最后一天
                //strEdate = strMonth + "/31";

                //找选择月份的起始日期
                //string strSql = @"SELECT CONVERT(varchar(7),dateadd(month,-0, CONVERT(DATETIME,'"+strEdate+"')     ),111)+'/01'  ";;
                //SqlCommand cmd = new SqlCommand(strSql, cn);
                //strSdate = Convert.ToString( cmd.ExecuteScalar());
                //strSdate = strMonth + "/01";



                //CASE remark WHEN 'SCH' THEN '中信银行' ELSE remark END


                //年級	中文班級	英文班級	學號	姓名	缴款日期	缴款金额	汇入银行
                string strSql = @" Select convert(int, grade_level) as [年级], [班級]=Home_Room, [學號]=aa.accountID, [姓名]=Sname,
                                 [消费时间]=CONVERT(varchar(100), BillTime, 120), [消费项目]=ProjectID, [消费子项]=ItemID, [消费金额]=Price, [交易帐号]=CreateAccount  
                                  from webapp.dbo.OA_Payment_Bill aa inner join Webapp.dbo.OA_StudentProfile bb on aa.accountID=bb.accountID 
                                where Status='Y' and ProjectID='" + strMenu1 + "' and CONVERT(varchar(100), BillTime, 111) BETWEEN '" + strBeginDay + "' AND '" + strEndDay + "' order by BillTime desc ";

                log.Debug("strSql=" + strSql);
                SqlCommand cmd = new SqlCommand(strSql, cn);

                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_MyForm = new DataTable();
                dt_MyForm.Load(dr);
                dr.Dispose();
                cmd.Dispose();

                //组装NPOI开始
                string fileName = "report.xlsx";
                IWorkbook workbook = null;
                if (fileName.IndexOf(".xlsx") > 0)
                {
                    workbook = new XSSFWorkbook();
                }
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                    workbook = new HSSFWorkbook();



                ISheet sheet = null;
                if (workbook != null)
                {
                    sheet = workbook.CreateSheet("Sheet1");
                }


                //设置列宽
                int[] columnWidth = { 6, 8, 8, 12, 25, 30, 30, 10, 20 };

                for (int k = 0; k < columnWidth.Length; k++)
                {
                    //设置列宽度，256*字符数，因为单位是1/256个字符
                    sheet.SetColumnWidth(k, 256 * columnWidth[k]);
                }


                ICellStyle style1 = workbook.CreateCellStyle();
                style1.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                style1.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                style1.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                style1.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;

                IFont font = workbook.CreateFont();
                //font.FontHeightInPoints = 16;
                font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                //font.FontName = "標楷體";
                style1.SetFont(font);
                style1.WrapText = true;

                //第二组style
                ICellStyle style2 = workbook.CreateCellStyle();
                style2.FillBackgroundColor = IndexedColors.Red.Index;

                int i = 0;
                int j = 0;
                int count = 0;
                if (true)
                { //写入DataTable的列名  
                    IRow row = sheet.CreateRow(0);
                    for (j = 0; j < dt_MyForm.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(dt_MyForm.Columns[j].ColumnName);
                        row.GetCell(j).CellStyle = style1;
                    }
                    count = 1;
                }


                for (i = 0; i < dt_MyForm.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < dt_MyForm.Columns.Count; ++j)
                    {

                        if (dt_MyForm.Columns[j].ColumnName.Equals("消费金额"))
                        {  //score
                            row.CreateCell(j).SetCellType(CellType.Numeric);
                            string strTemp = Convert.ToDouble(dt_MyForm.Rows[i][j]).ToString("0.00");
                            //log.Debug("strTemp=" + strTemp);
                            row.CreateCell(j).SetCellValue(Convert.ToDouble(strTemp));
                        }
                        else
                        {
                            row.CreateCell(j).SetCellType(CellType.String);
                            row.CreateCell(j).SetCellValue(dt_MyForm.Rows[i][j].ToString());
                        }


                        row.GetCell(j).CellStyle = style2;
                    }
                    ++count;
                }

                log.Debug("---------------S4");
                using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
                {
                    workbook.Write(ms);
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", Server.UrlEncode("SalesMonthReport_report" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx")));
                    Response.Clear();
                    Response.BinaryWrite(ms.ToArray());
                    Response.End();
                }
                log.Debug("---------------S5");





            }
            catch (Exception e)
            {
                log.Error(e.ToString());
                return null;
            }

            finally
            {
                if (cn.State != ConnectionState.Closed)
                    cn.Close();
            }
            return Content("文档导出错误！");
        }//end of func

        /**--------------------------------------------------------------------------------------------------------**/


        [Common.ActionFilter.CheckSessionFilter]
        public ActionResult Finance01Report_Main()
        {
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            try
            {
                cn.Open();
                //string strSql = "";
                //SqlCommand cmd = null;
                //string strOption = "";
                //for (int i = 0; i < 12; i++)
                //{
                //    strSql = @"SELECT   CONVERT(varchar(7),dateadd(month,-"+i+",getdate()),111)  "; ;
                //    cmd = new SqlCommand(strSql, cn);
                //    string str = Convert.ToString(cmd.ExecuteScalar());
                //    cmd.Dispose();
                //    strOption += "<option value='" + str + "'>" + str + "</option>";

                //}
                //ViewBag.SelectItem = strOption;

                string strSQL = @"select CONVERT(varchar(100), GETDATE(), 111) ";
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                string strEndDay = Convert.ToString(cmd.ExecuteScalar());
                cmd.Dispose();

                strSQL = @"select CONVERT(varchar(100), dateadd(month,-1, GETDATE()), 111) ";
                cmd = new SqlCommand(strSQL, cn);
                string strBeginDay = Convert.ToString(cmd.ExecuteScalar());
                cmd.Dispose();


                ViewBag.BeginDay = strBeginDay;
                ViewBag.EndDay = strEndDay;


            }
            catch (Exception e)
            {
                log.Error(e.ToString());
                return null;
            }

            finally
            {
                if (cn.State != ConnectionState.Closed)
                    cn.Close();
            }

            return View();
        }


        public ActionResult LoadFinance01Report(string strBeginDay, string strEndDay)
        {

            

            
           // unis.Models.UserModel user = (unis.Models.UserModel)Session["UserProfile"];
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            try
            {


                cn.Open();

                string strSdate="", strEdate="";
                //选择月费的最后一天
                //strEdate = strMonth + "/31";

                //找选择月份的起始日期
                //string strSql = @"SELECT CONVERT(varchar(7),dateadd(month,-0, CONVERT(DATETIME,'"+strEdate+"')     ),111)+'/01'  ";;
                //SqlCommand cmd = new SqlCommand(strSql, cn);
                //strSdate = Convert.ToString( cmd.ExecuteScalar());
                //strSdate = strMonth + "/01";

    

                //CASE remark WHEN 'SCH' THEN '中信银行' ELSE remark END


                //年級	中文班級	英文班級	學號	姓名	缴款日期	缴款金额	汇入银行
                string strSql = @"  Select convert(int, grade_level) as [年级], 中文班級='', [英文班級]=Home_Room, [學號]=aa.accountID, [姓名]=Sname,
                                    [缴款日期]=fdate, [缴款金额]=Price, SourceName as [汇入银行]   from webapp.dbo.OA_Payment_Fill aa inner join Webapp.dbo.OA_StudentProfile bb on aa.accountID=bb.accountID 
                                    where convert(int, grade_level)<=9
                                  and CONVERT(varchar(100), CreateTime, 111) BETWEEN '" + strBeginDay + "' AND '" + strEndDay + "' ";
 
                strSql += "  Order by CreateTime ";


                log.Debug("strSql=" + strSql);
                SqlCommand cmd = new SqlCommand(strSql, cn);

                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_MyForm = new DataTable();
                dt_MyForm.Load(dr);
                dr.Dispose();
                cmd.Dispose();

                //组装NPOI开始
                string fileName = "report.xlsx";
                IWorkbook workbook = null;
                if (fileName.IndexOf(".xlsx") > 0)
                {
                    workbook = new XSSFWorkbook();
                }
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                    workbook = new HSSFWorkbook();



                ISheet sheet = null;
                if (workbook != null)
                {
                    sheet = workbook.CreateSheet("Sheet1");
                }


                //设置列宽
                int[] columnWidth = { 10, 15, 15, 15, 12, 12, 15, 20};

                for (int k = 0; k < columnWidth.Length; k++)
                {
                    //设置列宽度，256*字符数，因为单位是1/256个字符
                    sheet.SetColumnWidth(k, 256 * columnWidth[k]);
                }


                ICellStyle style1 = workbook.CreateCellStyle();
                style1.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                style1.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                style1.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                style1.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;

                IFont font = workbook.CreateFont();
                //font.FontHeightInPoints = 16;
                font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                //font.FontName = "標楷體";
                style1.SetFont(font);
                style1.WrapText = true;

                //第二组style
                ICellStyle style2 = workbook.CreateCellStyle();
                style2.FillBackgroundColor = IndexedColors.Red.Index;

                int i = 0;
                int j = 0;
                int count = 0;
                if (true)
                { //写入DataTable的列名  
                    IRow row = sheet.CreateRow(0);
                    for (j = 0; j < dt_MyForm.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(dt_MyForm.Columns[j].ColumnName);
                        row.GetCell(j).CellStyle = style1;
                    }
                    count = 1;
                }


                for (i = 0; i < dt_MyForm.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < dt_MyForm.Columns.Count; ++j)
                    {

                        if (j == 6)
                        {  //score
                            row.CreateCell(j).SetCellType(CellType.Numeric);
                            string strTemp = Convert.ToDouble(dt_MyForm.Rows[i][j]).ToString("0.00");
                            //log.Debug("strTemp=" + strTemp);
                            row.CreateCell(j).SetCellValue(Convert.ToDouble(strTemp));
                        }
                        else { 
                            row.CreateCell(j).SetCellType(CellType.String);
                            row.CreateCell(j).SetCellValue(dt_MyForm.Rows[i][j].ToString());
                        }
                            

                        row.GetCell(j).CellStyle = style2;
                    }
                    ++count;
                }

                log.Debug("---------------S4");
                using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
                {
                    workbook.Write(ms);
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", Server.UrlEncode("Finance01_report" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx")));
                    Response.Clear();
                    Response.BinaryWrite(ms.ToArray());
                    Response.End();
                }
                log.Debug("---------------S5");





            }
            catch (Exception e)
            {
                log.Error(e.ToString());
                return null;
            }

            finally
            {
                if (cn.State != ConnectionState.Closed)
                    cn.Close();
            }
            return Content("文档导出错误！");
        }//end of func

        [Common.ActionFilter.CheckSessionFilter]
        public ActionResult Finance02Report_Main()
        {
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            try
            {
                cn.Open();
                string strSQL = @"select CONVERT(varchar(100), GETDATE(), 111) ";
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                string strEndDay = Convert.ToString(cmd.ExecuteScalar());
                cmd.Dispose();

                strSQL = @"select CONVERT(varchar(100), dateadd(month,-1, GETDATE()), 111) ";
                cmd = new SqlCommand(strSQL, cn);
                string strBeginDay = Convert.ToString(cmd.ExecuteScalar());
                cmd.Dispose();


                ViewBag.BeginDay = strBeginDay;
                ViewBag.EndDay = strEndDay;

            }
            catch (Exception e)
            {
                log.Error(e.ToString());
                return null;
            }

            finally
            {
                if (cn.State != ConnectionState.Closed)
                    cn.Close();
            }

            return View();
        }



        public ActionResult LoadFinance02Report(string strBeginDay, string strEndDay)
        {

            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            try
            {


                cn.Open();
                //CASE remark WHEN 'SCH' THEN '中信银行' ELSE remark END
                //年級	中文班級	英文班級	學號	姓名	缴款日期	缴款金额	汇入银行
                string strSql = @"  Select convert(int, grade_level) as [年级], 中文班級='', [英文班級]=Home_Room, [學號]=aa.accountID, [姓名]=Sname,
                                    [缴款日期]=fdate, [缴款金额]=Price, SourceName as [汇入银行]   from webapp.dbo.OA_Payment_Fill aa inner join Webapp.dbo.OA_StudentProfile bb on aa.accountID=bb.accountID 
                                    where convert(int, grade_level)>9
                                  and CONVERT(varchar(100), CreateTime, 111) BETWEEN '" + strBeginDay + "' AND '" + strEndDay + "' "; 

                strSql += "  Order by CreateTime ";


                log.Debug(">>>strSql=" + strSql);
                SqlCommand cmd = new SqlCommand(strSql, cn);

                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_MyForm = new DataTable();
                dt_MyForm.Load(dr);
                dr.Dispose();
                cmd.Dispose();

                //组装NPOI开始
                string fileName = "report.xlsx";
                IWorkbook workbook = null;
                if (fileName.IndexOf(".xlsx") > 0)
                {
                    workbook = new XSSFWorkbook();
                }
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                    workbook = new HSSFWorkbook();



                ISheet sheet = null;
                if (workbook != null)
                {
                    sheet = workbook.CreateSheet("Sheet1");
                }


                //设置列宽
                int[] columnWidth = { 10, 15, 15, 15, 12, 12, 15, 20 };

                for (int k = 0; k < columnWidth.Length; k++)
                {
                    //设置列宽度，256*字符数，因为单位是1/256个字符
                    sheet.SetColumnWidth(k, 256 * columnWidth[k]);
                }


                ICellStyle style1 = workbook.CreateCellStyle();
                style1.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                style1.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                style1.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                style1.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;

                IFont font = workbook.CreateFont();
                //font.FontHeightInPoints = 16;
                font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                //font.FontName = "標楷體";
                style1.SetFont(font);
                style1.WrapText = true;

                //第二组style
                ICellStyle style2 = workbook.CreateCellStyle();
                style2.FillBackgroundColor = IndexedColors.Red.Index;

                int i = 0;
                int j = 0;
                int count = 0;
                if (true)
                { //写入DataTable的列名  
                    IRow row = sheet.CreateRow(0);
                    for (j = 0; j < dt_MyForm.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(dt_MyForm.Columns[j].ColumnName);
                        row.GetCell(j).CellStyle = style1;
                    }
                    count = 1;
                }


                for (i = 0; i < dt_MyForm.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < dt_MyForm.Columns.Count; ++j)
                    {

                        if (j == 6)
                        {  //score
                            row.CreateCell(j).SetCellType(CellType.Numeric);
                            string strTemp = Convert.ToDouble(dt_MyForm.Rows[i][j]).ToString("0.00");
                            log.Debug("strTemp=" + strTemp);
                            row.CreateCell(j).SetCellValue(Convert.ToDouble(strTemp));
                        }
                        else
                        {
                            row.CreateCell(j).SetCellType(CellType.String);
                            row.CreateCell(j).SetCellValue(dt_MyForm.Rows[i][j].ToString());
                        }


                        row.GetCell(j).CellStyle = style2;
                    }
                    ++count;
                }

                log.Debug("---------------S4");
                using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
                {
                    workbook.Write(ms);
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", Server.UrlEncode("Finance02_report" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx")));
                    Response.Clear();
                    Response.BinaryWrite(ms.ToArray());
                    Response.End();
                }
                log.Debug("---------------S5");





            }
            catch (Exception e)
            {
                log.Error(e.ToString());
                return null;
            }

            finally
            {
                if (cn.State != ConnectionState.Closed)
                    cn.Close();
            }
            return Content("文档导出错误！");
        }//end of func



        [Common.ActionFilter.CheckSessionFilter]
        public ActionResult Finance03Report_Main()
        {
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            try
            {
                cn.Open();
                string strOption = "";

                string strSql = @"Select YYYYMM from OA_Payment_CalculateList group by YYYYMM order by YYYYMM desc"; ;
                SqlCommand cmd = new SqlCommand(strSql, cn);
                //cmd.Parameters.Add("@YYYYMM", SqlDbType.NVarChar).Value = strLastMonth;
                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_Remain = new DataTable();
                dt_Remain.Load(dr);
                dr.Dispose();
                cmd.Dispose();

                for (int i = 0; i < dt_Remain.Rows.Count; i++)
                {

                    strOption += "<option value='" + dt_Remain.DefaultView[i]["YYYYMM"] + "'>" + dt_Remain.DefaultView[i]["YYYYMM"] + "</option>";

                }

                //log.Debug("strOption=" + strOption);
                ViewBag.SelectItem = strOption;


            }
            catch (Exception e)
            {
                log.Error(e.ToString());
                return null;
            }

            finally
            {
                if (cn.State != ConnectionState.Closed)
                    cn.Close();
            }

            return View();
        }



        public ActionResult LoadFinance03Report(string strMonth)
        {

            // unis.Models.UserModel user = (unis.Models.UserModel)Session["UserProfile"];
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            try
            {

                cn.Open();

                string strSdate = "", strEdate = "";
                //选择月费的最后一天
                strEdate = strMonth + "/31";

                //找选择月份的起始日期
                //string strSql = @"SELECT CONVERT(varchar(7),dateadd(month,-0, CONVERT(DATETIME,'" + strEdate + "')     ),111)+'/01'  ";
                //log.Debug("---sql01=" +strSql);
                //SqlCommand cmd = new SqlCommand(strSql, cn);
                //strSdate = Convert.ToString(cmd.ExecuteScalar());
                strSdate = strMonth + "/01";

 

                //CASE remark WHEN 'SCH' THEN '中信银行' ELSE remark END


                //年級	中文班級	英文班級	學號	姓名	缴款日期	缴款金额	汇入银行
                string strSql = @" Select [年级]=convert(int, grade_level), 中文班級='schoolid='+schoolid +', exitdate='+exitdate , [英文班級]=Home_Room, [學號]=aa.accountID, [姓名]=Sname, [利润中心]=FinanceCatalogName,
                                     [上期结余数]=PreviousRemainAmount, [本期充值金额]=InAmount, [财务本期充值]=isnull(FinanceFillAmount,0),
[输出中心]=[国际部输出中心]+[双语部输出中心],
[寝具加购]=[国际部寝具加购]+[双语部寝具加购],
[制服加购]=[国际部制服加购]+[双语部制服加购],
[学生证（补办）]=[双语部学生证（补办）]+[国际部学生证（补办）],
[泳衣泳帽加购] =[双语部泳衣泳帽加购]+[国际部泳衣泳帽加购],
[打印或复印],
[双语部G3-G6补购中文书本费],
[双语部阅读存折遗失补发],
[双语部人接卡（补办）],
[国际部点心],
[国际部快递费],
[国际部乐器耗材],
[国际部乐团],
[国际部校服学号标烫印],

                                     [本期消费金额]=OutAmount, [本期结余数]=RemainAmount
                                        from webapp.dbo.OA_Payment_CalculateList aa inner join Webapp.dbo.OA_StudentProfile bb on aa.accountID=bb.accountID 
                                      where convert(int, grade_level)<=9 and (OutAmount<>0 or InAmount<>0 or PreviousRemainAmount<>0) and  YYYYMM='" + strMonth + "' Order by grade_level";
 
                log.Debug("strSql=" + strSql);
                SqlCommand cmd = new SqlCommand(strSql, cn);

                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_MyForm = new DataTable();
                dt_MyForm.Load(dr);
                dr.Dispose();
                cmd.Dispose();


                //取empty datatable
                DataTable dtclone = dt_MyForm.Clone();


                //组装NPOI开始
                string fileName = "report.xlsx";
                IWorkbook workbook = null;
                if (fileName.IndexOf(".xlsx") > 0)
                {
                    workbook = new XSSFWorkbook();
                }
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                    workbook = new HSSFWorkbook();



                ISheet sheet = null;
                if (workbook != null)
                {
                    sheet = workbook.CreateSheet("Sheet1");
                }


                //设置列宽 19
                int[] columnWidth = { 10, 15, 15, 15, 16, 12, 16, 16, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 21, 20, 20, 20, 15, 15, 15 };

                for (int k = 0; k < columnWidth.Length; k++)
                {
                    //设置列宽度，256*字符数，因为单位是1/256个字符
                    sheet.SetColumnWidth(k, 256 * columnWidth[k]);
                }


                ICellStyle style1 = workbook.CreateCellStyle();
                style1.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                style1.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                style1.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                style1.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;

                IFont font = workbook.CreateFont();
                //font.FontHeightInPoints = 16;
                font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                //font.FontName = "標楷體";
                style1.SetFont(font);
                style1.WrapText = true;

                //第二组style
                ICellStyle style2 = workbook.CreateCellStyle();
                style2.FillBackgroundColor = IndexedColors.Red.Index;

                int i = 0;
                int j = 0;
                int count = 0;
                if (true)
                { //写入DataTable的列名  
                    IRow row = sheet.CreateRow(0);
                    for (j = 0; j < dt_MyForm.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(dt_MyForm.Columns[j].ColumnName);
                        row.GetCell(j).CellStyle = style1;
                    }
                    count = 1;
                }

 
                System.Data.DataRow dRow = dtclone.NewRow();
                dRow["年级"] = 1;
                dRow["中文班級"] = "中文班級";
                dRow["英文班級"] = "英文班級";
                dRow["學號"] = "學號";
                dRow["姓名"] = "姓名";
                dRow["利润中心"] = "小计";


                for (i = 0; i < dt_MyForm.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < dt_MyForm.Columns.Count; ++j)
                    {

                        if (j >= 6)
                        {  //score
                            row.CreateCell(j).SetCellType(CellType.Numeric);
                            string strTemp = Convert.ToDouble(dt_MyForm.Rows[i][j]).ToString("0.00");
                            log.Debug("strTemp=" + strTemp);
                            row.CreateCell(j).SetCellValue(Convert.ToDouble(strTemp));
                        }
                        else
                        {
                            row.CreateCell(j).SetCellType(CellType.String);
                            row.CreateCell(j).SetCellValue(dt_MyForm.Rows[i][j].ToString());
                        }


                        row.GetCell(j).CellStyle = style2;

                        //累加各栏位消费总额, 6刚好是[双语部G3-G6补购中文书本费]的index
                        if (j >= 6)
                        {
       
                            double dbSum = Convert.ToDouble(dt_MyForm.Rows[i][j]);
                            if (dRow[j] == System.DBNull.Value)
                                dRow[j] = dbSum;
                            else
                                dRow[j] = Convert.ToDouble(dRow[j]) + dbSum;
                        }


                    }//End of J loop
                    ++count;  //完成1笔

                }//end of loop of dt_MyForm

                dtclone.Rows.Add(dRow);

                //显示加总
                IRow rowTotal = sheet.CreateRow(dt_MyForm.Rows.Count + 1);
                int index = 5;
                for (j = index; j < dtclone.Columns.Count; ++j)
                {
                    if (dtclone.Rows[0][j] == System.DBNull.Value)
                        continue;

                    if (j == 5)
                        rowTotal.CreateCell(j).SetCellValue(Convert.ToString(dtclone.Rows[0][j]));
                    else
                    {
                        rowTotal.CreateCell(j).SetCellType(CellType.Numeric);
                        rowTotal.CreateCell(j).SetCellValue(Convert.ToDouble(dtclone.Rows[0][j]));
                    }
                    rowTotal.GetCell(j).CellStyle = style2;
                }

              
                using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
                {
                    workbook.Write(ms);
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", Server.UrlEncode("Finance03_report" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx")));
                    Response.Clear();
                    Response.BinaryWrite(ms.ToArray());
                    Response.End();
                }
        





            }
            catch (Exception e)
            {
                log.Error(e.ToString());
                return null;
            }

            finally
            {
                if (cn.State != ConnectionState.Closed)
                    cn.Close();
            }
            return Content("文档导出错误！");
        }//end of func


        [Common.ActionFilter.CheckSessionFilter]
        public ActionResult Finance04Report_Main()
        {
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            try
            {

                cn.Open();
                string strOption = "";
          
                string strSql = @"Select YYYYMM from OA_Payment_CalculateList group by YYYYMM order by YYYYMM desc"; ;
                SqlCommand cmd = new SqlCommand(strSql, cn);
                //cmd.Parameters.Add("@YYYYMM", SqlDbType.NVarChar).Value = strLastMonth;
                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_Remain = new DataTable();
                dt_Remain.Load(dr);
                dr.Dispose();
                cmd.Dispose();

                for (int i = 0; i < dt_Remain.Rows.Count; i++)
                {

                    strOption += "<option value='" + dt_Remain.DefaultView[i]["YYYYMM"] + "'>" + dt_Remain.DefaultView[i]["YYYYMM"] + "</option>";

                }

                //log.Debug("strOption=" + strOption);
                ViewBag.SelectItem = strOption;

            }
            catch (Exception e)
            {
                log.Error(e.ToString());
                return null;
            }

            finally
            {
                if (cn.State != ConnectionState.Closed)
                    cn.Close();
            }

            return View();
        }



        public ActionResult LoadFinance04Report(string strMonth)
        {
 
            // unis.Models.UserModel user = (unis.Models.UserModel)Session["UserProfile"];
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            try
            {


                cn.Open();

                string strSdate = "", strEdate = "";
                //选择月费的最后一天
                strEdate = strMonth + "/31";

                //找选择月份的起始日期
                //string strSql = @"SELECT CONVERT(varchar(7),dateadd(month,-0, CONVERT(DATETIME,'" + strEdate + "')     ),111)+'/01'  ";
                //log.Debug("---sql01=" +strSql);
                //SqlCommand cmd = new SqlCommand(strSql, cn);
                //strSdate = Convert.ToString(cmd.ExecuteScalar());
                strSdate = strMonth + "/01";

                 
 

                //年級	中文班級	英文班級	學號	姓名	缴款日期	缴款金额	汇入银行
                string strSql = @" Select [年级]=convert(int, grade_level), 中文班級='schoolid='+schoolid +', exitdate='+exitdate   , [英文班級]=Home_Room, [學號]=aa.accountID, [姓名]=Sname, [利润中心]=FinanceCatalogName,
                                     [上期结余数]=PreviousRemainAmount, [本期充值金额]=InAmount, [财务本期充值]=isnull(FinanceFillAmount,0),
[输出中心]=[国际部输出中心]+[双语部输出中心],
[寝具加购]=[国际部寝具加购]+[双语部寝具加购],
[制服加购]=[国际部制服加购]+[双语部制服加购],
[学生证（补办）]=[双语部学生证（补办）]+[国际部学生证（补办）],
[泳衣泳帽加购] =[双语部泳衣泳帽加购]+[国际部泳衣泳帽加购],
[打印或复印],
[国际部点心],
[国际部快递费],
[国际部乐器耗材],
[国际部乐团],
[国际部校服学号标烫印],
                                     [本期消费金额]=OutAmount, [本期结余数]=RemainAmount
                                        from webapp.dbo.OA_Payment_CalculateList aa inner join Webapp.dbo.OA_StudentProfile bb on aa.accountID=bb.accountID 
                                      where convert(int, grade_level)>9 and (OutAmount<>0 or InAmount<>0 or PreviousRemainAmount<>0) and  YYYYMM='" + strMonth + "'";

                strSql += "  Order by grade_level ";


                log.Debug("strSql=" + strSql);
                SqlCommand cmd = new SqlCommand(strSql, cn);

                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_MyForm = new DataTable();
                dt_MyForm.Load(dr);
                dr.Dispose();
                cmd.Dispose();


                //取empty datatable
                //strSql = @"  Select * from webapp.dbo.OA_Payment_CalculateList where 1=0 ";
                //cmd = new SqlCommand(strSql, cn);
                //dr = cmd.ExecuteReader();
                //DataTable dt_Total = new DataTable();
                //dt_Total.Load(dr);
                //dr.Dispose();
                //cmd.Dispose();
                DataTable dtclone = dt_MyForm.Clone();


                //组装NPOI开始
                string fileName = "report.xlsx";
                IWorkbook workbook = null;
                if (fileName.IndexOf(".xlsx") > 0)
                {
                    workbook = new XSSFWorkbook();
                }
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                    workbook = new HSSFWorkbook();



                ISheet sheet = null;
                if (workbook != null)
                {
                    sheet = workbook.CreateSheet("Sheet1");
                }


                //设置列宽
                int[] columnWidth = { 10, 15, 15, 15, 16, 12, 16, 16, 20, 20, 20, 20, 20, 20, 20, 21, 20, 20, 20, 15, 15, 15 };


                for (int k = 0; k < columnWidth.Length; k++)
                {
                    //设置列宽度，256*字符数，因为单位是1/256个字符
                    sheet.SetColumnWidth(k, 256 * columnWidth[k]);
                }


                ICellStyle style1 = workbook.CreateCellStyle();
                style1.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                style1.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                style1.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                style1.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;

                IFont font = workbook.CreateFont();
                //font.FontHeightInPoints = 16;
                font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                //font.FontName = "標楷體";
                style1.SetFont(font);
                style1.WrapText = true;

                //第二组style
                ICellStyle style2 = workbook.CreateCellStyle();
                style2.FillBackgroundColor = IndexedColors.Red.Index;

                int i = 0;
                int j = 0;
                int count = 0;
                if (true)
                { //写入DataTable的列名  
                    IRow row = sheet.CreateRow(0);
                    for (j = 0; j < dt_MyForm.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(dt_MyForm.Columns[j].ColumnName);
                        row.GetCell(j).CellStyle = style1;
                    }
                    count = 1;
                }

                System.Data.DataRow dRow = dtclone.NewRow();
                dRow["年级"] = 1;
                dRow["中文班級"] = "中文班級";
                dRow["英文班級"] = "英文班級";
                dRow["學號"] = "學號";
                dRow["姓名"] = "姓名";
                dRow["利润中心"] = "小计";

                for (i = 0; i < dt_MyForm.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < dt_MyForm.Columns.Count; ++j)
                    {

                        if (j >= 6)
                        {  
                            row.CreateCell(j).SetCellType(CellType.Numeric);
                            string strTemp = Convert.ToDouble(dt_MyForm.Rows[i][j]).ToString("0.00");
                            row.CreateCell(j).SetCellValue(Convert.ToDouble(strTemp));
                        }
                        else
                        {
                            row.CreateCell(j).SetCellType(CellType.String);
                            row.CreateCell(j).SetCellValue(dt_MyForm.Rows[i][j].ToString());
                        }

                        row.GetCell(j).CellStyle = style2;


                        //累加各栏位消费总额, 把来源数据存放于dt_Total第9开始的栏位, 9刚好是[双语部G3-G6补购中文书本费]的index
                        if (j >=6 )
                        {
                            //log.Debug("dt_MyForm.Columns[j].ColumnName=" + dt_MyForm.Columns[j].ColumnName);
                            double dbSum = Convert.ToDouble(dt_MyForm.Rows[i][j]);
                            if (dRow[j] == System.DBNull.Value)
                                dRow[j] = dbSum;
                            else
                                dRow[j] =  Convert.ToDouble(dRow[j] ) + dbSum;
                             
                                
                       
                        }


                    }//End of J loop
                    ++count;  //完成1笔

                }//end of loop of dt_MyForm

                dtclone.Rows.Add(dRow);

                //显示加总
                IRow rowTotal = sheet.CreateRow(dt_MyForm.Rows.Count+1);
                int index =5;
                for (j = index; j < dtclone.Columns.Count; ++j)
                {
                    if (dtclone.Rows[0][j] == System.DBNull.Value)
                        continue;

                    if (j == 5)
                        rowTotal.CreateCell(j).SetCellValue(Convert.ToString(dtclone.Rows[0][j]));
                    else
                    {
                        rowTotal.CreateCell(j).SetCellType(CellType.Numeric);
                        rowTotal.CreateCell(j).SetCellValue(Convert.ToDouble(dtclone.Rows[0][j]));
                    }
                    rowTotal.GetCell(j).CellStyle = style2;
                }
                   
 
 
                log.Debug("---------------S4");
                using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
                {
                    workbook.Write(ms);
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", Server.UrlEncode("Finance04_report" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx")));
                    Response.Clear();
                    Response.BinaryWrite(ms.ToArray());
                    Response.End();
                }
                log.Debug("---------------S5");





            }
            catch (Exception e)
            {
                log.Error(e.ToString());
                return null;
            }

            finally
            {
                if (cn.State != ConnectionState.Closed)
                    cn.Close();
            }
            return Content("文档导出错误！");
        }//end of func
    }
}
