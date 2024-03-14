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

using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel; //ref NPOI.OOXML + OpenXml4
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;

using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Microsoft.Office.Core;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;

namespace Kcis.Controllers
{
    public class CarApplyController : Controller
    {
        private static ILog log = LogManager.GetLogger(typeof(CarApplyController));

        //外部引用
        public ActionResult ManageJSPartial()
        {
            return PartialView();   
        } 

        //管理专区专用选单
        public ActionResult ManageMenuPartial()
        {
            return PartialView();
        }

        //主申请页
        [Common.ActionFilter.CheckSessionFilter]
        public ActionResult Fill_MainPage()
        {
            Kcis.Models.UserModel user = (Kcis.Models.UserModel)Session["UserProfile"];
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["CarDB"].ConnectionString);
            SqlTransaction sTrans = null;
 
            try
            {

                //取申请人帐号,姓名

                cn.Open();

                string strSQL = @"Select ParentName, email, today=CONVERT(varchar(100), GETDATE(), 111) from Shuttle.dbo.Car_ApplyForm where StudentNO='" + user.UserId + "'  ";
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_Apply = new DataTable();
                dt_Apply.Load(dr);
                dr.Dispose();
                cmd.Dispose();

                string strEmail = "";
                string strParentName = "";
                string strToDay = DateTime.Now.ToString("yyyy/MM/dd");
                if (dt_Apply.Rows.Count > 0)
                {
                    strParentName = dt_Apply.DefaultView[0]["ParentName"].ToString();
                    strEmail = dt_Apply.DefaultView[0]["Email"].ToString();
                    strToDay = dt_Apply.DefaultView[0]["today"].ToString();
                }


                string strNeedCheck = Convert.ToString(System.Web.Configuration.WebConfigurationManager.AppSettings["NeedCheck"]);

                ViewBag.strNeedCheck = strNeedCheck;
                ViewBag.ParentName = strParentName;
                ViewBag.Email = strEmail;
                ViewBag.Today = strToDay;

            }
            catch (Kcis.Models.KcisException e)
            {

                //sTrans.Rollback();

                log.Error(e.ToString());
               
            }
            catch (Exception e)
            {

               //sTrans.Rollback();

                log.Error(e.ToString());
            
            }

            finally
            {
                if (cn!=null  && cn.State != ConnectionState.Closed)
                    cn.Close();

            }


            return View();
        }

        [Common.ActionFilter.CheckSessionFilter]
        public ActionResult GetCarListAjax()
        {

            Kcis.Models.UserModel user = (Kcis.Models.UserModel)Session["UserProfile"];
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["CarDB"].ConnectionString);
            string strMessage = "";


            DataTable dt_Dates = new DataTable();
            dt_Dates.Columns.Add("strStatus", typeof(string));
            dt_Dates.Columns.Add("strMessage", typeof(string));

            dt_Dates.Columns.Add("CurrCount", typeof(string));
            dt_Dates.Columns.Add("MaxLevel", typeof(string));
            dt_Dates.Columns.Add("PicID", typeof(string));


            System.Data.DataRow dRow = dt_Dates.NewRow();
            dRow["CurrCount"] = "";
            dRow["MaxLevel"] = System.Web.Configuration.WebConfigurationManager.AppSettings["CarMaxLevel"];
            dRow["PicID"] = "";

            try
            {
                cn.Open();


                string strSQL = @"Select aa.CarNo, CarColor=replace(isnull(CarColor,''),'色',''), PhoneNumber=isnull(PhoneNumber,'暂无电话'), Phase=isnull(Phase,'已审批'),  PicID=isnull(PicID,''),
                                    enabled= CASE bb.enabled WHEN 'Y' THEN '正常'  ELSE '停用' END
                                    from Car_Base aa inner join  Car_CarPerson bb on aa.CarNo=bb.CarNo Where (bb.enabled = 'Y' and aa.enabled='Y' or Phase='C' or Phase='X')  and bb.Personid = '" + user.UserId + "'";
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_Man = new DataTable();
                dt_Man.Load(dr);
                dr.Dispose();
                cmd.Dispose();
                int iTotal = dt_Man.Rows.Count;
                string strCarList = @"<h3>已登录车牌号</h3>
                        <div class='row' style='margin-top:20px;text-align:left'>
                           ";
                for (int i = 0; i < dt_Man.Rows.Count; i++)
                {
                    if(dt_Man.DefaultView[0]["PicID"].ToString().Length>0)
                       dRow["PicID"] = dt_Man.DefaultView[0]["PicID"];

                    string strStatus = "";
                    if (!dt_Man.DefaultView[i]["enabled"].ToString().Equals("正常"))
                        strStatus = "<span style='color:red'><b>[已停用]</b></span>";

                    if (dt_Man.DefaultView[i]["Phase"].ToString().Equals("C"))  //有姊妹者，当一人已过。另一人申请时将不会自动-已审批 避免有家长盗用别人车牌入校
                        strStatus = "<span style='color:red'><b>[审批中]</b></span>";
                    else if (dt_Man.DefaultView[i]["Phase"].ToString().Equals("X"))  //有姊妹者，当一人已过。另一人申请时将不会自动-已审批 避免有家长盗用别人车牌入校
                        strStatus = "<span style='color:red'><b>[已驳回]</b></span>";

                    strCarList += @"<div class='col-xs-12 col-lg-12'>@Status
                                    <label>车牌号:<span class='text - danger'>@CarNO(@Color色)</span></label>
                                    <label>手机号:<span class='text - danger'>@Phone</span></label>
                                    <button type='button' class='btn btn-danger btn-xs' name='btn_Del' id='@ID' onclick='DelCarNo(this)'>删除(DEL)</button></div>";
                    strCarList = strCarList.Replace("@Status", strStatus);
                    strCarList = strCarList.Replace("@CarNO", dt_Man.DefaultView[i]["CarNo"].ToString());
                    strCarList = strCarList.Replace("@Color色", dt_Man.DefaultView[i]["CarColor"].ToString());
                    strCarList = strCarList.Replace("@ID", dt_Man.DefaultView[i]["CarNo"].ToString());
                    strCarList = strCarList.Replace("@Phone", dt_Man.DefaultView[i]["PhoneNumber"].ToString());
                }

                //strSQL = @"Select CarNo=CarNumber, CarColor=replace(isnull(CarColor,''),'色',''), PhoneNumber=isnull(PhoneNumber,'暂无电话') from   Link247.[WebApp].[dbo].[OA_CarNumberRegister]
                //            Where flag = 'N' and accountid = '" + user.UserId + "'";
                //cmd = new SqlCommand(strSQL, cn);
                //dr = cmd.ExecuteReader();
                //dt_Man = new DataTable();
                //dt_Man.Load(dr);
                //dr.Dispose();
                //cmd.Dispose();
                //iTotal += dt_Man.Rows.Count;
                //for (int i = 0; i < dt_Man.Rows.Count; i++)
                //{
                //    //车牌号:苏E9C93X(深蓝) 手机号:13916323985       
                //    strCarList += @"<div class='col-xs-12 col-lg-12'>
                //                    <label>车牌号:<span class='text - danger'>@CarNO(@Color色)</span></label>
                //                    <label>手机号:<span class='text - danger'>@Phone</span></label>[审核中]
                //                    <button type='button' class='btn btn-danger btn-xs' name='btn_Del' id='@ID' onclick='DelCarNo(this)'>删除(DEL)</button></div>";
                //    strCarList = strCarList.Replace("@CarNO", dt_Man.DefaultView[i]["CarNo"].ToString());
                //    strCarList = strCarList.Replace("@Color色", dt_Man.DefaultView[i]["CarColor"].ToString());
                //    strCarList = strCarList.Replace("@ID", dt_Man.DefaultView[i]["CarNo"].ToString());
                //    strCarList = strCarList.Replace("@Phone", dt_Man.DefaultView[i]["PhoneNumber"].ToString());
                //}

                strCarList += @"</div></div>";

                dRow["CurrCount"] = iTotal;
                if (iTotal == 0)
                    strCarList = "";

            

                dRow["strStatus"] = "[ok]";
                dRow["strMessage"] = strCarList;
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
        public ActionResult DelCarListAjax(IEnumerable<HttpPostedFileBase> files, FormCollection collection)
        {

            Kcis.Models.UserModel user = (Kcis.Models.UserModel)Session["UserProfile"];
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["CarDB"].ConnectionString);
            string strMessage = "";
 
     
            string strCarNo= collection["inp_CarNo"].ToString();
            string strKind = collection["inp_Kind"].ToString();
                try
            {
                cn.Open();
                SqlCommand cmd = null;
                string strSQL = "";

                //只有审核中的enabled='N' 才能被家长删除
                //若关系删光了 连车牌也须连动删除避免家长入校

                //删除时只要发现flag=N的记录就直接删除
                strSQL = @"Select Ctype from Car_Base Where CarNo=N'" + strCarNo + "' and (remark<>N'审批驳回，禁止入校' and remark<>'新车牌等待审核中') ";
                cmd = new SqlCommand(strSQL, cn);
                string strCtype = Convert.ToString(cmd.ExecuteScalar());
                cmd.Dispose();

                if (strCtype.ToUpper().Equals("X"))
                    throw new Kcis.Models.KcisException("此车牌由于管理需求无法进行删除，有问题可与学校交通组联系！");

                //留记录
                strSQL = @"Insert into Car_CarPersonLog([Personid] ,[CarNO],[Enabled] ,[Remark],[CreateTime] ,[CreateUser])
                           SELECT Personid , CarNO, Enabled , Remark, getdate() ,CreateUser FROM [dbo].[Car_CarPerson] where  Personid='" + user.UserId + "' and CarNo=N'" + strCarNo + "'";
                cmd = new SqlCommand(strSQL, cn);
                cmd.ExecuteNonQuery();
                cmd.Dispose();

                //标记删除 并取消审核注记(否则仍会出现在审批页面与家长页)
                //strSQL = @"Update Car_CarPerson set Enabled='N', Phase='N' , PicID='', PicName='' where  Personid='" + user.UserId + "' and CarNo=N'" + strCarNo + "'";
                //cmd = new SqlCommand(strSQL, cn);
                //cmd.Parameters.Add("@Remark", SqlDbType.NVarChar).Value = "家长已自行删除";
                //cmd.ExecuteNonQuery();
                //cmd.Dispose();

                strSQL = @"Delete from Car_Base where  CarNo=N'" + strCarNo + "'";
                cmd = new SqlCommand(strSQL, cn);
                cmd.ExecuteNonQuery();
                cmd.Dispose();

                strSQL = @"Delete from Car_CarPerson where  CarNo=N'" + strCarNo + "'";
                cmd = new SqlCommand(strSQL, cn);
                cmd.ExecuteNonQuery();
                cmd.Dispose();

                //是否还存在 有效或审核中的 绑订关系？
                //strSQL = @"SELECT count(*) FROM [dbo].[Car_CarPerson] where (isnull(enabled,'N')='Y' or isnull(Phase,'')='C') and  Personid='" + user.UserId + "' and CarNo=N'" + strCarNo + "'";
                //cmd = new SqlCommand(strSQL, cn);
                //int iCount = Convert.ToInt32(cmd.ExecuteScalar());
                //cmd.Dispose();

                //if (iCount == 0)  //已无绑任何学生或审批中
                //{
                //    //留记录
                //    strSQL = @"Insert into Car_BaseLog([CarNo] ,[CarColor] ,[Ctype],[SourceType] ,[PhoneNumber] ,[CreateTime] ,[CreateUser],[UpdateTime],[Enabled] ,[Remark],[IP])
                //           SELECT [CarNo] ,[CarColor] ,[Ctype],[SourceType] ,[PhoneNumber] ,[CreateTime] ,[CreateUser],[UpdateTime],[Enabled] ,[Remark],[IP] FROM [dbo].[Car_Base] where CarNo=N'" + strCarNo + "'";
                //    cmd = new SqlCommand(strSQL, cn);
                //    cmd.ExecuteNonQuery();
                //    cmd.Dispose();

                //    //标记删除
                //    strSQL = @"Delete from Car_Base where  CarNo=N'" + strCarNo + "'";
                //    cmd = new SqlCommand(strSQL, cn);
                //    cmd.ExecuteNonQuery();
                //    cmd.Dispose();

                //    strSQL = @"Delete from Car_CarPerson where  CarNo=N'" + strCarNo + "'";
                //    cmd = new SqlCommand(strSQL, cn);
                //    cmd.ExecuteNonQuery();
                //    cmd.Dispose();

                //}



                strMessage = "车牌已删除完成！";

            }
            catch (Kcis.Models.KcisException e)
            {
                log.Error(e.ToString());
                return Content("{Error}" + e.Message);
            }
            catch (Exception e)
            {
                log.Error(e.ToString());
                return Content("[Error]");
            }

            finally
            {
                if (cn.State != ConnectionState.Closed)
                    cn.Close();

            }

            return Content(strMessage);
        }//end of func 

        //增加车牌人员绑定
        [Common.ActionFilter.CheckSessionFilter]
        public ActionResult AddCarListAjax(IEnumerable<HttpPostedFileBase> files, FormCollection collection)
        {

            string strExceptionCanAdd = "[家长已自行删除],[审批驳回，禁止入校]";

            Kcis.Models.UserModel user = (Kcis.Models.UserModel)Session["UserProfile"];
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["CarDB"].ConnectionString);
            //SqlTransaction sTrans = null;
            string strMessage = "";


            string strCarNo = collection["inp_CarNoNew"].ToString();
            string strTel = collection["inp_Tel"].ToString();
            string strColor = collection["inp_Color"].ToString().Replace("色","");
            try
            {
                cn.Open();
                int iMaxLevel = Convert.ToInt32(System.Web.Configuration.WebConfigurationManager.AppSettings["CarMaxLevel"]);
                string strNeedCheck = Convert.ToString(System.Web.Configuration.WebConfigurationManager.AppSettings["NeedCheck"]);

                string strSQL = @"Select * from Shuttle.dbo.OA_StudentsPS where remark<>'转出' and  accountid='" + user.UserId + "'  ";
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_Apply = new DataTable();
                dt_Apply.Load(dr);
                dr.Dispose();
                cmd.Dispose();

                if (dt_Apply.Rows.Count ==0)
                    throw new Kcis.Models.KcisException("您身分非本校学生或已经转出！");

                //判断是否已经填写家长同意书
                strSQL = @"Select * from Shuttle.dbo.Car_ApplyForm where StudentNO='" + user.UserId + "'  ";
                cmd = new SqlCommand(strSQL, cn);
                dr = cmd.ExecuteReader();
                dt_Apply = new DataTable();
                dt_Apply.Load(dr);
                dr.Dispose();
                cmd.Dispose();

                if (dt_Apply.Rows.Count == 0)
                    throw new Kcis.Models.KcisException("您尚未提交线上家长同意书无法申请车牌！");


                //数总车牌数已经等于上线数，不接受再新增 
                //strSQL = @"Select * from Car_Base aa inner join Car_CarPerson bb on aa.CarNo=bb.CarNo
                //           Where  bb.enabled='Y' and bb.Personid='" + user.UserId + "'";
                // 不可以判断nabled='Y' 已避免 经后台拉黑关闭后 继续申请
                strSQL = @"Select * from Car_CarPerson Where enabled='Y' and Personid='" + user.UserId + "'";
                log.Debug("----------------------------strSQL1=" + strSQL);
                cmd = new SqlCommand(strSQL, cn);
                dr = cmd.ExecuteReader();
                dt_Apply = new DataTable();
                dt_Apply.Load(dr);
                dr.Dispose();
                cmd.Dispose();

                int iTotalCar = dt_Apply.Rows.Count;
                if (iTotalCar >= iMaxLevel)
                    throw new Kcis.Models.KcisException("申请失败！您已大于" + iMaxLevel + "个车牌数限制无法再登录新车牌～");


                //检查车牌：告知无法申请的原因避免一直打电话来资讯
                strSQL = @"Select Ctype, enabled, remark from Car_Base Where CarNo='" + strCarNo + "' ";
                cmd = new SqlCommand(strSQL, cn);
                dr = cmd.ExecuteReader();
                DataTable dt_Car_Base = new DataTable();
                dt_Car_Base.Load(dr);
                dr.Dispose();
                cmd.Dispose();
                if (dt_Car_Base.Rows.Count > 0) {
                    
                    if (dt_Car_Base.DefaultView[0]["Ctype"].ToString().Equals("X") && strExceptionCanAdd.IndexOf(dt_Car_Base.DefaultView[0]["remark"].ToString())<0)
                        throw new Kcis.Models.KcisException("申请失败！[" + strCarNo + "]此车牌已被[学校交通组]限制入校～");

                }


                //检查车牌与学号关系
                strSQL = @"Select enabled=isnull(enabled,'N') from Car_CarPerson  Where Personid='" + user.UserId + "' and CarNo='"+strCarNo+"'";
                cmd = new SqlCommand(strSQL, cn);
                dr = cmd.ExecuteReader();
                DataTable dt_CarPerson = new DataTable();
                dt_CarPerson.Load(dr);
                dr.Dispose();
                cmd.Dispose();

                //图片上传
                //Convert.ToString(System.Guid.NewGuid());
                string[] matrix_FileName = { "inp_DBPicID1", "inp_DBPicID2" };
                string[] matrix_FileNameValue = { "", "" };

                string[] matrix_HTMLFile = System.Text.RegularExpressions.Regex.Split(collection["inp_PicAdd"].ToString(), ",", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                HttpFileCollectionBase myfiles = Request.Files;
                int index = 0;
                //因为myfiles若第一个无档案，第二个会自动补位故不能用myfiles当loop顺序
                string strFileName = "";
                for (int i = 0; i < matrix_HTMLFile.Length; i++)
                {
                    if (matrix_HTMLFile[i].Equals(""))
                    {
                        matrix_FileNameValue[i] = Common.Models.Utility.UtilityString.TrimNull(collection[matrix_FileName[i]]);
                        continue;
                    }
                    HttpPostedFileBase upfile = myfiles[index];
                    matrix_FileNameValue[i] = Common.Models.Utility.UtilityString.TrimNull(collection[matrix_FileName[i]]);  //若没上传就是维持原档名
                    if (upfile != null && upfile.ContentLength > 0)
                    {
                        ////给新档名
                        matrix_FileNameValue[i] = Convert.ToString(System.Guid.NewGuid()) + Path.GetExtension(upfile.FileName);
                        string filePath = Path.Combine(HttpContext.Server.MapPath("~/ArtPIC"), matrix_FileNameValue[i]);
                        upfile.SaveAs(filePath);
                        strFileName = upfile.FileName;
                    }
                    index++;
                }


                if (dt_CarPerson.Rows.Count == 0)
                {
                    strSQL = @"Insert into Car_CarPerson([Personid] ,[CarNO],[Enabled] ,[Remark],[CreateTime] ,[CreateUser], Phase, IP, PICID, PicName)
                            SELECT @Personid ,@CarNO, @Enabled , @Remark, getdate() , @CreateUser, @Phase, @IP, @PICID, @PicName    ";
                    cmd = new SqlCommand(strSQL, cn);
                    cmd.Parameters.Add("@Personid", SqlDbType.NVarChar).Value = user.UserId;
                    cmd.Parameters.Add("@CarNO", SqlDbType.NVarChar).Value = strCarNo;
                    if (!strNeedCheck.Equals("Y"))
                    {
                        cmd.Parameters.Add("@Enabled", SqlDbType.NVarChar).Value = "Y";
                        cmd.Parameters.Add("@Phase", SqlDbType.NVarChar).Value = "Y";
                    }
                    else
                    {
                        cmd.Parameters.Add("@Enabled", SqlDbType.NVarChar).Value = "Y";   
                        cmd.Parameters.Add("@Phase", SqlDbType.NVarChar).Value = "C";
                    }
                    cmd.Parameters.Add("@Remark", SqlDbType.NVarChar).Value = "";
                    cmd.Parameters.Add("@PICID", SqlDbType.NVarChar).Value = matrix_FileNameValue[0];
                    cmd.Parameters.Add("@PicName", SqlDbType.NVarChar).Value = strFileName;
                    cmd.Parameters.Add("@CreateUser", SqlDbType.NVarChar).Value = user.UserId;
                    cmd.Parameters.Add("@IP", SqlDbType.NVarChar).Value = HttpContext.Request.UserHostAddress;


                    

                    cmd.ExecuteNonQuery();
                    cmd.Dispose();

                }
                else
                {
                    //if(dt_CarPerson.DefaultView[0]["enabled"].ToString().Equals("Y"))
                    //    throw new Kcis.Models.KcisException("此车牌已存在不需重复申请～");

                    //复原人车关系
                    strSQL = @"Update [dbo].[Car_CarPerson] set enabled=@Enabled, Phase=@Phase, CreateTime=getdate() [PIC] where  Personid='" + user.UserId + "' and CarNo='" + strCarNo + "'";

                    //处理图档
                    if (strFileName.Length > 0)
                    {
                        //, PICID='" + matrix_FileNameValue[0] + "', PicName='" + strFileName + "'
                        strSQL = strSQL.Replace("[PIC]", ", PICID = '" + matrix_FileNameValue[0] + "', PicName = '" + strFileName + "'");
                    }
                    else
                        strSQL = strSQL.Replace("[PIC]", "");

                    cmd = new SqlCommand(strSQL, cn);
                    if (!strNeedCheck.Equals("Y"))
                    {
                        cmd.Parameters.Add("@Enabled", SqlDbType.NVarChar).Value = "Y";
                        cmd.Parameters.Add("@Phase", SqlDbType.NVarChar).Value = "Y";
                    }
                    else
                    {
                        cmd.Parameters.Add("@Enabled", SqlDbType.NVarChar).Value = "Y";  //避免未审核数据出现在前台与后台
                        cmd.Parameters.Add("@Phase", SqlDbType.NVarChar).Value = "C";
                    }

                    cmd.ExecuteNonQuery();
                    cmd.Dispose();

                }//end of 关系


                //开始处理车牌：需考虑车牌是否需审核
                if (dt_Car_Base.Rows.Count == 0)
                {
                    strSQL = @" Select Sourcetype=CASE isnull(Sourcetype,'B') WHEN 'B' THEN 'B' WHEN 'K' THEN 'K' ELSE 'B' END from kcis_account Where Accountid='" + user.UserId + "' ";
                    cmd = new SqlCommand(strSQL, cn);
                    string strSourceType = Convert.ToString(cmd.ExecuteScalar());
                    cmd.Dispose();

                    //处理新车牌
                    strSQL = @"Insert into Car_Base([CarNo] ,[CarColor] ,[Ctype],[SourceType] ,[PhoneNumber] ,[CreateTime] ,[CreateUser],[UpdateTime],[Enabled] ,[Remark],[IP])
                                SELECT @CarNo, @CarColor, @Ctype, @SourceType ,@PhoneNumber , getdate() ,@CreateUser, getdate(), @Enabled , @Remark, @IP ";
                    cmd = new SqlCommand(strSQL, cn);
                    cmd.Parameters.Add("@CarNo", SqlDbType.NVarChar).Value = strCarNo;
                    cmd.Parameters.Add("@CarColor", SqlDbType.NVarChar).Value = strColor;
                    cmd.Parameters.Add("@Ctype", SqlDbType.NVarChar).Value = strSourceType;
                    cmd.Parameters.Add("@SourceType", SqlDbType.NVarChar).Value = "F";
                    cmd.Parameters.Add("@PhoneNumber", SqlDbType.NVarChar).Value = strTel;

                    if (strNeedCheck.Equals("Y"))  //先关闭待审核的新车牌
                    {
                        cmd.Parameters.Add("@Enabled", SqlDbType.NVarChar).Value = 'N';
                        cmd.Parameters.Add("@Remark", SqlDbType.NVarChar).Value = "新车牌等待审核中";
                    }
                    else
                    {
                        cmd.Parameters.Add("@Enabled", SqlDbType.NVarChar).Value = 'Y';
                        cmd.Parameters.Add("@Remark", SqlDbType.NVarChar).Value = "新车牌无需审批立即生效";
                    }
                    
                    cmd.Parameters.Add("@CreateUser", SqlDbType.NVarChar).Value = user.UserId;
                    cmd.Parameters.Add("@IP", SqlDbType.NVarChar).Value = HttpContext.Request.UserHostAddress; 
                    cmd.ExecuteNonQuery();
                    cmd.Dispose();

                }
                else
                {
                    //是否还存在 有效的绑订关系？
                    strSQL = @"SELECT count(*) FROM [dbo].[Car_CarPerson] where isnull(enabled,'N')='Y' and Phase='Y' and  CarNo=N'" + strCarNo + "'";
                    cmd = new SqlCommand(strSQL, cn);
                    int iCount = Convert.ToInt32(cmd.ExecuteScalar());
                    cmd.Dispose();
 
                    //老车牌
                    strSQL = @"Update [dbo].[Car_Base] set enabled=@Enabled, PhoneNumber=@PhoneNumber, CarColor=@CarColor, Remark=@Remark  where CarNo = '" + strCarNo + "'";
                    cmd = new SqlCommand(strSQL, cn);
                    cmd.Parameters.Add("@CarColor", SqlDbType.NVarChar).Value = strColor;
                    cmd.Parameters.Add("@PhoneNumber", SqlDbType.NVarChar).Value = strTel;

                    if (iCount > 0)  //还存在 有效的绑订关系
                    {
                        cmd.Parameters.Add("@Enabled", SqlDbType.NVarChar).Value = 'Y';
                        cmd.Parameters.Add("@Remark", SqlDbType.NVarChar).Value = "";
                    }
                    else if (strNeedCheck.Equals("Y"))  // 无有效开放者，且要求审核
                    {
                        cmd.Parameters.Add("@Enabled", SqlDbType.NVarChar).Value = 'N';
                        cmd.Parameters.Add("@Remark", SqlDbType.NVarChar).Value = "新车牌等待审核中";
                    }
                    else
                    {
                        cmd.Parameters.Add("@Enabled", SqlDbType.NVarChar).Value = 'Y';
                        cmd.Parameters.Add("@Remark", SqlDbType.NVarChar).Value = "新车牌无需审批立即生效";
                    }
                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                }//end of carno


                if (strNeedCheck.Equals("Y"))
                    strMessage = "车牌提交申请完成，请等待审核作业大约需要1个工作天！";
                else
                    strMessage = "车牌已经申请完成与审核成功，欢迎入校！";


            }
            catch (Kcis.Models.KcisException e)
            {
               // sTrans.Rollback();
                log.Error(e.ToString());
                return Content("{Error}" + e.Message);
            }
            catch (Exception e)
            {
               // sTrans.Rollback();
                log.Error(e.ToString());
                return Content("[Error]");
            }

            finally
            {
                if (cn.State != ConnectionState.Closed)
                    cn.Close();

            }

            return Content(strMessage);
        }//end of func 


        //send form
        [Common.ActionFilter.CheckSessionFilter]
        public ActionResult Fill_SendFormAjax(IEnumerable<HttpPostedFileBase> files, FormCollection collection)
        {
            Kcis.Models.UserModel user = (Kcis.Models.UserModel)Session["UserProfile"];
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["CarDB"].ConnectionString);
            SqlTransaction sTrans = null;
           
            try
            {
           
                //取申请人帐号,姓名

                cn.Open();
                sTrans = cn.BeginTransaction();    //get insert script from visual management

                string strSQL = @"Select * from Shuttle.dbo.OA_StudentsPS where AccountID='" + user.UserId + "'";
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                int iCount = Convert.ToInt32(cmd.ExecuteScalar());
                cmd.Dispose();
                if (iCount == 0)
                    throw new Kcis.Models.KcisException("学生家长才能申请入校车接！");

                strSQL = @"Select * from Shuttle.dbo.Car_ApplyForm where StudentNO='" + user.UserId + "'  ";
                cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt_Apply = new DataTable();
                dt_Apply.Load(dr);
                dr.Dispose();
                cmd.Dispose();

                if (dt_Apply.Rows.Count > 0)
                {
                    strSQL = @"Delete from Shuttle.dbo.Car_ApplyForm where StudentNO='" + user.UserId + "'  ";
                    cmd = new SqlCommand(strSQL, cn);
                    cmd.Transaction = sTrans;
                    cmd.ExecuteNonQuery();

                }
                     
 
                string strParentName = collection["inp_ParentName"].ToString();
                string strEmail = collection["inp_Email"].ToString();


                if (strEmail.IndexOf("@")<0)
                    throw new Kcis.Models.KcisException("邮箱格式错误！");

                strSQL = @"INSERT INTO Shuttle.dbo.Car_ApplyForm
                                ([StudentNo]
                                ,[ParentName]
                                ,[Email], IP
                                ,[CreateTime])
                            VALUES
                                (@StudentNo 
                                ,@ParentName
                                ,@Email, @IP
                                ,getdate() )";

                cmd = new SqlCommand(strSQL, cn);
                cmd.Transaction = sTrans;
                cmd.Parameters.Add("@StudentNo", SqlDbType.NVarChar).Value = user.UserId;                         
                cmd.Parameters.Add("@ParentName", SqlDbType.NVarChar).Value = strParentName;
                cmd.Parameters.Add("@Email", SqlDbType.NVarChar).Value = strEmail;
                cmd.Parameters.Add("@IP", SqlDbType.NVarChar).Value = HttpContext.Request.UserHostAddress;

                cmd.ExecuteNonQuery();
                cmd.Dispose();
               
                sTrans.Commit();
                sTrans.Dispose();

                return Content("回执单已提交成功，请继续登录车牌！(Submit form successfully, thank you!)");
            }
            catch (Kcis.Models.KcisException e)
            {
        
                sTrans.Rollback();

                log.Error(e.ToString());
                return Content("{Error}" + e.Message);
            }
            catch (Exception e)
            {
           
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



        public ActionResult ApplyMessage()
        {

            return View();
        }


    }//end of class
}
