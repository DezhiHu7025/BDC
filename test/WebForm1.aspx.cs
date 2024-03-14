using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace test
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            //  Word2Html(@"C:\Users\wenjie_liang\Desktop\1007.docx", "http://portal.kcistz.org.cn/BDC/Manager/", "1007.docx");

            byte[] fileData = System.IO.File.ReadAllBytes(Server.MapPath("1007.docx"));
            Response.Clear();
            Response.ContentType = "application/msword";
            Response.AddHeader("Content-Disposition", "attachment; filename=ZooAnimals.docx");
            Response.OutputStream.Write(fileData, 0, fileData.Length);
            Response.End();
        }



    }
}