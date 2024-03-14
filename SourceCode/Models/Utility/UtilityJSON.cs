using System;
using System.Collections.Generic;
using System.Web;
using System.Data.SqlClient;
using System.IO;
using System.Net.Mail;
using System.Data;
using System.Collections;

namespace Kcis.Models.Utility
{
    public class UtilityJSON
    {

 

        public static string RetJsonMessage(string strKey, string strMessage)
        {
            string strJson = "";
            DataTable dt_Dates2 = new DataTable();
            dt_Dates2.Columns.Add("strStatus", typeof(string));
            dt_Dates2.Columns.Add("strMessage", typeof(string));


            System.Data.DataRow dRow = dt_Dates2.NewRow();
            dRow["strStatus"] = strKey;
            dRow["strMessage"] = strMessage;


            dt_Dates2.Rows.Add(dRow);

            strJson = Newtonsoft.Json.JsonConvert.SerializeObject(dt_Dates2);

            return strJson;

        }// end of method
    }//end of class

}