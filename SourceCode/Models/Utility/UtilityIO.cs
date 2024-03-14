using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using log4net;
using System.Data.SqlClient;
using System.IO;
using System.Net.Mail;
using System.Data;

namespace Kcis.Models.Utility
{
    public class UtilityIO
    {

        public static string RequestPage(string strUrlLink)
        {
            string strContent = "";
            try
            {
                System.Net.HttpWebRequest fr;
                Uri targetUri = new Uri(strUrlLink);

                fr = (System.Net.HttpWebRequest)System.Net.HttpWebRequest.Create(targetUri);


                System.IO.StreamReader str = new System.IO.StreamReader(fr.GetResponse().GetResponseStream());

                strContent = str.ReadToEnd();

                if (str != null) str.Close();


                System.Console.WriteLine("WebService Finish...");
                return strContent;
            }
            catch (Exception e)
            {
                return "Error:" + e.ToString();
                throw e;
            }
        }


        public static bool SendEmail(string toAddr, string userName, string subject, string content)
        {
            //if (!mes.Utility.Config.IsEmail)
            //    return true;

            MailMessage em = new MailMessage();
            //em.From = new System.Net.Mail.MailAddress(mes.Utility.Config.client_FromAddr, mes.Utility.Config.client_FromName, System.Text.Encoding.UTF8);
            //em.To.Add(new MailAddress(toAddr, userName));

            //em.SubjectEncoding = System.Text.Encoding.UTF8;
            //em.BodyEncoding = System.Text.Encoding.UTF8;
            //em.Subject = subject;
            //em.Body = content;
            //em.IsBodyHtml = true;

            //System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();
            //client.Credentials = new System.Net.NetworkCredential(mes.Utility.Config.client_Credentials_Account, mes.Utility.Config.client_Credentials_Password);
            //client.Port = mes.Utility.Config.client_Port;
            //client.Host = mes.Utility.Config.client_Host;
            //client.Send(em);

            return true;

        }// end of func




    }//end of class
}