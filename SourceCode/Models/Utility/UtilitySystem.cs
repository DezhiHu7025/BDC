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
    public class UtilitySystem
    {

        public static string TransNavigation(string strVersion ,string strSource)
        {

            if (strVersion.Equals("eng"))
                strSource = "KCIS System";

            return strSource;
        }// end of method

        public static bool CheckUserExist(Kcis.Models.UserModel user)
        {
            bool flag = true;

            if (user == null || user.UserId == null || user.UserId.Equals(""))
                flag = false;


            return flag;
        }// end of method
    }//end of class

}