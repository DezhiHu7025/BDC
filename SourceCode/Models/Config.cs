using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Kcis.Models
{
    public class Config
    {
        public static string client_Host = "";
        public static int client_Port = 25;
        public static string client_Credentials_Account = "";
        public static string client_Credentials_Password = "";
        public static string client_FromAddr = "";
        public static string client_FromName = "";
        public static string FilePath = "";
        public static string WebURL = "";
        public static string FMWebURL = "";
        public static string FMWeb_Port = "";
        public static string VirtualFolder = "doc";
        public static string HeadTitle = "";
        public static string GroupTitle = "";
        public static bool IsSharing = false;
        public static bool IsEmail = false;
        public static string SEGMENT_NO = "";
        public static string SEGMENT_NO_SZ = "";
        public static bool IsNextTerm = true;
        public static int iAutoCancelTime = 0;
    }
}