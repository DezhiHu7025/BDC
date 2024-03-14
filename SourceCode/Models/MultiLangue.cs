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

    public sealed class MultiLangue
    {
        private static ILog log = LogManager.GetLogger(typeof(MultiLangue));

        private static MultiLangue _instance = null;
        // Creates an syn object.
        private static readonly object SynObject = new object();
        private int iCount = 1;
        private Hashtable ht_en = new Hashtable();
        private Hashtable ht_cn = new Hashtable();
        MultiLangue()
        {
            //只有初始时执行一次
            log.Debug("---MultiLangue()初始化执行开始～, iCount="+iCount);
 
            SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["unis"].ConnectionString);
            try
            {
                cn.Open();
                string strSQL = @"Select FKey, en_key, cn_key from webapp.dbo.OA_xxx_Lang ";
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                SqlDataReader dr = cmd.ExecuteReader();
                                
                while (dr.Read())
                {

                    ht_en[dr["FKey"].ToString()] =dr["en_key"].ToString();
                    ht_cn[dr["FKey"].ToString()] = dr["cn_key"].ToString();
                    //log.Debug(dr["FKey"].ToString());
                }
                dr.Dispose();
                cmd.Dispose();
                log.Debug("---MultiLangue()初始化执行完成！");
            }
            catch (Exception e)
            {
                log.Error(e.ToString());
                 
            }
            finally
            {
                if (cn.State != ConnectionState.Closed)
                    cn.Close();
            }


        }//end of construtor

        //从数据库要中英翻译
        public string WriteText(string lang, string fkey)
        {
            //log.Debug("lang=" + lang);
            if (ht_en.ContainsKey(fkey) && ht_cn.ContainsKey(fkey))
                if ((lang.ToLower().Equals("zh-cn") || lang.ToLower().Equals("zh-tw")))
                    return ht_cn[fkey].ToString();
                else
                    return ht_en[fkey].ToString();

            else
                return fkey;
        }

        //如果是中文环境，就直接输出kjey, 否则输出evalue 与数据库无关  是用于单一参考的情况
        public string WriteText(string lang, string fkey, string strEValue)
        {
            //log.Debug("lang=" + lang);
  
            if ((lang.ToLower().Equals("zh-cn") || lang.ToLower().Equals("zh-tw")))
                return fkey;
            else
                return strEValue;

        }

        public static MultiLangue Instance
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
                            _instance = new MultiLangue();

                        }//end of if
                    }
                }
                return _instance;
            }
        }//end of sttaic method
    }//end of class
}