using System.Web;
using System.Web.Optimization;

namespace Kcis
{
    public class BundleConfig
    {
        // 有关 Bundling 的详细信息，请访问 http://go.microsoft.com/fwlink/?LinkId=254725
        public static void RegisterBundles(BundleCollection bundles)
        {
            //作用于_ManagerLayout.cshtml 的引用设置
            //bundles.Add(new ScriptBundle("~/bundles/jquery").Include("~/Scripts/app_js/jquery-1.4.4.js"));
             

            //bundles.Add(new StyleBundle("~/Content/app_css").Include("~/Content/app_css/j*"));  //客制化需求手工维护,若可自动化则优先采用

        }//end of method
    }
}