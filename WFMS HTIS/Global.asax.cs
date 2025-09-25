using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Web;
using System.Web.Helpers;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;
using System.Globalization;
namespace Payroll.portal
{
    public class MvcApplication : System.Web.HttpApplication
    {
        //string siteUrl = System.Configuration.ConfigurationManager.AppSettings["SiteUrl"];
        //double directUrlTimeOut = Convert.ToDouble(System.Configuration.ConfigurationManager.AppSettings["DirectUrlTimeOut"]);
        //string qs = commonFunctions.encryptStringStatic(DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);
            //  UnityConfig.RegisterComponents();

            //  AntiForgeryConfig.UniqueClaimTypeIdentifier = ClaimTypes.Name;
        }
        //protected void Application_BeginRequest(object sender, EventArgs e)
        //{
        //    //if (Request.Url.Authority.StartsWith("www"))
        //    //    return; 

        //    //var url = string.Format("{0}://www.{1}{2}",
        //    //            Request.Url.Scheme,
        //    //            Request.Url.Authority,
        //    //            Request.Url.PathAndQuery);

        //    string r = Request.HttpMethod.ToString();
        //    if (r == "GET")
        //    {
        //        CultureInfo provider = CultureInfo.InvariantCulture;
        //        string format = "dd-MMM-yyyy HH:mm:ss";
        //        DateTime sessionDate;
        //        string url = HttpContext.Current.Request.Url.ToString();
        //        commonFunctions comfun = new commonFunctions();
        //        string cDate = qs;
        //        if (url.Contains("qs"))
        //        {
        //            cDate = comfun.decryptString(Request.QueryString["qs"].ToString().Trim()).Trim();
        //            if (cDate.Contains("#"))
        //            {
        //                cDate = cDate.Replace("#", "");
        //            }
        //            if (DateTime.TryParseExact(cDate, format, provider, DateTimeStyles.None, out sessionDate))
        //            {
        //                TimeSpan span = (DateTime.Now - sessionDate);
        //                if (span.Minutes > directUrlTimeOut)
        //                {
        //                    string qstime = comfun.encryptString(DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
        //                    url = siteUrl + "/account/SessionExpired?qs=" + qstime;
        //                    Response.Redirect(url.ToString(), true);
        //                }
        //            }
        //            else
        //            {
        //                string qstime = comfun.encryptString(DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                         
        //                url = siteUrl + "/account/SessionExpired?qs=" + qstime;
        //                Response.Redirect(url.ToString(), true);
        //            }
        //        }

        //        if (!url.ToString().Contains("qs"))
        //        {
        //            string qstime = comfun.encryptString(DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
        //            if (url.Contains("?"))
        //            {
        //                url = url.Trim() + "&qs=" + qstime;
        //            }
        //            else
        //            {
        //                url = url.Trim() + "?qs=" + qstime;
        //            }
        //            Response.Redirect(url.ToString(), true);
        //        }
        //    }

        //}
        //protected void Application_BeginRequest(Object source, EventArgs e)
        //{
        //     var url = HttpContext.Current.Request.Url.AbsoluteUri;
        //    // validate thi surl using Regex and re found faulty then redirect... it

        //    //{
        //    //    HttpApplication app = (HttpApplication)source;
        //    //    HttpContext context = app.Context;

        //    //    string host = FirstRequestInitialisation.Initialise(context);
        //    //}
        //}
        class FirstRequestInitialisation
        {
            private static string host = null;

            private static Object s_lock = new Object();

            // Initialise only on the first request
            public static string Initialise(HttpContext context)
            {
                if (string.IsNullOrEmpty(host))
                {
                    lock (s_lock)
                    {
                        if (string.IsNullOrEmpty(host))
                        {
                            Uri uri = HttpContext.Current.Request.Url;
                            host = uri.Scheme + Uri.SchemeDelimiter + uri.Host + ":" + uri.Port;
                        }
                    }
                }

                return host;
            }
        }
    }
}
