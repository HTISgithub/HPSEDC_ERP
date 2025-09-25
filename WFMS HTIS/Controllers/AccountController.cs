using Payroll.portal.Models;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using Microsoft.Web.Administration;
using Site = Microsoft.Web.Administration.Site;
using System.Data.SqlClient;
using System.Configuration;
using Newtonsoft.Json;


namespace Payroll.portal.Controllers
{
    // [Authorize]
    public class AccountController : Controller
    {
        payrollFunctions payfun = new payrollFunctions();
        commonFunctions comfun = new commonFunctions();
        string Validate = System.Configuration.ConfigurationManager.AppSettings["Validate"];
        string qs = commonFunctions.encryptStringStatic(DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));

        string mainUrl = System.Configuration.ConfigurationManager.AppSettings["SiteUrl"].ToString();
        string baseUrl = System.Configuration.ConfigurationManager.AppSettings["BaseUrl"].ToString();
        SqlConnection conLogin = new SqlConnection(ConfigurationManager.ConnectionStrings["cnLogin"].ConnectionString);//new SqlConnection(@"data source=15.207.246.29,1533; initial catalog=dbsHTIS_wfmslogin; user id=sa; password=data@5_htis#9002; connection timeout=0;pooling=true;Min Pool Size=0;Max Pool Size=100;Pooling=true; ");

        //[AllowAnonymous]
        ////[OutputCache(NoStore = true, Duration = 0, VaryByParam = "None")]
        //public ActionResult Login()
        //{
        //    DataTable dt = comfun.fillDataTable("", "select fvSessionName as [label], fiSessionId as [value] from sysVIKSATSession", null);
        //    List<DropdownModal> studentList = new List<DropdownModal>();
        //    studentList = (from DataRow dr in dt.Rows
        //                   select new DropdownModal()
        //                   {
        //                       Label = dr["label"].ToString(),
        //                       Value = dr["value"].ToString(),
        //                   }).OrderByDescending(d => d.Value).ToList();
        //    ViewBag.SessionList = new SelectList(studentList, "Value", "Label");
        //    if (User.Identity.IsAuthenticated)
        //    {
        //        FormsAuthentication.SignOut();
        //        return RedirectToAction("Login", new { qs = qs });
        //        //return RedirectToAction("redirecting", "Account", new { qs = qs });
        //    }
        //    ViewData["error"] = "";
        //    return View();
        //    //comfun.SendEmail("Error-  Login()", "kuldeep.singh@horizontelecom.in", "redirecting from User.Identity.Name == ", "Login()");

        //    //if (commonFunctions.isLocalHost())
        //    //{
        //    //    var mView = "Login";
        //    //    return View(mView);
        //    //}

        //    //return Redirect(mainUrl + "/Account/Login");
        //}

        [AllowAnonymous]
        public ActionResult NewLogin()
        {
            DataTable dt = comfun.fillDataTable("", "select fvSessionName as [label], fiSessionId as [value] from sysVIKSATSession", null);
            List<DropdownModal> studentList = new List<DropdownModal>();
            studentList = (from DataRow dr in dt.Rows
                           select new DropdownModal()
                           {
                               Label = dr["label"].ToString(),
                               Value = dr["value"].ToString(),
                           }).OrderByDescending(d => d.Value).ToList();
            ViewBag.SessionList = new SelectList(studentList, "Value", "Label");
            if (User.Identity.IsAuthenticated)
            {
                FormsAuthentication.SignOut();
                return RedirectToAction("NewLogin", new { qs = qs });
                //return RedirectToAction("redirecting", "Account", new { qs = qs });
            }
            ViewData["error"] = "";
            return View();
            //comfun.SendEmail("Error-  Login()", "kuldeep.singh@horizontelecom.in", "redirecting from User.Identity.Name == ", "Login()");

            //if (commonFunctions.isLocalHost())
            //{
            //    var mView = "Login";
            //    return View(mView);
            //}

            //return Redirect(mainUrl + "/Account/Login");
        }

        public ActionResult HpsedcLogin()
        {

            return View();
        }

        //[AllowAnonymous]
        //[HttpPost]
        ////[ValidateAntiForgeryToken]
        //public ActionResult Login(LoginViewModel l, string ReturnUrl = "")
        //{
        //    //string subject = "WFMS Login done";
        //    //string domainId = "";

        //    string IsSuccess = "N";
        //    ViewData["error"] = "";

        //    // get session
        //    DataTable dtsess = comfun.fillDataTable("", "select fvSessionName as [label], fiSessionId as [value] from sysVIKSATSession", null);
        //    List<DropdownModal> studentList = new List<DropdownModal>();
        //    studentList = (from DataRow dr in dtsess.Rows
        //                   select new DropdownModal()
        //                   {
        //                       Label = dr["label"].ToString(),
        //                       Value = dr["value"].ToString(),
        //                   }).OrderByDescending(d => d.Value).ToList();
        //    ViewBag.SessionList = new SelectList(studentList, "Value", "Label");
        //    //
        //    if (ModelState.IsValid)
        //    {
        //        //var isValidUser = Membership.ValidateUser(l.UserName, l.Password);
        //        //if (isValidUser)
        //        {
        //            HttpRuntime.Cache.Remove("dtMenu");
        //            List<string> keys = new List<string>();
        //            IDictionaryEnumerator enumerator = HttpRuntime.Cache.GetEnumerator();
        //            while (enumerator.MoveNext())
        //            {
        //                keys.Add(enumerator.Key.ToString());
        //            }
        //            for (int i = 0; i < keys.Count; i++)
        //            {
        //                HttpRuntime.Cache.Remove(keys[i]);
        //            }
        //            if (l.UserName != "sa")
        //            {
        //                SortedList list = new SortedList();
        //                list.Add("@Email", l.UserName);
        //                list.Add("@PassWord", l.Password);
        //                System.Data.DataTable dt = comfun.fillDataTable("wfms_LoginDetail", "", list, conLogin);

        //                if (dt.Rows.Count > 0)
        //                {
        //                    if (dt.Rows[0]["Error"].ToString() == "")
        //                    {

        //                        HttpRuntime.Cache.Remove("dtMenu");
        //                        Session["DomainId"] = dt.Rows[0]["Domainid"].ToString();
        //                        Session["Domain"] = dt.Rows[0]["fvWebDomain"].ToString();
        //                        Session["EmpId"] = dt.Rows[0]["Id"].ToString();
        //                        Session["Email"] = l.UserName;
        //                        Session["Customer"] = dt.Rows[0]["CustomerId"].ToString();
        //                        Session["AmcId"] = 21;
        //                        Session["RoleId"] = dt.Rows[0]["RollId"].ToString();
        //                        Session["UserName"] = l.UserName;
        //                        Session["EmpName"] = dt.Rows[0]["EmpName"].ToString();
        //                        IsSuccess = "Y";

        //                        if (!mainUrl.Contains("//"))
        //                        {
        //                            mainUrl = "https://" + mainUrl;
        //                        }
        //                        payfun.setValuesInCookiesUsingSession();
        //                    }
        //                    else
        //                    {
        //                        ViewData["error"] = dt.Rows[0]["Error"].ToString();
        //                        ViewData["error"] = "Invalid Employee Code or Password";
        //                        ModelState.Remove("Password");
        //                        return View();
        //                    }

        //                    FormsAuthentication.SetAuthCookie(l.UserName, false);
        //                    if (IsSuccess == "Y")
        //                    {
        //                        return Redirect("/TimeSheet/Dashboard");
        //                    }

        //                }
        //                else
        //                {
        //                    ViewData["error"] = "Invalid id  or Password";
        //                    ModelState.Remove("Password");
        //                    return View();
        //                }
        //            }
        //            else
        //            {
        //                Session["EmpId"] = "1";
        //                Session["EmpName"] = "sa";
        //                Session["EmpCode"] = "sa";
        //            }

        //            FormsAuthentication.SetAuthCookie(l.UserName, false);




        //            //return Redirect(mainUrl);
        //            //  subject = "WFMS Login done-"+ mainUrl;
        //            //subject = comfun.SendEmail("WFMS Login done", "kuldeep.singh@horizontelecom.in", subject, subject);
        //            //comfun.decryptString(Session["DomainId"].ToString());
        //            //if (commonFunctions.isLocalHost())
        //            //{
        //            //    return Redirect(baseUrl + "/account/logind?id=" + comfun.encryptString(l.UserName) + "&did=" + comfun.encryptString(Session["DomainId"].ToString()) + "&session=" + l.SessionList);

        //            //}
        //            //return Redirect(mainUrl + "/account/logind?id=" + comfun.encryptString(l.UserName) + "&did=" + comfun.encryptString(Session["DomainId"].ToString()) + "&session=" + l.SessionList);
        //            return View();
        //        }
        //    }


        //    ModelState.Remove("Password");
        //    return View();
        //}

        [AllowAnonymous]
        [HttpPost]
        public ActionResult NewLogin(LoginViewModel l, string ReturnUrl = "")
        {
            //string subject = "WFMS Login done";
            //string domainId = "";

            string IsSuccess = "N";
            ViewData["error"] = "";

            // get session
            DataTable dtsess = comfun.fillDataTable("", "select fvSessionName as [label], fiSessionId as [value] from sysVIKSATSession", null);
            List<DropdownModal> studentList = new List<DropdownModal>();
            studentList = (from DataRow dr in dtsess.Rows
                           select new DropdownModal()
                           {
                               Label = dr["label"].ToString(),
                               Value = dr["value"].ToString(),
                           }).OrderByDescending(d => d.Value).ToList();
            ViewBag.SessionList = new SelectList(studentList, "Value", "Label");
            //
            if (ModelState.IsValid)
            {
                //var isValidUser = Membership.ValidateUser(l.UserName, l.Password);
                //if (isValidUser)
                {
                    HttpRuntime.Cache.Remove("dtMenu");
                    List<string> keys = new List<string>();
                    IDictionaryEnumerator enumerator = HttpRuntime.Cache.GetEnumerator();
                    while (enumerator.MoveNext())
                    {
                        keys.Add(enumerator.Key.ToString());
                    }
                    for (int i = 0; i < keys.Count; i++)
                    {
                        HttpRuntime.Cache.Remove(keys[i]);
                    }
                    if (l.UserName != "sa")
                    {
                        SortedList list = new SortedList();
                        list.Add("@Email", l.UserName);
                        list.Add("@PassWord", l.Password);
                        System.Data.DataTable dt = comfun.fillDataTable("wfms_LoginDetail", "", list, conLogin);

                        if (dt.Rows.Count > 0)
                        {
                            if (dt.Rows[0]["Error"].ToString() == "")
                            {

                                HttpRuntime.Cache.Remove("dtMenu");
                                Session["DomainId"] = dt.Rows[0]["Domainid"].ToString();
                                Session["Domain"] = dt.Rows[0]["fvWebDomain"].ToString();
                                Session["EmpId"] = dt.Rows[0]["Id"].ToString();
                                Session["Email"] = l.UserName;
                                Session["Customer"] = dt.Rows[0]["CustomerId"].ToString();
                                Session["AmcId"] = 21;
                                Session["RoleId"] = dt.Rows[0]["RollId"].ToString();
                                Session["UserName"] = l.UserName;
                                Session["EmpName"] = dt.Rows[0]["EmpName"].ToString();
                                IsSuccess = "Y";

                                if (!mainUrl.Contains("//"))
                                {
                                    mainUrl = "https://" + mainUrl;
                                }
                                payfun.setValuesInCookiesUsingSession();
                            }
                            else
                            {
                                ViewData["error"] = dt.Rows[0]["Error"].ToString();
                                ViewData["error"] = "Invalid Employee Code or Password";
                                ModelState.Remove("Password");
                                return View();
                            }

                            FormsAuthentication.SetAuthCookie(l.UserName, false);
                            if (IsSuccess == "Y")
                            {
                                return Redirect("/TimeSheet/Dashboard");
                            }

                        }
                        else
                        {
                            ViewData["error"] = "Invalid id  or Password";
                            ModelState.Remove("Password");
                            return View();
                        }
                    }
                    else
                    {
                        Session["EmpId"] = "1";
                        Session["EmpName"] = "sa";
                        Session["EmpCode"] = "sa";
                    }

                    FormsAuthentication.SetAuthCookie(l.UserName, false);




                    //return Redirect(mainUrl);
                    //  subject = "WFMS Login done-"+ mainUrl;
                    //subject = comfun.SendEmail("WFMS Login done", "kuldeep.singh@horizontelecom.in", subject, subject);
                    //comfun.decryptString(Session["DomainId"].ToString());
                    //if (commonFunctions.isLocalHost())
                    //{
                    //    return Redirect(baseUrl + "/account/logind?id=" + comfun.encryptString(l.UserName) + "&did=" + comfun.encryptString(Session["DomainId"].ToString()) + "&session=" + l.SessionList);

                    //}
                    //return Redirect(mainUrl + "/account/logind?id=" + comfun.encryptString(l.UserName) + "&did=" + comfun.encryptString(Session["DomainId"].ToString()) + "&session=" + l.SessionList);
                    return View();
                }
            }


            ModelState.Remove("Password");
            return View();
        }

        [AllowAnonymous]
        public ActionResult Logind()
        {

            string userName = comfun.decryptString(Request.QueryString["id"].ToString());
            string domainId = comfun.decryptString(Request.QueryString["did"].ToString());

            HttpRuntime.Cache.Remove("dtMenu");
            List<string> keys = new List<string>();
            IDictionaryEnumerator enumerator = HttpRuntime.Cache.GetEnumerator();
            while (enumerator.MoveNext())
            {
                keys.Add(enumerator.Key.ToString());
            }
            for (int i = 0; i < keys.Count; i++)
            {
                HttpRuntime.Cache.Remove(keys[i]);
            }


            SortedList list = new SortedList();
            list.Add("@Email", userName);
            System.Data.DataTable dt = comfun.fillDataTable("APPEmp_LoginWithEmailDOMAIN", "", list);
            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["Error"].ToString() == "")
                {
                    //comfun.SendEmail("Error-   Logind", "kuldeep.singh@horizontelecom.in", " Logind", " Logind");
                    HttpRuntime.Cache.Remove("dtMenu");

                    //Session["DomainId"] = dt.Rows[0]["DomainId"].ToString();
                    //Session["EmpId"] = dt.Rows[0]["Id"].ToString();
                    //Session["EmpName"] = dt.Rows[0]["EmpName"].ToString();
                    //Session["EmpCode"] = dt.Rows[0]["EmpCode"].ToString();
                    //Session["Email"] = dt.Rows[0]["Email"].ToString();
                    //Session["loginID"] = dt.Rows[0]["MainloginID"].ToString();
                    //Session["AmcId"] = 21;

                    //  Session["CompanyName"] = dt.Rows[0]["CompanyName"].ToString();
                    // Session["CompanyId"] = dt.Rows[0]["CompanyId"].ToString();

                    Session["Customer"] = dt.Rows[0]["CustomerID"].ToString();
                    if (Session["SessionId"] == null)
                    {
                        SortedList lists = new SortedList();
                        lists.Add("@sessionId", Convert.ToInt32(Request.QueryString["session"].ToString()));
                        System.Data.DataTable dtSession = comfun.fillDataTable("stpgetSessionById", "", lists);
                        Session["SessionId"] = dtSession.Rows[0]["SessionId"].ToString();
                        Session["SessionName"] = dtSession.Rows[0]["SessionName"].ToString();
                    }


                    // Session["CostCenterName"] = dt.Rows[0]["CostCenterName"].ToString();
                    // Session["CostCenterId"] = dt.Rows[0]["CostCenterId"].ToString();

                    // Session["LocationId"] = dt.Rows[0]["LocationId"].ToString();
                    // Session["LocationName"] = dt.Rows[0]["LocationName"].ToString();

                    Session["EmpCategoryId"] = dt.Rows[0]["EmpCategoryId"].ToString();

                    //Session["RoleName"] = Enum.Parse(typeof(payrollFunctions.roleIds), "Employee").ToString();
                    Session["RoleName"] = dt.Rows[0]["RoleName"].ToString();
                    //Session["RoleId"] = ((int)Enum.Parse(typeof(payrollFunctions.roleIds), "Employee")).ToString();
                    Session["RoleId"] = dt.Rows[0]["RoleId"].ToString();

                    payfun.setValuesInCookiesUsingSession();
                }
                else
                {
                    ViewData["error"] = dt.Rows[0]["Error"].ToString();
                }
            }
            else
            {
                ViewData["error"] = "Invalid Employee Code or Password";
                mainUrl = System.Configuration.ConfigurationManager.AppSettings["SiteUrl"].ToString();
                return Redirect(mainUrl + "/account/login");
            }

            try
            {
                FormsAuthentication.SetAuthCookie(userName, false);
            }
            catch (Exception ex)
            {
                string subject = "WFMS Logind";
                // subject = comfun.SendEmail("WFMS Logind", "kuldeep.singh@horizontelecom.in", subject, ex.Message);
            }
            if (Session["RoleId"].ToString() == "2")
            {
                return Redirect("/TimeSheet/Dashboard");
            }
            else
            {
                return Redirect(mainUrl + "/account/login");

            }
            //return RedirectToAction("setup", "Account", new { qs = Request.QueryString["id"].ToString(),did= Request.QueryString["did"].ToString() });
        }

        public bool SetAppSettings(string domain)
        {
            try
            {
                string key = "";


                System.Configuration.Configuration appConfig = null;// System.Configuration.ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

                if (System.Web.HttpContext.Current != null)
                {
                    appConfig =
                        System.Web.Configuration.WebConfigurationManager.OpenWebConfiguration("~");
                }
                else
                {
                    System.Configuration.ExeConfigurationFileMap map = new ExeConfigurationFileMap { ExeConfigFilename = $"{System.AppDomain.CurrentDomain.BaseDirectory}Web.Config" };
                    appConfig = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None);
                }
                AppSettingsSection appSettings = (AppSettingsSection)appConfig.GetSection("appSettings");
                key = "dbsHTIS_wfms" + commonFunctions.getSubDomain(domain);
                if (appSettings.Settings[key] == null)
                {
                    /*************
                    ****Create DB****
                    *************/

                    SortedList list = new SortedList();
                    list.Add("@Db_Name", key);
                    comfun.executeNonQuery("stp_Database", "", list, conLogin);


                    string cn = @"data source=15.207.246.29,1533; initial catalog=" + key + "; user id=sa; password=data@5_htis#9002; connection timeout=0;pooling=true;Min Pool Size=0;Max Pool Size=100;Pooling=true;";

                    appSettings.Settings.Add(key, cn);
                }

                appConfig.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection("appSettings");
                return true;
            }
            catch (Exception ex)
            {
                string subject = "WFMS error";
                //subject = comfun.SendEmail("WFMS error", "kuldeep.singh@horizontelecom.in", subject, ex.Message);
                return false;
            }
        }

        //[Route("/redirecting")]
        [AllowAnonymous]

        public ActionResult setup()
        {
            //Check if setup already done
            string domainId = comfun.decryptString(Request.QueryString["did"].ToString());
            SortedList list = new SortedList();
            list.Add("@DomainId", domainId);
            ViewData["DomainId"] = domainId;
            System.Data.DataTable dt = comfun.fillDataTable("DbSetUpCheckIfDone", "", list, conLogin);
            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["IsDatabaseCreated"].ToString() == "N")
                {
                    ViewData["DB"] = 0;
                    string mes = "";
                    try
                    {
                        string db = "dbsHTIS_wfms" + commonFunctions.getSubDomain();

                        FileInfo file = new FileInfo(@"D:\Webserver\wfmbase.htistelecom.in\DB\script-tables.sql");
                        FileInfo file2 = new FileInfo(@"D:\Webserver\wfmbase.htistelecom.in\DB\script-proc.sql");
                        FileInfo file3 = new FileInfo(@"D:\Webserver\wfmbase.htistelecom.in\DB\script-table-types.sql");
                        //FileInfo file4 = new FileInfo(@"D:\Webserver\wfmbase.htistelecom.in\DB\script-default-entry.sql");

                        //FileInfo file = new FileInfo(@"D:\vss_work\HTIS\Wfms 2\Payroll\DB\script-tables.sql");
                        //FileInfo file2 = new FileInfo(@"D:\vss_work\HTIS\Wfms 2\Payroll\DB\script-proc.sql");
                        //FileInfo file3 = new FileInfo(@"D:\vss_work\HTIS\Wfms 2\Payroll\DB\script-table-types.sql");
                        //FileInfo file4 = new FileInfo(@"D:\vss_work\HTIS\Wfms 2\Payroll\DB\script-default-entry.sql");

                        //FileInfo file = new FileInfo(@"D:\vss_work\HTIS\Wfms 2\Payroll\DB\script.sql");

                        string baseDB = ConfigurationManager.AppSettings["BaseDb"].ToString();
                        string scriptTables = file.OpenText().ReadToEnd();
                        scriptTables = scriptTables.Replace("##BaseDB##", baseDB);
                        scriptTables = scriptTables.Replace("##DbName##", db);

                        string scriptProc = file2.OpenText().ReadToEnd();
                        scriptProc = scriptProc.Replace("##BaseDB##", baseDB);
                        scriptProc = scriptProc.Replace("##DbName##", db);

                        string scriptTableTypes = file3.OpenText().ReadToEnd();
                        scriptTableTypes = scriptTableTypes.Replace("##BaseDB##", baseDB);
                        scriptTableTypes = scriptTableTypes.Replace("##DbName##", db);



                        if (ConfigurationManager.AppSettings.AllKeys.Contains(db))
                        {
                            string subject = "WFMS- connection done";
                            subject = comfun.SendEmail("WFMS connection", "kuldeep.singh@horizontelecom.in", subject, subject);
                            SqlConnection newDBCon = new SqlConnection(ConfigurationManager.AppSettings[db].ToString());

                            if (newDBCon == null)
                            {
                                subject = "WFMS- connection null";
                                subject = comfun.SendEmail("WFMS connection null", "kuldeep.singh@horizontelecom.in", subject, subject);
                                mes = "Error";
                            }
                            if (newDBCon.State != ConnectionState.Open)
                            {
                                try
                                {
                                    newDBCon.Open();
                                    subject = "WFMS- connection open";
                                    newDBCon.Close();
                                    subject = comfun.SendEmail("WFMS connection open", "kuldeep.singh@horizontelecom.in", subject, subject);
                                }
                                catch (Exception ex)
                                {
                                    subject = "WFMS- connection error";
                                    subject = comfun.SendEmail("WFMS connection error", "kuldeep.singh@horizontelecom.in", subject, ex.Message);
                                    mes = "Error";
                                }
                            }
                            try
                            {
                                string qry = "update tblDomain set IsDatabaseCreated='Y' where DomainId=" + domainId;
                                comfun.executeNonQuery("", scriptTables, null, newDBCon); //Tables
                                comfun.executeNonQuery("", scriptTableTypes, null, newDBCon); //Tables Types
                                comfun.executeNonQuery("", scriptProc, null, newDBCon);//Stored Proc  & Functions

                                comfun.executeNonQuery("", qry, null, conLogin);//set DB Create TAG
                                mes = "Done";
                            }
                            catch (Exception ex)
                            {
                                subject = "WFMS-acc redirec Exception";
                                subject = comfun.SendEmail("WFMS script exe error", "kuldeep.singh@horizontelecom.in", subject, ex.Message);
                                mes = "Error";
                            }
                        }
                        else
                        {
                            string subject = "WFMS- cant find connection to new db";
                            subject = comfun.SendEmail("WFMS connection error", "kuldeep.singh@horizontelecom.in", subject, subject);
                            mes = "Error";
                        }
                        // mes = "done";
                    }
                    catch (Exception ex)
                    {
                        mes = "Error: " + ex.Message;
                    }
                }
                else
                {
                    ViewData["DB"] = 1;
                }
            }
            else
            {
                ViewData["DB"] = 0;
            }



            // return View();

            return RedirectToAction("Index", "Home");
        }

        [AllowAnonymous]
        public string DatabaseSetup(string domainId, string domain)
        {
            string db = "";
            try
            {
                db = "dbsHTIS_wfms" + commonFunctions.getSubDomain(domain);
            }
            catch (Exception ex)
            {

            }
            string mes = string.Empty;
            SortedList list = new SortedList();
            list.Add("@DomainId", domainId);
            list.Add("@DataBase", db);
            ViewData["DomainId"] = domainId;
            System.Data.DataTable dt = comfun.fillDataTable("DbSetUpCheckIfDone", "", list, conLogin);
            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["IsDatabaseCreated"].ToString() == "N")
                {

                    try
                    {
                        db = "dbsHTIS_wfms" + commonFunctions.getSubDomain(domain);

                        FileInfo file = new FileInfo(@"C:\Webserver\wfmbase.htistelecom.in\DB\script-tables.sql");
                        FileInfo file2 = new FileInfo(@"C:\Webserver\wfmbase.htistelecom.in\DB\script-proc.sql");
                        FileInfo file3 = new FileInfo(@"C:\Webserver\wfmbase.htistelecom.in\DB\script-table-types.sql");
                        //FileInfo file4 = new FileInfo(@"D:\Webserver\wfmbase.htistelecom.in\DB\script-default-entry.sql");

                        //FileInfo file = new FileInfo(@"D:\vss_work\HTIS\Wfms 2\Payroll\DB\script-tables.sql");
                        //FileInfo file2 = new FileInfo(@"D:\vss_work\HTIS\Wfms 2\Payroll\DB\script-proc.sql");
                        //FileInfo file3 = new FileInfo(@"D:\vss_work\HTIS\Wfms 2\Payroll\DB\script-table-types.sql");
                        //FileInfo file4 = new FileInfo(@"D:\vss_work\HTIS\Wfms 2\Payroll\DB\script-default-entry.sql");

                        //FileInfo file = new FileInfo(@"D:\vss_work\HTIS\Wfms 2\Payroll\DB\script.sql");

                        string baseDB = ConfigurationManager.AppSettings["BaseDb"].ToString();
                        string scriptTables = file.OpenText().ReadToEnd();
                        scriptTables = scriptTables.Replace("##BaseDB##", baseDB);
                        scriptTables = scriptTables.Replace("##DbName##", db);

                        string scriptProc = file2.OpenText().ReadToEnd();
                        scriptProc = scriptProc.Replace("##BaseDB##", baseDB);
                        scriptProc = scriptProc.Replace("##DbName##", db);

                        string scriptTableTypes = file3.OpenText().ReadToEnd();
                        scriptTableTypes = scriptTableTypes.Replace("##BaseDB##", baseDB);
                        scriptTableTypes = scriptTableTypes.Replace("##DbName##", db);

                        string cn = @"data source=15.207.246.29,1533; initial catalog=" + db + "; user id=sa; password=data@5_htis#9002; connection timeout=0;pooling=true;Min Pool Size=0;Max Pool Size=100;Pooling=true;";


                        //if (ConfigurationManager.AppSettings.AllKeys.Contains(db))
                        {
                            string subject = "WFMS- connection done";
                            subject = comfun.SendEmail("WFMS connection", "kuldeep.singh@horizontelecom.in", subject, subject);
                            SqlConnection newDBCon = new SqlConnection(cn);

                            if (newDBCon == null)
                            {
                                subject = "WFMS- connection null";
                                subject = comfun.SendEmail("WFMS connection null", "kuldeep.singh@horizontelecom.in", subject, subject);
                                mes = "Error";
                            }
                            if (newDBCon.State != ConnectionState.Open)
                            {
                                try
                                {
                                    newDBCon.Open();
                                    subject = "WFMS- connection open";
                                    newDBCon.Close();
                                    subject = comfun.SendEmail("WFMS connection open", "kuldeep.singh@horizontelecom.in", subject, subject);
                                }
                                catch (Exception ex)
                                {
                                    subject = "WFMS- connection error";
                                    subject = comfun.SendEmail("WFMS connection error", "kuldeep.singh@horizontelecom.in", subject, ex.Message);
                                    mes = "Error";
                                }
                            }
                            try
                            {
                                string qry = "update tblDomain set IsDatabaseCreated='Y' where DomainId=" + domainId;
                                comfun.executeNonQuery("", scriptTables, null, newDBCon); //Tables
                                comfun.executeNonQuery("", scriptTableTypes, null, newDBCon); //Tables Types

                                subject = comfun.SendEmail("WFMS scriptProc", "kuldeep.singh@horizontelecom.in", scriptProc, scriptProc);
                                comfun.executeNonQuery("", scriptProc, null, newDBCon);//Stored Proc  & Functions

                                comfun.executeNonQuery("", qry, null, conLogin);//set DB Create TAG
                                mes = "Done";
                            }
                            catch (Exception ex)
                            {
                                subject = "WFMS-acc redirec Exception";
                                subject = comfun.SendEmail("WFMS script exe error", "kuldeep.singh@horizontelecom.in", subject, ex.Message);
                                mes = "Error";
                            }
                        }
                        //else
                        //{
                        //    string subject = "WFMS- cant find connection to new db->"+ db;
                        //    subject = comfun.SendEmail("WFMS connection error", "kuldeep.singh@horizontelecom.in", subject, subject);
                        //    mes = "Error";
                        //}
                        // mes = "done";
                    }
                    catch (Exception ex)
                    {
                        mes = "Error: " + ex.Message;
                    }
                }
            }
            return mes;
        }

        [Authorize]
        public ActionResult Logout()
        {
            Session.RemoveAll();
            Session.Abandon();
            foreach (string key in Request.Cookies.AllKeys)
            {
                HttpCookie c = Request.Cookies[key];
                c.Expires = DateTime.Now.AddMonths(-1);
                Response.AppendCookie(c);
            }
            HttpRuntime.Cache.Remove("dtMenu");
            List<string> keys = new List<string>();
            IDictionaryEnumerator enumerator = HttpRuntime.Cache.GetEnumerator();
            while (enumerator.MoveNext())
            {
                keys.Add(enumerator.Key.ToString());
            }
            for (int i = 0; i < keys.Count; i++)
            {
                HttpRuntime.Cache.Remove(keys[i]);
            }
            FormsAuthentication.SignOut();

            mainUrl = System.Configuration.ConfigurationManager.AppSettings["SiteUrl"].ToString();
            if (mainUrl.ToLower().Contains(commonFunctions.getSubDomain().ToLower()))
            {
                //return RedirectToAction("Login", "Account");
                return RedirectToAction("NewLogin", "Account");
            }
            //return RedirectToAction("Login", "Account");
            return Redirect(mainUrl + "/account/logout");
        }

        [AllowAnonymous]
        public ActionResult SessionExpired()
        {
            Session.RemoveAll();
            Session.Abandon();
            foreach (string key in Request.Cookies.AllKeys)
            {
                HttpCookie c = Request.Cookies[key];
                c.Expires = DateTime.Now.AddMonths(-1);
                Response.AppendCookie(c);
            }
            HttpRuntime.Cache.Remove("dtMenu");
            List<string> keys = new List<string>();
            IDictionaryEnumerator enumerator = HttpRuntime.Cache.GetEnumerator();
            while (enumerator.MoveNext())
            {
                keys.Add(enumerator.Key.ToString());
            }
            for (int i = 0; i < keys.Count; i++)
            {
                HttpRuntime.Cache.Remove(keys[i]);
            }
            FormsAuthentication.SignOut();

            return View("SessionExpired");
        }


        [AllowAnonymous]

        public ActionResult forgot()
        {
            if (Request.QueryString["eid"] == null)
            {
                return RedirectToAction("Login");
            }
            SortedList list = new SortedList();

            // list.Add("@EmpId", Request.QueryString["Email"]);

            list.Add("@EmpId", Request.QueryString["eid"].ToString());

            System.Data.DataTable dt = comfun.fillDataTable("Emp_SelectEmails", "", list);
            Session["Forget"] = Guid.NewGuid();
            ViewData["EmpName"] = Request.QueryString["eid"].ToString();
            return View(dt);
        }

        [AllowAnonymous]
        [HttpPost]
        public JsonResult _sendPassword()
        {
            string mes = "";
            try
            {
                Random r = new Random();

                string password = r.Next(1111, 9999).ToString();
                string empcode = comfun.encryptString(Request.Form["UserName"].ToString());
                string email = comfun.encryptString(Request.Form["Email"].ToString());
                string refDate = comfun.encryptString(DateTime.Now.AddMinutes(30).ToString("dd-MMM-yyyy HH:mm:ss"));
                SortedList list = new SortedList();
                list.Add("@UserCode", Request.Form["UserName"].ToString());
                list.Add("@Email", Request.Form["Email"].ToString());
                list.Add("@Password", password);

                string url = mainUrl + "/Account/passwordreset?refid=" + empcode + "&refe=" + email + "&reftd=" + refDate;
                mes = comfun.executeNonQueryWMessage("User_ResetPassword", "", list).ToString();
                if (!mes.Contains("Error"))
                {
                    StreamReader reader = new StreamReader(Server.MapPath("~/EmailTemplates/forget.htm"));
                    mes = reader.ReadToEnd();

                    mes = mes.Replace("#Link#", url);
                    //comfun.SendEmail("CMR Password", "jagvirjb@gmail.com", "CMR Password", mes);
                    payfun.SendEmail("CMR Password", Request.Form["Email"].ToString(), "CMR Password", mes);
                    if (!mes.Contains("Error"))
                    {
                        mes = "Password sent, Please check your email";
                    }
                }
            }
            catch (Exception ex)
            {
                mes = "Error:" + comfun.errorMessage("Forget password", ex.Message);
                mes = "Error: Please try later";
            }
            return Json(mes);
        }

        public ActionResult ChangePassword()
        {
            return View();
        }
        [AllowAnonymous]
        public JsonResult _Changepassword()
        {
            string empCode = "";
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                if (Request.Form["EmpCode"] == null)
                {
                    empCode = User.Identity.Name;
                }
                else
                {
                    empCode = Request.Form["EmpCode"].ToString();
                }
                list.Add("@EmpCode", empCode);
                list.Add("@UserPassword", Request.Form["Password"].ToString());
                list.Add("@RefDate", Request.Form["RefDate"].ToString());

                mes = comfun.executeNonQueryWMessage("stpViksatChangePassword", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_Changepassword", ex.Message);

            }
            return Json(mes);

        }

        [AllowAnonymous]
        public ActionResult passwordreset()
        {
            if (Request.QueryString["refid"] == null || Request.QueryString["refe"] == null || Request.QueryString["reftd"] == null)
            {
                return View("linkexpired");
            }
            if (Request.QueryString["refid"].ToString() == "" || Request.QueryString["refe"].ToString() == "" || Request.QueryString["reftd"].ToString() == "")
            {
                return View("linkexpired");
            }

            string date = Request.QueryString["reftd"].ToString();
            string empcode = Request.QueryString["refid"].ToString();
            string email = Request.QueryString["refe"].ToString();


            date = comfun.decryptString(date).Trim();
            empcode = comfun.decryptString(empcode).Trim();
            email = comfun.decryptString(email).Trim();
            ViewData["EmpCode"] = empcode.Trim();
            ViewData["RefDate"] = date.Trim();
            DateTime dateTime;
            bool chk = false;
            chk = DateTime.TryParseExact(date,
                                   "dd-MMM-yyyy HH:mm:ss",
                                   CultureInfo.InvariantCulture,
                                   DateTimeStyles.None,
                                   out dateTime);
            if (chk == false)
            {
                return View("linkexpired");
            }

            int empCode = 0;
            chk = int.TryParse(empcode, out empCode);
            if (chk == false)
            {
                return View("linkexpired");
            }
            if (dateTime < DateTime.Now)
            {
                return View("linkexpired");
            }

            return View();
        }
        [AllowAnonymous]
        public ActionResult sessionRestore()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return RedirectToAction("Login", "account");
            }
            return Json("done");
        }
        [AllowAnonymous]
        public JsonResult sessionTimeCheck()
        {
            return Json(Session.Timeout);
        }
        [AllowAnonymous]
        public ActionResult sessionDestroy()
        {
            Session.RemoveAll();
            Session.Abandon();

            foreach (string key in Request.Cookies.AllKeys)
            {
                HttpCookie c = Request.Cookies[key];
                c.Expires = DateTime.Now.AddMonths(-1);
                Response.AppendCookie(c);
            }
            HttpRuntime.Cache.Remove("dtMenu");
            FormsAuthentication.SignOut();
            return PartialView("_sessionExpired");
        }
        public ActionResult ProjectunderShift()
        {
            return View();
        }

        [AllowAnonymous]
        //[OutputCache(NoStore = true, Duration = 0, VaryByParam = "None")]
        public ActionResult SignUp()
        {
            string mes = "";
            ViewData["error"] = "";
            return Redirect(mainUrl + "/account/signup");
        }

        // [Route("/")]
        public ActionResult IndexBase()
        {
            string host = Request.Url.Host;
            SortedList list = new SortedList();
            list.Add("@websiteTitle", "website");
            System.Data.DataTable dt = comfun.fillDataTable("WebsiteTemplate", "", list);
            ViewData["host"] = host;
            ViewData["Login"] = true;
            return View(dt);
        }

        ServerManager serverMgr = new ServerManager();
        [AllowAnonymous]
        public ActionResult AddWeb()
        {
            string msg = "";
            try
            {
                //string strWebsitename = Request.Form["Web"].ToString()+ ".earnmoney.blog"; // abc 
                string strApplicationPool = "DefaultAppPool";  // set your deafultpool :4.0 in IIS
                string strhostname = System.Configuration.ConfigurationManager.AppSettings["HostName"]; //"base.earnmoney.blog"; //abc.com
                string stripaddress = System.Configuration.ConfigurationManager.AppSettings["IpAddress"]; ;// ip address
                string subDomainSuffix = System.Configuration.ConfigurationManager.AppSettings["SubDomainSuffix"]; ;// ip address
                string strWebsitename = Request.Form["Domain"].ToString();
                strWebsitename = strWebsitename + subDomainSuffix;
                Random rnd = new Random();
                string password = rnd.Next(1111, 9999).ToString();
                SortedList list = new SortedList();
                list.Add("@Domain", strWebsitename);
                list.Add("@FirstName", Request.Form["FirstName"].ToString());
                list.Add("@LastName", Request.Form["LastName"].ToString());
                list.Add("@Email", Request.Form["Email"].ToString());
                list.Add("@MobileNo", Request.Form["Phone"].ToString());
                list.Add("@NofEmp", Request.Form["NofEmp"].ToString());
                list.Add("CompanyName", Request.Form["CompanyName"].ToString());
                list.Add("CountryID", Request.Form["CountryID"].ToString());
                list.Add("JobTitle", Request.Form["JobTitle"].ToString());
                list.Add("@Password", password);

                msg = comfun.executeNonQueryWMessage("stpwfmsSignUp_Accept", "", list, conLogin).ToString();
                if (msg.ToLower().Contains("error"))
                {
                    return Json(msg);
                }




                string bindinginfo = stripaddress + ":" + System.Configuration.ConfigurationManager.AppSettings["Port"] + ":" + strWebsitename;

                //check if website name already exists in IIS
                Boolean bWebsite = IsWebsiteExists(strWebsitename);
                if (!bWebsite)
                {
                    Site mySite = serverMgr.Sites.Add(strWebsitename.ToString(), "http", bindinginfo, @"D:\webserver\" + strhostname);
                    mySite.ApplicationDefaults.ApplicationPoolName = strApplicationPool;
                    mySite.TraceFailedRequestsLogging.Enabled = true;
                    mySite.TraceFailedRequestsLogging.Directory = @"D:\webserver\" + strWebsitename;
                    serverMgr.CommitChanges();
                    msg = "Done";
                }
                else
                {
                    msg = "Error: domain already exists";
                }

                if (msg.ToLower().Contains("done"))
                {
                    //msg = "Website created successfully";
                    string name = Request.Form["FirstName"].ToString() + " " + Request.Form["LastName"].ToString();
                    string email = Request.Form["Email"].ToString();
                    msg = sendWelcomeEmail(email, name, password, "0");
                }
            }
            catch (Exception ex)
            {
                msg = "Error:" + ex.Message;
            }
            return Json(msg);
        }

        public bool IsWebsiteExists(string strWebsitename)
        {
            Boolean flagset = false;
            SiteCollection sitecollection = serverMgr.Sites;
            foreach (Site site in sitecollection)
            {
                if (site.Name == strWebsitename.ToString())
                {
                    flagset = true;
                    break;
                }
                else
                {
                    flagset = false;
                }
            }
            return flagset;
        }

        private string sendWelcomeEmail(string email, string name, string password, string domainId)
        {

            string mes = "";
            string subject = "Welcome to WFMS";

            HtmlString htmlString = new HtmlString(mes);

            StreamReader reader = new StreamReader(Server.MapPath("~/EmailTemplates/welcome.htm"));
            mes = reader.ReadToEnd();

            mes = mes.Replace("#EmployeeName#", name);
            mes = mes.Replace("#Password#", password);
            mes = mes.Replace("#Email#", email);
            mes = mes.Replace("#Portal#", "wfms.htistelecom.in");

            string link = mainUrl + "/account/confirmemail?code=" + comfun.encryptString(email).Trim() + "&domain=" + comfun.encryptString(domainId).Trim() + "&p=" + comfun.encryptString(password).Trim();

            mes = mes.Replace("#Link#", link);

            htmlString = new HtmlString(mes);
            reader.Close();
            reader.Dispose();

            mes = comfun.SendEmail("WFMS", email, subject, mes);
            if (!mes.ToLower().Contains("error"))
            {
                mes = "Your account has been created successfully, Please check your inbox to verify your Email Address";
            }

            return mes;
        }
        [AllowAnonymous]
        public ActionResult confirmemail()
        {
            string db = "";
            string cn = "";
            string scriptDefaultEntry = "";
            string mes = "";
            string firstName = "", lastName = "";
            try
            {
                SortedList list = new SortedList();
                list.Add("@Email", comfun.decryptString(Request.QueryString["code"].ToString().Trim()).Trim());
                list.Add("@DomainId", comfun.decryptString(Request.QueryString["domain"].ToString().Trim()).Trim());
                list.Add("@Password", comfun.decryptString(Request.QueryString["p"].ToString().Trim()).Trim());

                mes = comfun.executeNonQueryWMessage("AppEmail_verifyJ", "", list, conLogin).ToString();

                if (!mes.ToLower().Contains("error"))
                {
                    ViewData["Mes"] = "Your Email has been verified successfully!";
                    HtmlString html = new HtmlString("<a href='" + mainUrl + "'>Click here</a> to continue login");
                    ViewData["Login"] = html;

                    /*****************
                    *****Connection setup****
                    ****************/
                    string[] d = mes.Split('~');
                    SetAppSettings(d[0]);
                    //0=DomainName,1=DomainId,2=FirstName,3=LastName,4=RoleId

                    DatabaseSetup(d[1], d[0]);

                    firstName = d[2];
                    lastName = d[3];

                    try
                    {
                        db = "dbsHTIS_wfms" + commonFunctions.getSubDomain(d[0]);
                        cn = @"data source=15.207.246.29,1533; initial catalog=" + db + "; user id=sa; password=data@5_htis#9002; connection timeout=0;pooling=true;Min Pool Size=0;Max Pool Size=100;Pooling=true;";

                        SqlConnection newDBCon = new SqlConnection(cn);
                        //string baseDB = ConfigurationManager.AppSettings["BaseDb"].ToString();
                        string domainId = d[1];
                        if (Request.QueryString["domain"] != null)
                        {
                            domainId = comfun.decryptString(Request.QueryString["domain"].ToString().Trim()).Trim();
                        }

                        FileInfo file4 = new FileInfo(@"D:\Webserver\wfmbase.htistelecom.in\DB\script-default-entry.sql");
                        scriptDefaultEntry = file4.OpenText().ReadToEnd();
                        scriptDefaultEntry = scriptDefaultEntry.Replace("##FirstName##", firstName);
                        scriptDefaultEntry = scriptDefaultEntry.Replace("##LastName##", lastName);
                        scriptDefaultEntry = scriptDefaultEntry.Replace("##FullName##", firstName + " " + lastName);
                        scriptDefaultEntry = scriptDefaultEntry.Replace("##DomainId##", domainId);
                        scriptDefaultEntry = scriptDefaultEntry.Replace("##RoleId##", d[4]);

                        scriptDefaultEntry = scriptDefaultEntry.Replace("##Email##", comfun.decryptString(Request.QueryString["code"].ToString()).Trim());

                        mes = comfun.SendEmail("WFMS create user function info", "kuldeep.singh@horizontelecom.in", mes, scriptDefaultEntry + "<br><br>" + cn + "<br><br>" + db + "<br><br>" + scriptDefaultEntry);
                        comfun.executeNonQuery("", scriptDefaultEntry, null, newDBCon);//User Setup
                    }
                    catch (Exception ex)
                    {
                        mes = "WFMS- confirmemail error";
                        mes = comfun.SendEmail("WFMS connection error", "kuldeep.singh@horizontelecom.in", mes, ex.Message + "<br><br>" + ex.InnerException + "<br><br>" + cn + "<br><br>" + db + "<br><br>" + scriptDefaultEntry);
                        mes = "Error";
                    }
                }
                else
                {
                    ViewData["Mes"] = "Email can not be verified";
                    ViewData["Login"] = "";
                }
            }
            catch (Exception ex)
            {
                mes = "Error:" + ex.Message;
                mes = mes + ", Email not verified";
            }
            return View();
        }

        private void createDB(string dbName)
        {
            FileInfo file = new FileInfo("C:\\script.sql");
            string script = file.OpenText().ReadToEnd();
            script = script.Replace("##DbName##", dbName);
            comfun.executeNonQuery("", script, null);
        }

        public ActionResult CreateUser()
        {
            comfun.saveformname("CreateUser", "/Users/CreateUser", "Emp Management/CreateUser", "Create User form", "Y", "CreateUser", "CreateUser", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("CreateUser");
            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("stpRoleddl", null, null);
            return View(dt);
        }


        public JsonResult _Customerddl()
        {
            // comfun.saveformname("_CreateUserList", "/Master/_CreateUserList", "Emp Management/CreateUser", "Create User List", "N", "CreateUser", "CreateUser", "List", 4);

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@EntityId", Request.Form["EntityId"].ToString());
            dt = comfun.fillDataTable("stpCustomerforCreateUser", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _CreateUserList()
        {
            comfun.saveformname("_CreateUserList", "/Master/_CreateUserList", "Emp Management/CreateUser", "Create User List", "N", "CreateUser", "CreateUser", "List", 4);

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@DomainId", Session["DomainId"]);

            dt = comfun.fillDataTable("stpEmployeeUser_list", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _CreateUserSubmit()
        {
            comfun.saveformname("_CreateUserSubmit", "/Users/_CreateUserSubmit", "Emp Management/Create User", "Create User  Add", "N", "CreateUser", "CreateUser", "Add", 2);
            comfun.saveformname("_CreateUserSubmit", "/Users/_CreateUserSubmit", "Emp Management/Create User", "Create User  Edit", "N", "CreateUser", "CreateUser", "Edit", 3);

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                SqlConnection conLogin = new SqlConnection(ConfigurationManager.ConnectionStrings["cnLogin"].ConnectionString);
                Random rnd = new Random();
                string password = rnd.Next(1111, 9999).ToString();
                list.Add("@DomainId", Request.Cookies["DomainId"].Value);
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@FirstName", Request.Form["FirstName"].ToString());
                list.Add("@LastName", Request.Form["LastName"].ToString());
                list.Add("@RoleId", Request.Form["RoleId"].ToString());
                list.Add("@Email", Request.Form["Email"].ToString());
                list.Add("@CustomerId", Request.Form["CustomerId"].ToString());
                list.Add("@EntityId", Request.Form["EntityId"].ToString());
                list.Add("@Mobile", Request.Form["Mobile"].ToString());
                list.Add("@Password", password);
                mes = comfun.executeNonQueryWMessage("stpwfmsSignUp_AddDomainUser", "", list, conLogin).ToString();
                if (!mes.Contains("Error"))
                {
                    sendWelcomeEmail(Request.Form["Email"].ToString(), Request.Form["FirstName"].ToString() + " " + Request.Form["LastName"].ToString(), mes, Request.Cookies["DomainId"].Value);
                }
            }

            catch (Exception ex)
            {

                mes = "Error:" + ex.Message;
            }

            return Json(mes);
        }


        public JsonResult _CreateUserStatusUpdate()
        {
            comfun.saveformname("_CreateUserStatusUpdate", "/Users/_CreateUserStatusUpdate", "Emp Management/CreateUser", "Create User Status", "N", "CreateUser", "CreateUser", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpEmployeeUserStatus", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }
        public JsonResult _CreateUserEdit()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("stpEmployeeUser_Edit", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }


        /////////*****Eliglibity*****///////
        public ActionResult Eligibility()
        {

            return View();
        }
        public JsonResult _EligibilityList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("stpEligibilityList", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);


        }

        public JsonResult _EligibilitySubmit()
        {
            //comfun.saveformname("_CreateUserSubmit", "/Users/_CreateUserSubmit", "Emp Management/Create User", "Create User  Add", "N", "CreateUser", "CreateUser", "Add", 2);
            // comfun.saveformname("_CreateUserSubmit", "/Users/_CreateUserSubmit", "Emp Management/Create User", "Create User  Edit", "N", "CreateUser", "CreateUser", "Edit", 3);

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {

                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@ElCode", Request.Form["Code"].ToString());
                list.Add("@Eligibility", Request.Form["Eliglibilty"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());

                mes = comfun.executeNonQueryWMessage("stpEligibilityAccept", "", list).ToString();

            }

            catch (Exception ex)
            {

                mes = "Error:" + ex.Message;
            }

            return Json(mes);
        }
        public JsonResult _EligibilityEdit()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("stpEligibilityList", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _EligibilityDelete()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["RecordId"].ToString());
            String mes = comfun.executeNonQueryWMessage("stpEligibility_delete", "", list).ToString();


            var json = JsonConvert.SerializeObject(mes);
            return Json(json);
        }


        //****Course*****//
        public ActionResult Course()
        {
            return View();
        }
        public JsonResult _EligibilityList2()

        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("stpEligibilityList", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);


        }

        public JsonResult _EligibilitySubmit2()
        {
            //comfun.saveformname("_CreateUserSubmit", "/Users/_CreateUserSubmit", "Emp Management/Create User", "Create User  Add", "N", "CreateUser", "CreateUser", "Add", 2);
            // comfun.saveformname("_CreateUserSubmit", "/Users/_CreateUserSubmit", "Emp Management/Create User", "Create User  Edit", "N", "CreateUser", "CreateUser", "Edit", 3);

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {

                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@ElCode", Request.Form["Code"].ToString());
                list.Add("@Eligibility", Request.Form["Eliglibilty"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());

                mes = comfun.executeNonQueryWMessage("stpEligibilityAccept", "", list).ToString();

            }

            catch (Exception ex)
            {

                mes = "Error:" + ex.Message;
            }

            return Json(mes);
        }
        public JsonResult _EligibilityEdit2()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("stpEligibilityList", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }


        //****Bank***///
        public ActionResult Bank()
        {
            return View();
        }
        public JsonResult _BankSubmit()
        {
            //comfun.saveformname("_CreateUserSubmit", "/Users/_CreateUserSubmit", "Emp Management/Create User", "Create User  Add", "N", "CreateUser", "CreateUser", "Add", 2);
            // comfun.saveformname("_CreateUserSubmit", "/Users/_CreateUserSubmit", "Emp Management/Create User", "Create User  Edit", "N", "CreateUser", "CreateUser", "Edit", 3);

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {

                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@ElCode", Request.Form["Code"].ToString());
                list.Add("@Eligibility", Request.Form["Eliglibilty"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());

                mes = comfun.executeNonQueryWMessage("stpEligibilityAccept", "", list).ToString();

            }

            catch (Exception ex)
            {

                mes = "Error:" + ex.Message;
            }

            return Json(mes);
        }
        public JsonResult _BankList()

        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("stpEligibilityList", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);


        }

        #region PiLogin
        public JsonResult _SchoolRegistrationSubmit()
        {
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@ID", "0");
                list.Add("@Name", Request.Form["Name"].ToString());
                list.Add("@emailId", Request.Form["email"].ToString());
                list.Add("@ContactNO", Request.Form["Contact"].ToString());
                list.Add("@Orgname", Request.Form["organization"].ToString());

                mes = comfun.executeNonQueryWMessage("PiSchoolLoginRequist_AcceptUpdate", "", list).ToString();
            }

            catch (Exception ex)
            {

                mes = "Error:" + ex.Message;
            }

            return Json(mes);
        }
        #endregion

        #region RefreshSession
        [Authorize]
        [HttpPost]
        public JsonResult KeepAlive()
        {
            return Json(new { status = "success", message = "Session renewed." });
        }
        #endregion

    }
}
