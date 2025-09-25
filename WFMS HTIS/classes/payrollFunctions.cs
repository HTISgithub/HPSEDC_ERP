
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using System.Web.UI.WebControls;
using System.Net;
using System.Net.Mail;
using System.Linq;
using System.Collections.Generic;
using cls_Encrypt;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Xml.Linq;
using System.Text;
using System.Net.Mime;

using System.Threading;

using System.ComponentModel;

using System.Text;

using System.Web.Util;
using System.IO;
using Microsoft.Web.Administration;

public class payrollFunctions
{
    commonFunctions comfun = new commonFunctions();
    public const string _rolesGlobal = "SA,SuperAdmin,Manager,Employee,PayrollTeam,Auditor,Accounts";
    public const string roleManger = "Manager";
    public const string rolePayrollTeam = "PayrollTeam";
    public const string roleAdmin = "SuperAdmin";
    public const string roleEmployee = "Employee";
    public const string roleAuditor = "Auditor";
    public const string roleSa = "SA";
    public const string roleRegionalHr = "RegionalHr";

    string siteUrl = System.Configuration.ConfigurationManager.AppSettings["SiteUrl"];
    public static string sessionIdFixed = "3";
    public static string sessionNameFixed = "2019-20";
    double sessionTimeInMin = Convert.ToDouble(System.Configuration.ConfigurationManager.AppSettings["SessionTimeOut"]);
    double directUrlTimeOut = Convert.ToDouble(System.Configuration.ConfigurationManager.AppSettings["DirectUrlTimeOut"]);

    public string sessionRecreate()
    {
        string str = "Done";

        if (checkIfSessionExists() == "expires")
        {
            if (checkifCookiesExists() == "expires")
            {
                str = setValuesUsingTable();
            }
            else
            {
                setValuesInSessionUsingCookies();
            }
        }
        else
        {
            str = "Done";
        }

        if (str == "expires")
        {
            HttpContext.Current.Session.RemoveAll();
            foreach (string key in HttpContext.Current.Request.Cookies.AllKeys)
            {
                HttpCookie c = HttpContext.Current.Request.Cookies[key];
                c.Expires = DateTime.Now.AddMonths(-1);
                HttpContext.Current.Response.AppendCookie(c);
            }
            HttpRuntime.Cache.Remove("dtMenu");

            FormsAuthentication.SignOut();
        }
        return str;
    }

    private string checkifCookiesExists()
    {
        foreach ( var item in cookieArrayList())
        {
            if (getCookieValueOrDefault(item.ToString()) == "")
            {
                return "expires";
            }
        }
        return "Y";
    }

    private string checkIfSessionExists()
    { 
        foreach (var item in cookieArrayList())
        {
            if (HttpContext.Current.Session[item.ToString()] == null)
            {
                return "expires";
            }
        }
        return "Y";
    }

    private void setValuesInSessionUsingCookies()
    {

        foreach (var item in cookieArrayList())
        {
            if (HttpContext.Current.Request.Cookies[item.ToString()] != null)
            {
                HttpContext.Current.Session[item.ToString()] = HttpContext.Current.Request.Cookies[item.ToString()].Value;
            }
        }
    }

    public void setValuesInCookiesUsingSession()
    {
        foreach (var item in cookieArrayList())
        {
            if (HttpContext.Current.Session[item.ToString()] != null)
            {
                HttpContext.Current.Response.Cookies[item.ToString()].Value = HttpContext.Current.Session[item.ToString()].ToString();
                HttpContext.Current.Response.Cookies[item.ToString()].Expires = DateTime.Now.AddMinutes(sessionTimeInMin);
            }
        }
    }
    

    private string setValuesUsingTable()
    {
        if (HttpContext.Current.User.Identity.Name != "")
        {
            SortedList list = new SortedList();
            list.Add("@EmpCode", HttpContext.Current.User.Identity.Name);
            DataTable dt = comfun.fillDataTable("Htis_LoginDetail", "", list);
            if (dt.Rows.Count > 0)
            {
                HttpContext.Current.Session["EmpId"] = dt.Rows[0]["Id"].ToString();
                HttpContext.Current.Session["EmpName"] = dt.Rows[0]["EmpName"].ToString();
                HttpContext.Current.Session["EmpCode"] = dt.Rows[0]["EmpCode"].ToString();

                HttpContext.Current.Session["CompanyId"] = dt.Rows[0]["CompanyId"].ToString();
                HttpContext.Current.Session["CompanyName"] = dt.Rows[0]["CompanyName"].ToString();

                HttpContext.Current.Session["CostCenterId"] = dt.Rows[0]["CostCenterId"].ToString();
                HttpContext.Current.Session["CostCenterName"] = dt.Rows[0]["CostCenterName"].ToString();

                HttpContext.Current.Session["LocationId"] = dt.Rows[0]["LocationId"].ToString();
                HttpContext.Current.Session["LocationName"] = dt.Rows[0]["LocationName"].ToString();

                HttpContext.Current.Session["EmpCategoryId"] = dt.Rows[0]["EmpCategoryId"].ToString(); 

                HttpContext.Current.Session["SessionId"] = sessionIdFixed;
                HttpContext.Current.Session["SessionName"] = sessionNameFixed;// dt.Rows[0]["SessionId"].ToString();  

                HttpContext.Current.Session["RoleName"] = Enum.Parse(typeof(payrollFunctions.roleIds), "Employee").ToString();
                HttpContext.Current.Session["RoleId"] = ((int)Enum.Parse(typeof(payrollFunctions.roleIds), "Employee")).ToString();

                setValuesInCookiesUsingSession();
            }
            else
            {
                return "expires";
            }
        }
        else
        {
            return "expires";
        }
        return "Y";
    }


    public DataTable companySelectByEmp()
    {
        DataTable dt = new DataTable();
        return dt;
    }

    public DataTable employeeSearch(string txt, string roleName)
    {
        SortedList list = new SortedList();
        list.Clear();
        list.Add("@LoginEmpId", HttpContext.Current.Session["EmpId"].ToString());
        list.Add("@keyword", txt);
        list.Add("@Role", roleName);
        if (roleName != "SuperAdmin")
        {
            list.Add("@CostCenterId", HttpContext.Current.Session["CostCenterId"].ToString());
        } 
         
        DataTable dt = comfun.fillDataTable("Emp_search", "", list);
        return dt;
    }


    public Boolean IsExistsDepartment(string txt)
    {
        SortedList list = new SortedList();
        list.Clear();
        list.Add("@DepartmentCode", txt);
        DataTable dt = comfun.fillDataTable("Department_search", "", list);
        if (dt.Rows.Count > 0)
        {
            return true;
        }
        else
        {
            return false;
        }
    }
    public Boolean IsExistsDesignation(string txt)
    {
        SortedList list = new SortedList();
        list.Clear();
        list.Add("@DesignationCode", txt);
        DataTable dt = comfun.fillDataTable("Designation_Search", "", list);
        if (dt.Rows.Count > 0)
        {
            return true;
        }
        else
        {
            return false;
        }
    }

    public Boolean IsExistsEmployee(string txt)
    {
        SortedList list = new SortedList();
        list.Clear();
        list.Add("@EmpCode", txt);
        DataTable dt = comfun.fillDataTable("Employee_SearchByCode", "", list);
        if (dt.Rows.Count > 0)
        {
            return true;
        }
        else
        {
            return false;
        }
    }
    public Boolean IsExistsGrade(string txt)
    {
        SortedList list = new SortedList();
        list.Clear();
        list.Add("@GradeCode", txt);
        DataTable dt = comfun.fillDataTable("Grade_SearchByCode", "", list);
        if (dt.Rows.Count > 0)
        {
            return true;
        }
        else
        {
            return false;
        }
    }
    public Boolean IsExistsCompany(string txt)
    {
        SortedList list = new SortedList();
        list.Clear();
        list.Add("@CompanyCode", txt);
        DataTable dt = comfun.fillDataTable("Company_SearchByCode", "", list);
        if (dt.Rows.Count > 0)
        {
            return true;
        }
        else
        {
            return false;
        }
    }

    public Boolean IsExistsEmpCategory(string txt)
    {
        SortedList list = new SortedList();
        list.Clear();
        list.Add("@EmpCategoryCode", txt);
        DataTable dt = comfun.fillDataTable("EmpCategory_SearchByCode", "", list);
        if (dt.Rows.Count > 0)
        {
            return true;
        }
        else
        {
            return false;
        }
    }
    public Boolean IsExistsBank(string txt)
    {
        SortedList list = new SortedList();
        list.Clear();
        list.Add("@BankName", txt);
        DataTable dt = comfun.fillDataTable("Bank_SearchByCode", "", list);
        if (dt.Rows.Count > 0)
        {
            return true;
        }
        else
        {
            return false;
        }
    }
    public Boolean IsExistsCostCentre(string txt)
    {
        SortedList list = new SortedList();
        list.Clear();
        list.Add("@CostCentreCode", txt);
        DataTable dt = comfun.fillDataTable("CostCentre_SearchByCode", "", list);
        if (dt.Rows.Count > 0)
        {
            return true;
        }
        else
        {
            return false;
        }
    }
    public Boolean IsExistsDivision(string txt)
    {
        SortedList list = new SortedList();
        list.Clear();
        list.Add("@DivisionCode", txt);
        DataTable dt = comfun.fillDataTable("Division_SearchByCode", "", list);
        if (dt.Rows.Count > 0)
        {
            return true;
        }
        else
        {
            return false;
        }
    }
    public Boolean IsExistsLocation(string txt)
    {
        SortedList list = new SortedList();
        list.Clear();
        list.Add("@LocationCode", txt);
        DataTable dt = comfun.fillDataTable("Location_SearchByCode", "", list);
        if (dt.Rows.Count > 0)
        {
            return true;
        }
        else
        {
            return false;
        }
    }
    public Boolean IsExistsEmployeeId(string txt)
    {
        SortedList list = new SortedList();
        list.Clear();
        list.Add("@EmpCode", txt);
        DataTable dt = comfun.fillDataTable("stpviksatSearchbyEmpCode", "", list);
        if (dt.Rows.Count > 0)
        {
            return true;
        }
        else
        {
            return false;
        }
    }
    public string getCookieValueOrDefault(string cookieName)
    {
        HttpCookie cookie = HttpContext.Current.Request.Cookies[cookieName];
        if (cookie == null)
        {
            return "";
        }
        else
        if (cookie.Value == "")
        {
            return "";
        }
        return cookie.Value;
    }
    public string getQueryStringValueOrDefault(string qsName)
    {

        if (HttpContext.Current.Request.QueryString[qsName] == null)
        {
            return "";
        }
        else
        if (HttpContext.Current.Request.QueryString[qsName].ToString() == "")
        {
            return "";
        }
        return HttpContext.Current.Request.QueryString[qsName].ToString();
    }
    public ArrayList cookieArrayList()
    {
        ArrayList alist = new ArrayList();
        alist.Add("EmpId");
        alist.Add("EmpName");
        alist.Add("EmpCode");
        //alist.Add("CostCenterId");
        //alist.Add("CostCenterName");
        //alist.Add("LocationId");
        //alist.Add("CompanyId");
        //alist.Add("CompanyName");
    
        alist.Add("SessionId");
        alist.Add("SessionName");
        alist.Add("RoleName");
        alist.Add("RoleId");
        alist.Add("EmpCategoryId");
        alist.Add("DomainId");
        alist.Add("AmcId");
        return alist;
    }

    public enum MonthNames
    {
        Jan = 1,
        Feb,
        Mar,
        Apr,
        May,
        Jun,
        Jul,
        Aug,
        Sep,
        Oct,
        Nov,
        Dec
    }

    public enum DayNames
    {
        Sunday = 1,
        Monday,
        Tuesday,
        Wednesday,
        Thursday,
        Friday,
        Saturday
    }

    public enum WeekAlias
    {
        First= 1,
        Second,
        Third,
        Fourth,
        Last 
    }

    public enum roleIds
    {

        SuperAdmin = 2,
        PayrollTeam = 4,
        Employee = 3,
        Manager = 5,
        Auditor = 6
    }
    public enum clearanceIds
    {

        Saleorder = 1,
        Claim = 2,
        Attedance = 3,
        salary = 4,
        Onboard = 5,
        FullFinal=6
    }
    //public enum AccountsHeadId
    //{

    //    CapitalAccount = 1,
    //    PayrollTeam = 4,
    //    Employee = 3,
    //    Manager = 5,
    //    Auditor = 6
    //}

    public static string captalize(string s)
    {
        if (string.IsNullOrWhiteSpace(s))
        {
            return s;
        }

        return s.ToLower();
    }

  
    public string LeaveApply(DateTime Applydate)
    {
        DateTime Date = DateTime.Now;        
        DateTime CurrentMonthDate = new DateTime(Date.Year, Date.Month, 2);
        Date = Date.AddMonths(-1);
        DateTime OldMonthDate = new DateTime(Date.Year, Date.Month,2);
        if (Applydate < OldMonthDate)
        {
            return "N";
        }

        if (Applydate < CurrentMonthDate && Applydate.Month < CurrentMonthDate.Month)
        {
            return "N";
        }

        return "Y";
    }

    public static string ExcelContentType
    {
        get
        { return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"; }
    }

    public SortedList globalParameterList(FormCollection form,SortedList list)
    { 
        string keyname;
        string keyvalue;

        for (int i = 0; i <= form.Count - 1; i++)
        {
            keyname = form.AllKeys[i];
            keyvalue = form[i];

            if(!keyname.StartsWith("@"))
            {
                keyname = "@"+keyname;
            }

            if (keyvalue != "" && keyvalue.ToString() != "0")
            {
                if (!list.ContainsKey(keyname) && keyname!="@Proc")
                {
                    list.Add(keyname, keyvalue); 
                }
            }
        } 
        return list;
    }


     
    public string SendEmail(string displayName, string toEmail, string subject, string message)
    {
        string m = "";
        try
        {
            

            string email = "Wfms@horizontelecom.in";
            string password = "wfms@9002"; ;
            displayName = "WFMS";

            var loginInfo = new NetworkCredential(email, password);
            var msg = new MailMessage();
            var smtpClient = new SmtpClient("smtp.rediffmailpro.com");
            //var smtpClient = new SmtpClient("smtp.gmail.com");
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtpClient.Port = 587;

            msg.From = new MailAddress("hrms@wfms.in", displayName);
            msg.To.Add(new MailAddress(toEmail));
            //msg.To.Add(new MailAddress("jagvirjb@gmail.com"));
            //msg.Bcc.Add(new MailAddress("htis.comp111@gmail.com"));
            //msg.Bcc.Add(new MailAddress("htis.comp111@gmail.com"));
            //msg.CC.Add(new MailAddress("info@horizontelecom.in"));
            //msg.CC.Add(new MailAddress("anish.dhiman@horizontelecom.in"));
            HtmlString htmlString = new HtmlString(message);
     //       AlternateView htmlView =
     //AlternateView.CreateAlternateViewFromString(message, Encoding.UTF8, "text/html");
     //       msg.AlternateViews.Add(htmlView); // And a html attachment to make sure.

            msg.Subject = subject;
            msg.Body = message;
            msg.ReplyTo = new MailAddress("hrms@wfms.in");
            msg.Sender = new MailAddress("hrms@wfms.in", displayName);
            msg.IsBodyHtml = true;
          
            smtpClient.EnableSsl = true;
            smtpClient.UseDefaultCredentials = true;
            smtpClient.Credentials = loginInfo;
            smtpClient.Timeout = 600000;
            smtpClient.Send(msg);
            m = "email sent";
        }
        catch (Exception ex)
        {
            m = "Error : " + ex.Message;
        }
        return m;
    }
    public string SendEmailOTP(string displayName, string toEmail, string subject, string message)
    {
        string m = "";
        try
        {
            if (siteUrl.Contains("localhost"))
            {

                toEmail = "kuldeep.wfms@gmail.com";

            }
            //string email = "jagvir.singh@horizontelecom.in";
            //string password = "Htis@@123";


            string email = "Wfms@horizontelecom.in";
            string password = "wfms@9002"; 
            //displayName = "WFMS";

            var loginInfo = new NetworkCredential(email, password);
            var msg = new MailMessage();
            var smtpClient = new SmtpClient("smtp.rediffmailpro.com");
          //  var smtpClient = new SmtpClient("smtp.gmail.com");
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtpClient.Port = 587;

            msg.From = new MailAddress("hrms@wfms.in", displayName);
            msg.To.Add(new MailAddress(toEmail));
            //msg.To.Add(new MailAddress("jagvirjb@gmail.com"));
            //msg.Bcc.Add(new MailAddress("htis.comp111@gmail.com"));
            //msg.Bcc.Add(new MailAddress("htis.comp111@gmail.com"));
            //msg.CC.Add(new MailAddress("info@horizontelecom.in"));
            //msg.CC.Add(new MailAddress("anish.dhiman@horizontelecom.in"));
            HtmlString htmlString = new HtmlString(message);
            //       AlternateView htmlView =
            //AlternateView.CreateAlternateViewFromString(message, Encoding.UTF8, "text/html");
            //       msg.AlternateViews.Add(htmlView); // And a html attachment to make sure.

            msg.Subject = subject;
            msg.Body = message;
            msg.ReplyTo = new MailAddress("hrms@wfms.in");
            msg.Sender = new MailAddress("hrms@wfms.in", displayName);
            msg.IsBodyHtml = true;

            smtpClient.EnableSsl = true;
            smtpClient.UseDefaultCredentials = true;
            smtpClient.Credentials = loginInfo;
            smtpClient.Timeout = 600000;
            smtpClient.Send(msg);
            m = "email sent";
        }
        catch (Exception ex)
        {
            m = "Error : " + ex.Message;
        }
        return m;
    }

    public string SendEmailWithAttach(string displayName, string toEmail, string subject, string message,string attachPath)
    {
        string m = "";
        try
        {
            if (siteUrl.Contains("localhost"))
            {
                toEmail = "kuldeep.viksat@gmail.com";
            }
            //string email = "jagvir.singh@horizontelecom.in";
            //string password = "Htis@@123";

            string email = "Wfms@horizontelecom.in";
            string password = "wfms@9002"; ;
            displayName = "Htis Telecom-HR";

            var loginInfo = new NetworkCredential(email, password);
            var msg = new MailMessage();
            //var smtpClient = new SmtpClient("smtp.rediffmailpro.com");
            var smtpClient = new SmtpClient("smtp.rediffmailpro.com");
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtpClient.Port = 587;

            msg.From = new MailAddress("hrms@wfms.in", displayName);
            msg.To.Add(new MailAddress(toEmail));

            //msg.To.Add(new MailAddress("jagvirjb@gmail.com"));
            //msg.Bcc.Add(new MailAddress("htis.comp111@gmail.com"));
            //msg.Bcc.Add(new MailAddress("htis.comp111@gmail.com"));

            HtmlString htmlString = new HtmlString(message);
            
            //AlternateView htmlView =
            //AlternateView.CreateAlternateViewFromString(message, Encoding.UTF8, "text/html");
            //msg.AlternateViews.Add(htmlView); // And a html attachment to make sure.

            msg.Subject = subject;
            msg.Body = message;
            msg.ReplyTo = new MailAddress("hrms@wfms.in");
            msg.Sender   = new MailAddress("hrms@wfms.in", displayName);
            msg.IsBodyHtml = true;

            System.Net.Mail.Attachment attachment;
            attachment = new System.Net.Mail.Attachment((HttpContext.Current.Server.MapPath("~/Emaildataupload/") + attachPath));
            //attachment = new System.Net.Mail.Attachment(attachPath);
            msg.Attachments.Add(attachment);
        
            smtpClient.EnableSsl = true;
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = loginInfo;
            smtpClient.Timeout = 600000;
            smtpClient.Send(msg);
            m = "email sent";
        }
        catch (Exception ex)
        {
            m = "Error : " + ex.Message;
            SendEmail("Error ", "jagvirjb@gmail.com", "Error", ex.Message);
        }
        return m;
    }

    public string getAccessRights(string pageName)
    {  
        SortedList list = new SortedList();
        list.Add("@EmpId", HttpContext.Current.Request.Cookies["EmpId"].Value );
        list.Add("@ActivityName", pageName);
        list.Add("@RoleId", HttpContext.Current.Request.Cookies["RoleId"].Value);
        DataTable dt =comfun.fillDataTable("UserPermissionCheckByEmpId", "", list);

        string x = "";
        list.Clear();
        for (int i = 0; i < dt.Rows.Count; i++)
        {
            //list.Add(dt.Rows[i]["AccessRight"].ToString(), dt.Rows[i]["HasPermission"].ToString());
            x += dt.Rows[i]["AccessRight"].ToString() + ":" + dt.Rows[i]["HasPermission"].ToString() + ",";
        }
        return x.TrimEnd(',');
    }
    public string SendEmailWithAttach1(string displayName, string toEmail, string subject, string message,string attachPath, MailMessage msg, string ReportingManager, string HrManager, string Hr)
    {
        string m = "";
        try
        {
            if (siteUrl.Contains("localhost"))
            {
               //toEmail = "jagvirjb@gmail.com";
          toEmail = "munish.kumar@horizontelecom.in";
            }
            //string email = "jagvir.singh@horizontelecom.in";
            //string password = "Htis@@123";

            string email = "Wfms@horizontelecom.in";
            string password = "wfms@9002";
            displayName = "Htis Telecom-HR";
            var loginInfo = new NetworkCredential(email, password);
            //var msg = new MailMessage();
            //var smtpClient = new SmtpClient("smtp.rediffmailpro.com");
            var smtpClient = new SmtpClient("smtp.rediffmailpro.com");
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtpClient.Port = 587;
            smtpClient.Credentials = new System.Net.NetworkCredential(email, password);
            msg.From = new MailAddress("hrms@wfms.in", displayName);
            msg.To.Add(new MailAddress(toEmail));

            msg.CC.Add(new MailAddress(Hr));
            msg.CC.Add(new MailAddress(HrManager));
            msg.CC.Add(new MailAddress(ReportingManager));
            msg.Bcc.Add(new MailAddress("anish.dhiman@horizontelecom.in"));
            //msg.To.Add(new MailAddress("jagvirjb@gmail.com"));
            //msg.Bcc.Add(new MailAddress("htis.comp111@gmail.com"));
            //msg.Bcc.Add(new MailAddress("htis.comp111@gmail.com"));
            HtmlString htmlString = new HtmlString(message);
            //       AlternateView htmlView =
            //AlternateView.CreateAlternateViewFromString(message, Encoding.UTF8, "text/html");
            //       msg.AlternateViews.Add(htmlView); // And a html attachment to make sure.

            msg.Subject = subject;
            msg.Body = message;
            msg.ReplyTo = new MailAddress("hrms@wfms.in");
            msg.Sender = new MailAddress("hrms@wfms.in", displayName);
            msg.IsBodyHtml = true;

            System.Net.Mail.Attachment attachment;

            attachment = new System.Net.Mail.Attachment((HttpContext.Current.Server.MapPath("~/") + attachPath));
            string Uplodefile = (HttpContext.Current.Server.MapPath("~/") + attachPath);
            string[] sentFiles = Directory.GetFiles((HttpContext.Current.Server.MapPath("~/") + attachPath));
            //string[] FileName = new string[] { attachPath };
            foreach (string File in sentFiles)
            {
                Attachment atch = new Attachment(File);
                msg.Attachments.Add(atch);
            }
            //  msg.Attachments.Add(attachment);

            smtpClient.UseDefaultCredentials = true;
                smtpClient.EnableSsl = true;

                smtpClient.Credentials = loginInfo;
                smtpClient.Credentials = loginInfo;
                smtpClient.Timeout = 600000;
                smtpClient.Send(msg);
         
            
         

             
            
            m = "email sent";
        }
        catch (Exception ex)
         {
            m = "Error : " + ex.Message;
            SendEmail("Error ", "jagvirjb@gmail.com", "Error", ex.Message);
        }
        return m;
    }

}



