using System;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Web;

using System.Collections;
using System.IO;
using System.Text;
//using Telerik.WebControls;
using System.Net;
using System.Net.Mail;
using System.Linq;
using System.Collections.Generic;
using cls_Encrypt;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Xml.Linq;

public class emails
{

    payrollFunctions comfun = new payrollFunctions();
   
    public string mail(SortedList list, string templateName, string toEmail, string subject, string displayName)
    {
        
        string mes = "";
        try
        {
            if (!templateName.Contains(".htm"))
            {
                templateName = templateName + ".htm";
            }
            StreamReader reader = new StreamReader(HttpContext.Current.Server.MapPath("~/EmailTemplates/" + templateName));
            mes = reader.ReadToEnd();
            string hash = string.Empty;
            foreach (string li in list.Keys)
            {
                if (!li.Trim().StartsWith("#") && !li.Trim().EndsWith("#"))
                {
                    hash = "#" + li.Trim()+"#";
                }
                else
                {
                    hash = li.Trim();
                }
                mes = mes.Replace(hash, list[li].ToString().Trim());
            }
            mes = comfun.SendEmail(displayName, toEmail, subject, mes).ToString();

        }
        catch (Exception ex)
        {
            mes = "Error: Please try later";
        }
        return mes;
    }
}
