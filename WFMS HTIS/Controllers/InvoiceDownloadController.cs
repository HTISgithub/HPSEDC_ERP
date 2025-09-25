using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Xml.Linq;
using static Payroll.portal.Controllers.AdminController;
using System.Globalization;
using ClosedXML.Excel;
using System.Web.Security;
using System.Threading;
using Newtonsoft.Json;
using System.Configuration;
using System.Data.OleDb;
using System.Net.Http;
using System.Web.Helpers;
using System.Net;
using System.Net.Mail;
using QRCoder;
using System.Drawing;
using System.Drawing.Imaging;
using System.Net.Mime;
using System.Text;
using System.Web.UI;
using iTextSharp.text;
using iTextSharp.text.html.simpleparser;
//using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml.pipeline.html;
using iTextSharp.tool.xml.parser;
using iTextSharp.tool.xml.pipeline.css;
using iTextSharp.tool.xml.pipeline.end;
using iTextSharp.tool.xml.html;
using ExcelDataReader;
using System.Threading.Tasks;
//using MasterIndia;
//using MasterIndia.Model;
using System.Web.Script.Serialization;
//using Payroll.portal.SessionLogout;
using Payroll.portal.Models;

namespace Payroll.portal.Controllers
{
    //[SessionExpire]
    public class InvoiceDownloadController  : Controller
    {

        commonFunctions comfun = new commonFunctions();
        payrollFunctions payfun = new payrollFunctions();
        //MasterIndiaFunction masterIndia = new MasterIndiaFunction();

        // GET: Deployment
        public ActionResult HpsedcInvoice1()
        {
            return View();
        }
        [HttpPost]
        public JsonResult HpsedcInvoieDetails1()

        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@DeptBillId", "0");
            list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
            dt = comfun.fillDataTable("TallyHpsedcDepartmentInvoice", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json, JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public JsonResult HpsedcInvoieDetails2()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@DeptBillId", "0");
            list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
            dt = comfun.fillDataTable("TallyHpsedcDepartmentInvoice", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json, JsonRequestBehavior.AllowGet);
        }
    }
}