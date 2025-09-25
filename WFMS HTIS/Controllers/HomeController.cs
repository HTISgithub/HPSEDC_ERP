using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Validation;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using System.Configuration;
//using LinqToExcel;
using System.Data.SqlClient;
using System.Collections;
using ClosedXML.Excel;
using Excelcon = Microsoft.Office.Interop.Excel;   //namespace
using Newtonsoft.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Net.Http;
using Payroll.portal.Models;
using System.Threading;
using System.Device.Location;
using System.Web.Script.Serialization;
using Newtonsoft.Json.Linq;
/*using ExcelDataReader;*/

namespace Payroll.portal.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        #region Project Management System
        public ActionResult ProjectAssignedToEmployee()
        {
            DataSet ds = new DataSet();
            ds = comfun.fillDataSet("stpEmployeeProjectddl", "", null);

            return View(ds);
        }
        public ActionResult _SubProject_ddl()
        {
            SortedList list = new SortedList();
            DataTable dt = new DataTable();
            list.Add("@Id", Request.Form["ProjectId"].ToString());
            dt = comfun.fillDataTable("stpSubprojectdll", "", list);
            return PartialView("_SubProject_dll", dt);
        }

        public ActionResult _Project_Employee_mapped_list()
        {
            //  comfun.saveformname("_emp_mapped_list", "/Home/_emp_mapped_list", "HR- Pm Employee List", "", "N");
            SortedList list = new SortedList();
            list.Add("@ProjectId", Request.Form["ProjectId"].ToString());
            DataTable dt = comfun.fillDataTable("ApiProjectEmployeeMappedList", "", list);
            return PartialView(dt);
        }
        public JsonResult _addProjectEmployeeToMapping()
        {
            // comfun.saveformname("_addToMapping", "/Home/_addToMapping", "HR- Pm Employee Mapping Add", "", "N");
            string mes = "";
            try
            {

                SortedList list = new SortedList();
                list.Add("@fiProjectID", Request.Form["ProjectId"].ToString());
                list.Add("@fiSubProjectId", Request.Form["SubProjectId"].ToString());
                list.Add("@EmpId", Request.Form["EmpID"].ToString());
                list.Add("@Date", Request.Form["Date"].ToString());
                list.Add("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                mes = comfun.executeNonQueryWMessage("ApiProject_EmpAddToMapping", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = "Error:" + ex.Message;
            }
            return Json(mes);
        }
        #endregion Properties
        #region Payroll Management System
        string siteUrl = System.Configuration.ConfigurationManager.AppSettings["SiteUrl"];
        public const string _roles = payrollFunctions._rolesGlobal;
        commonFunctions comfun = new commonFunctions();
        payrollFunctions payfun = new payrollFunctions();
        SqlConnection conLogin = new SqlConnection(ConfigurationManager.ConnectionStrings["cnLogin"].ConnectionString);

        public ActionResult Index()
        {
            string mView = "index";
            DataTable dt = new DataTable();
            SortedList list = new SortedList();

            DataSet ds = new DataSet();

            if (payfun.sessionRecreate() == "expires")
            {
                //return PartialView("_sessionExpired");
                return RedirectToAction("NewLogin", "Account");
            }
            if (User.Identity.Name == "")
            {
                //return PartialView("_sessionExpired");

            }

            Popup();

            if (Session["RoleId"].ToString() == "2" || Session["RoleId"].ToString() == "7" || Session["RoleId"].ToString() == "24")
            {
                ViewBag.Project = "ams";
                Session["Project"] = "ams";
                Session["ProjectName"] = "Attendance Management System";
            }

            else
            {
                ViewBag.Project = "esc";
                Session["Project"] = "esc";
                if (Session["RoleId"].ToString() == "3")
                {
                    mView = "IndexEsc";
                }
                else
                {
                    SortedList list1 = new SortedList();

                    list1.Add("@LoginId", Session["EmpId"].ToString());
                    ds = comfun.fillDataSet("stpPendingApproval", "", list1);
                    mView = "ApprovalDashBoard";
                }

                Session["ProjectName"] = "Employee Self Care";
            }

            if (Request.QueryString["id"] != null)
            {
                ViewBag.Project = comfun.decryptString(Request.QueryString["id"].ToString().Trim()).Trim();
                Session["Project"] = comfun.decryptString(Request.QueryString["id"].ToString().Trim()).Trim();
                Response.Cookies["Project"].Value = comfun.decryptString(Request.QueryString["id"].ToString().Trim()).Trim();
                Response.Cookies["Project"].Expires = DateTime.Now.AddDays(1);

                Response.Cookies["ProjectName"].Value = "Attendance Management System";
                Response.Cookies["ProjectName"].Expires = DateTime.Now.AddDays(1);

                if (Session["Project"].ToString().ToLower() == "ams")
                {
                    Session["ProjectName"] = "Payroll";
                    mView = "IndexAms";
                }
                if (Session["Project"].ToString().ToLower() == "payroll")
                {
                    Session["ProjectName"] = "Payroll";
                    mView = "IndexPayroll";
                }
                if (Session["Project"].ToString().ToLower() == "pms")
                {
                    Session["ProjectName"] = "Project Management System";
                    mView = "IndexPms";
                }
                //if (Session["Project"].ToString().ToLower() == "esc")
                //{
                //    mView = "IndexEsc";
                //    Session["ProjectName"] = "Employee Self Care";
                //}
                if (Session["Project"].ToString().ToLower() == "esc")
                {

                    if (Session["RoleId"].ToString() == "3")
                    {
                        mView = "IndexEsc";
                    }
                    else
                    {
                        SortedList list2 = new SortedList();
                        list2.Add("@LoginId", Session["EmpId"].ToString());
                        ds = comfun.fillDataSet("stpPendingApproval", "", list2);
                        mView = "ApprovalDashBoard";

                    }

                    Session["ProjectName"] = "Employee Self Care";
                }

                if (Session["Project"].ToString().ToLower() == "account")
                {
                    if (Session["RoleId"].ToString() == "14")
                    {
                        mView = "IndexSalePerson";
                    }
                    else if (Session["RoleId"].ToString() == "13")
                    {
                        mView = "IndexSaleHead";
                    }
                    else if (Session["RoleId"].ToString() == "15")
                    {
                        mView = "IndexSCM";
                    }
                    else
                    {
                        mView = "IndexAcMgt";
                    }
                    Session["ProjectName"] = "Financial Management";

                }

                if (Session["Project"].ToString().ToLower() == "invetory")
                {
                    mView = "IndexInvMgt";
                    Session["ProjectName"] = "Inventory Management";
                }
                if (Session["Project"].ToString().ToLower() == "compmgt")
                {
                    mView = "IndexCompMgt";
                    Session["ProjectName"] = "Complaint Management";
                }
                if (Session["Project"].ToString().ToLower() == "hiring")
                {
                    mView = "IndexHiriMgt";
                    Session["ProjectName"] = "Hiring Management";
                }
            }
            if (Session["Project"].ToString().ToLower() == "ams" || Session["Project"].ToString().ToLower() == "")
            {
                Session["ProjectName"] = "Attendance Management System";

                list.Add("@AsonDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy"));
                ds = comfun.fillDataSet("stpDashBoard_Ams", "", list);


                /******************Pie chart************************/
                string chart = "";
                if (ds.Tables.Count > 0)
                {
                    dt = ds.Tables[0];
                    chart = "labels: ['Present', 'Absent', 'Leave' ],";
                    chart += "datasets: [{";
                    chart += "label: '# of Votes',";
                    chart += "data:[";
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        chart += dt.Rows[i]["Present"].ToString() + ",";
                        chart += dt.Rows[i]["Absents"].ToString() + ",";
                        chart += dt.Rows[i]["Leave"].ToString();
                    }
                    chart += "],";
                    chart += @"
                        backgroundColor: ['#ffbd0c','#994499','#17b0fb'
                        ],
                        }]";
                }
                HtmlString htmlString = new HtmlString(chart);
                ViewBag.Pie = htmlString;

                /******************Line chart Trend************************/

                if (ds.Tables.Count > 0)
                {
                    dt = ds.Tables[2];
                    chart = "{ data:[";
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        chart += dt.Rows[i]["PP"].ToString() + ",";
                    }
                    chart += "],";
                    chart += @"
                        label: 'Present',
                        borderColor: '#ffbd0c',
                        fill: false
                    },
                    ";

                    chart += "{ data:[";
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        chart += dt.Rows[i]["AA"].ToString() + ",";
                    }
                    chart += "],";
                    chart += @"
                        label: 'Absent',
                        borderColor: '#994499',
                        fill: false
                    } 
                    ";

                    string attDates = "";
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        attDates += "'" + Convert.ToDateTime(dt.Rows[i]["AttendanceDate"]).ToString("dd-MMM-yyyy") + "',";
                    }
                    htmlString = new HtmlString(attDates);
                    ViewBag.TrendLabels = htmlString;
                }
                htmlString = new HtmlString(chart);
                ViewBag.Trend = htmlString;

                if (ds.Tables.Count > 0)
                {
                    dt = ds.Tables[1];

                    chart = @" 
                        labels: ['Indoor', 'Outdoor'],
                        datasets:[
                            {
                                label:'',
                                backgroundColor:['#ffbd0c','#8e5ea2' ],
                                data:[";
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {

                        chart += dt.Rows[i]["Indoor"].ToString() + ",";
                        chart += dt.Rows[i]["OutDoor"].ToString() + ",";
                        chart += "0";
                    }
                    chart += "]";
                    chart += "}]";
                }
                htmlString = new HtmlString(chart);
                ViewBag.PresentDetail = htmlString;


                list.Clear();
                list.Add("@AsonDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy"));
                dt = comfun.fillDataTable("Stpuserstatus", "", list);

                chart = @" 
                        labels: ['Total', 'Active','LoginUser','Indoor','Mobile','Web'],
                        datasets:[
                            {
                                label:'',
                                backgroundColor:['#ffbd0c','#8e5ea2','#17b0fb','#6ca329','#ff5400','#00476a' ],
                                data:[";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    chart += dt.Rows[i]["Total"].ToString() + ",";
                    chart += dt.Rows[i]["Active"].ToString() + ",";
                    chart += dt.Rows[i]["LoginUser"].ToString() + ",";
                    chart += dt.Rows[i]["Indoor"].ToString() + ",";
                    chart += dt.Rows[i]["Mobile"].ToString() + ",";
                    chart += dt.Rows[i]["Web"].ToString();
                }
                chart += "]";
                chart += "}]";

                htmlString = new HtmlString(chart);
                ViewBag.UserStatus = htmlString;
            }

            if (Session["Project"].ToString().ToLower() == "payroll")
            {
                Session["ProjectName"] = "Payroll Management System";

                list.Add("@AsonDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy"));
                ds = comfun.fillDataSet("stpDashboard_Payroll", "", null);
                HtmlString htmlString = new HtmlString("");
                string chart = "";
                if (ds.Tables.Count > 0)
                {
                    dt = ds.Tables[1];

                    chart =

                    chart += "{ data:[";
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        chart += dt.Rows[i]["Amount"].ToString() + ",";
                    }
                    chart += "],";
                    chart += @"
                        label: 'Amount',
                        borderColor: '#994499',
                        fill: false
                    } 
                    ";

                    string attDates = "";
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        attDates += "'" + dt.Rows[i]["MonthYear"].ToString() + "',";
                    }
                    htmlString = new HtmlString(attDates);
                    ViewBag.TrendLabels = htmlString;
                }
                htmlString = new HtmlString(chart);
                ViewBag.Trend = htmlString;
            }

            if (Session["Project"].ToString().ToLower() == "account")
            {
                Session["ProjectName"] = "Finance Management System [" + Session["SessionName"].ToString() + "]";

                list.Add("@AsonDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy"));
                ds = comfun.fillDataSet("stpDashboard_Payroll", "", null);
                HtmlString htmlString = new HtmlString("");
                string chart = "";
                if (ds.Tables.Count > 0)
                {
                    dt = ds.Tables[1];

                    chart =

                    chart += "{ data:[";
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        chart += dt.Rows[i]["Amount"].ToString() + ",";
                    }
                    chart += "],";
                    chart += @"
                        label: 'Amount',
                        borderColor: '#994499',
                        fill: false
                    } 
                    ";

                    string attDates = "";
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        attDates += "'" + dt.Rows[i]["MonthYear"].ToString() + "',";
                    }
                    htmlString = new HtmlString(attDates);
                    ViewBag.TrendLabels = htmlString;
                }
                htmlString = new HtmlString(chart);
                ViewBag.Trend = htmlString;
            }

            return View(mView, ds);
        }

        public ActionResult payrollGraphData()
        {
            DataSet ds = comfun.fillDataSet("stpDashboard_Payroll", null, null);
            DataTable dt = ds.Tables[0];
            DataTable dt2 = ds.Tables[1];

            var json1 = JsonConvert.SerializeObject(ds, new JsonSerializerSettings
            {
                ReferenceLoopHandling = ReferenceLoopHandling.Ignore
            });

            return Json(json1, JsonRequestBehavior.AllowGet);
        }

        public ActionResult Unauthorized()
        {
            return PartialView("NotAuthorized");
        }

        public ActionResult _homePagePanels()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            SortedList list = new SortedList();
            string proc = "";
            if (Request.Form["Method"].ToString() == "Birthday")
            {
                list.Clear();
                list.Add("@Month", DateTime.Now.Month);
                list.Add("@Day", DateTime.Now.Day);
                list.Add("@LocationId", Session["LocationId"].ToString());
                list.Add("@Employeecategory", "");
                proc = "HomePage_birthdayList";
            }
            if (Request.Form["Method"].ToString() == "ServiceAn")
            {
                list.Clear();
                list.Add("@Month", DateTime.Now.Month);
                list.Add("@Day", DateTime.Now.Day);
                list.Add("@LocationId", Session["LocationId"].ToString());
                proc = "HomePage_ServiceAnList";
            }
            if (Request.Form["Method"].ToString() == "WeddingAn")
            {
                list.Clear();
                list.Add("@Month", DateTime.Now.Month);
                list.Add("@Day", DateTime.Now.Day);
                proc = "HomePage_WeddingAnList";
            }
            if (Request.Form["Method"].ToString() == "WelcomeNew")
            {
                list.Clear();
                list.Add("@Date1", DateTime.Now.ToString("dd-MMM-yyyy"));
                list.Add("@Date2", DateTime.Now.AddDays(-7).ToString("dd-MMM-yyyy"));

                proc = "HomePage_WelcomeNew";
            }
            if (Request.Form["Method"].ToString() == "News")
            {
                proc = "HomePage_News";
            }
            list.Add("@CompanyId", Session["CompanyId"].ToString());
            DataTable dt = comfun.fillDataTable(proc, "", list);
            //if (Request.Form["Method"].ToString() == "Birthday")
            //{

            //    if (Convert.ToDateTime(dt.Rows[0]["Date"]).ToString("dd-MMM") == DateTime.Now.ToString("dd-MMM"))
            //    {
            //        string mes = string.Empty;
            //        string EmpId = dt.Rows[0]["EmpId"].ToString();
            //        string EmailId = dt.Rows[0]["Email"].ToString();
            //        string EmpName = dt.Rows[0]["EmpName"].ToString();
            //        string EmpCode = dt.Rows[0]["EmpCode"].ToString();
            //        //string Date = dt.Rows[0]["Date"].ToString();
            //        //string Remarks = dt.Rows[0]["BirthdayWishes"].ToString();
            //        StreamReader reader1 = new StreamReader(Server.MapPath("~/EmailTemplates/EmployeeBirthday.htm"));
            //        mes = reader1.ReadToEnd();
            //        mes = mes.Replace("#EmpName#", EmpName);
            //        mes = mes.Replace("#EmpCode#", EmpCode);
            //        //mes = mes.Replace("#Date#", Date);
            //        //mes = mes.Replace("#Remarks#", Remarks);

            //        payfun.SendEmail("Birthday Wishes", EmailId, "Birthday Wishes -" + EmpCode, mes);
            //    }
            //}
            return PartialView(dt);
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";
            return View();
        }
        public ActionResult Popup()
        {
            SortedList list = new SortedList();
            string mes = string.Empty;
            list.Add("@DomainId", Session["DomainId"]);
            mes = comfun.executeNonQueryWMessage("stpDomainExpire", null, list, conLogin).ToString();
            return Json(mes);
        }
        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";
            return View();
        }

        public ActionResult _ComplaintCount()
        {
            SortedList list = new SortedList();
            list.Add("@RoleId", Session["RoleId"].ToString());
            list.Add("@SessionId", Session["EmpId"].ToString());
            DataTable ds = new DataTable();
            ds = comfun.fillDataTable("stpTotalComplaint", null, list);
            var json1 = JsonConvert.SerializeObject(ds);
            return Json(json1);
        }
        public ActionResult _SaleCount()
        {
            DataSet ds = new DataSet();
            SortedList list = new SortedList();
            list.Add("@SessionId", Session["SessionId"]);
            list.Add("@LoginID", Session["EmpId"]);
            ds = comfun.fillDataSet("stpSaleExecutive_Dashboard", null, list);
            var json1 = JsonConvert.SerializeObject(ds);
            return Json(json1);
        }
        public ActionResult _SaleHeadCount()
        {
            DataSet ds = new DataSet();
            SortedList list = new SortedList();
            list.Add("@SessionId", Session["SessionId"]);
            list.Add("@LoginID", Session["EmpId"]);
            ds = comfun.fillDataSet("stpSaleHead_Dashboard", null, list);
            var json1 = JsonConvert.SerializeObject(ds);
            return Json(json1);
        }
        public ActionResult _SaleSCMCount()
        {
            DataSet ds = new DataSet();
            SortedList list = new SortedList();
            list.Add("@SessionId", Session["SessionId"]);
            list.Add("@LoginID", Session["EmpId"]);
            ds = comfun.fillDataSet("stpSCMHead_Dashboard", null, list);
            var json1 = JsonConvert.SerializeObject(ds);
            return Json(json1);
        }

        /*
        // get: user  
        public ActionResult excelimport()
        {
            comfun.saveformname("excelimport", "/Home/excelimport", "excel import", "excelimport", "Y", "excelimport", "excelimport", "List");
            return View();
        }

        //Step 3
        public ActionResult _excelSheetReader()
        {
            //Excel_ImportSetting
            SortedList list = new SortedList();
            list.Add("@ExcelTable", Request.Form["SheetName"].ToString());
            DataTable dt = comfun.fillDataTable("Excel_ImportSetting_select", "", list);

            DataTable dtExcel = new DataTable();
            string columns = string.Empty;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i]["ExcelCol"].ToString().Trim().Length > 0)
                {
                    if (dt.Rows[i]["ExcelCol"].ToString().Trim().Contains("["))
                    {
                        columns += dt.Rows[i]["ExcelCol"].ToString().Trim() + ",";
                    }
                    else
                    {
                        columns += "[" + dt.Rows[i]["ExcelCol"].ToString().Trim() + "], ";
                    }

                }

                //dtExcel.Columns.Add(dt.Rows[i]["ExcelCol"].ToString(), typeof(System.String));  
            }

            if (columns.Length > 2)
            {
                columns = columns.Trim().Remove(columns.Trim().Length - 1, 1);

                dtExcel = _excelReadSheetByName(Request.Form["ExcelName"].ToString(), Request.Form["SheetName"].ToString(), columns);
            }
            return PartialView("_excelSheet", dtExcel);

            //DataTable dtExcelTables = dt.DefaultView.ToTable(true, "ExcelTable");
            //DataTable dtDBtables = dt.DefaultView.ToTable(true, "DbTable");
        }

        public string columnBuilder(string excelSheetName, string colType)
        {
            SortedList list = new SortedList();
            list.Add("@ExcelTable", excelSheetName);
            DataTable dt = comfun.fillDataTable("Excel_ImportSetting_select", "", list);

            DataTable dtExcel = new DataTable();
            string columns = string.Empty;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (colType == "Excel")
                {
                    if (dt.Rows[i]["ExcelCol"].ToString().Trim().Length > 0)
                    {
                        if (!dt.Rows[i]["ExcelCol"].ToString().Trim().Contains("["))
                        {
                            columns += "[" + dt.Rows[i]["ExcelCol"].ToString().Trim() + "], ";
                        }
                        else
                        {
                            columns += dt.Rows[i]["ExcelCol"].ToString().Trim() + ", ";
                        }
                    }
                }
                if (colType == "DbCol")
                {
                    if (dt.Rows[i]["DbCol"].ToString().Trim().Length > 0)
                    {
                        if (!dt.Rows[i]["DbCol"].ToString().Trim().Contains("["))
                        {
                            columns += "[" + dt.Rows[i]["DbCol"].ToString().Trim() + "], ";
                        }
                        else
                        {
                            columns += dt.Rows[i]["DbCol"].ToString().Trim() + ", ";
                        }
                    }
                }
            }

            if (columns.Length > 2)
            {
                columns = columns.Trim().Remove(columns.Trim().Length - 1, 1);
            }
            return columns;
        }


        public DataTable _excelReadSheetByName(string excelName, string sheetName, string columns)
        {
            DataTable dt = new DataTable();
            string filepath = string.Empty;

            string path = Path.Combine(Server.MapPath("~/uploads/"));

            filepath = path + Path.GetFileName(excelName);
            string extension = Path.GetExtension(excelName);
            string constring = string.Empty;
            DataTable dtexcel = new DataTable();
            try
            {

                Excelcon.Application xlApp = new Excelcon.Application();
                Excelcon.Workbook xlWorkbook = xlApp.Workbooks.Open(filepath, 0,
                                    true,
                                    5,
                                    "",
                                    "",
                                    true,
                                    Excelcon.XlPlatform.xlWindows,
                                    "\t",
                                    false,
                                    false,
                                    0,
                                    true,
                                    1,
                                    0);
                // Excelcon.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excelcon.Worksheet xlWorksheet = (Excelcon.Worksheet)xlWorkbook.Worksheets[sheetName];
                Excelcon.Range xlRange = xlWorksheet.UsedRange;


                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;


                columns = "";

                for (int i = 1; i <= 1; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        //write the value to the Grid   
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            if (!xlRange.Cells[i, j].Value2.ToString().ToLower().Contains("."))
                            {
                                if (!xlRange.Cells[i, j].Value2.ToString().Contains("["))
                                {
                                    columns += "[" + xlRange.Cells[i, j].Value2.ToString() + "], ";
                                    dtexcel.Columns.Add(xlRange.Cells[i, j].Value2);
                                }
                                else
                                {
                                    columns += xlRange.Cells[i, j].Value2.ToString() + ", ";
                                    dtexcel.Columns.Add(xlRange.Cells[i, j].Value2);
                                }
                            }
                        }
                    }
                }
                if (columns.Trim().EndsWith(","))
                {
                    columns = columns.Trim().Remove(columns.Trim().Length - 1, 1);
                }
                DateTime result = new DateTime();
                DataRow dr;
                string str = "";
                string coln = "";
                for (int i = 2; i <= rowCount; i++)
                {
                    dr = dtexcel.NewRow();
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (dtexcel.Columns.Count >= j)
                        {
                            if (xlRange.Cells[i, 1].Value != null)
                            {
                                str = (xlRange.Cells[i, j].Value == null ? "" : xlRange.Cells[i, j].Value.ToString());
                                if (sheetName == "EmpMast")
                                {
                                    coln = xlRange.Cells[1, j].Value.ToString();
                                    if (coln == "DateOfBirth" || coln == "DateOfComfirmation")
                                    {
                                        dr[j - 1] = dateformat(str);
                                    }
                                    else
                                    {
                                        dr[j - 1] = str;
                                    }
                                }
                                else if (sheetName == "HrdTran")
                                {
                                    coln = xlRange.Cells[1, j].Value.ToString();
                                    if (coln == "Trn_WEF" || coln == "Trn_Date" || coln == "DOL" || coln == "DOS" || coln == "DOR" || coln == "DOC")
                                    {
                                        dr[j - 1] = dateformat(str);
                                    }
                                    else
                                    {
                                        dr[j - 1] = str;
                                    }
                                }
                                else
                                {
                                    dr[j - 1] = str;
                                }
                            }
                        }
                    }
                    dtexcel.Rows.Add(dr);
                }


                if (dtexcel.Rows.Count > 0)
                {
                    for (int i = dtexcel.Rows.Count - 1; i >= 0; i--)
                    {
                        if (dtexcel.Rows[i][0] == DBNull.Value)
                        {
                            dtexcel.Rows[i].Delete();
                        }
                    }
                    dtexcel.AcceptChanges();
                    dt = dtexcel;
                }
                else
                {
                    dt = dtexcel;
                }
                //IntPtr xAsIntPtr = new IntPtr(xlWorkbook.Application.Hwnd);
                killExcel(xlWorkbook);
                xlApp.Quit();

            }
            catch (Exception ex)
            {
                comfun.errorMessage("_excelReadSheetByName 123", ex.Message);
            }
            return dt;
        }


        private void killExcel(Excelcon.Workbook xl)
        {
            GetWindowThreadProcessId((IntPtr)xl.Application.Hwnd, out iProcessId);
            var process = System.Diagnostics.Process.GetProcessesByName("Excel");
            foreach (var p in process)
            {
                comfun.errorMessage("KillExcel", p.ProcessName);
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    if (p.Id == iProcessId)
                    {
                        try
                        {
                            p.Kill();
                        }
                        catch (Exception ex)
                        {
                            comfun.errorMessage("KillExcel", ex.Message);
                        }
                    }
                }
            }
        }

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        
        uint iProcessId = 0;

        public DataSet dsExcels = new DataSet();
        string excelCols = "";
        string idCols = "";
        string mcodCols = "";
        */

        /*
        public DataTable _excelSaveToDBPayTran(DataTable dtExcelData, string sheetName, string excelName)
        {
            dtExcelData.Columns.Add("~Id", typeof(System.Int32));
            dtExcelData.Columns.Add("~Saved", typeof(System.Int32));
            dtExcelData.Columns.Add("~Updated", typeof(System.Int32));
            dtExcelData.Columns.Add("~AlreadyExist", typeof(System.Int32));
            dtExcelData.Columns.Add("~Error", typeof(System.Int32));
            dtExcelData.Columns.Add("~Remarks", typeof(System.String));
            dtExcelData.Columns.Add("SheetName", typeof(System.String));

            SortedList list = new SortedList();
            list.Add("@ExcelTable", "PayTran");
            DataTable dtExcelSettings = comfun.fillDataTable("Excel_ImportSetting_select", "", list);

            string[] columnsExcels = columnBuilder(sheetName, "Excel").Split(',');
            string[] columnsDB = columnBuilder(sheetName, "DbCol").Split(',');


            string[] strCol = { };
            DataTable tempUpdate = new DataTable();

            DataTable dtChargeIds = comfun.fillDataTable("", "SELECT '' as Emp_Code,fvCharges,[fiCharges] PayChargeId,[fvImportCode] Col,'' as [fnChargeAmount],'' as Trn_WEF FROM[tblVIKSATpayCharges] where fvImportCode is not null", null);


            DataTable dtChargeHeadCols = new DataTable();

            for (int i = 0; i < dtChargeIds.Rows.Count; i++)
            {
                dtChargeHeadCols.Columns.Add(dtChargeIds.Rows[i]["Col"].ToString(), typeof(string));
            }

            //DataRow drChargesVals;
            //drChargesVals = dtChargeHeadCols.NewRow();
            //for (int i = 0; i < dtChargeIds.Rows.Count; i++)
            //{
            //    foreach (DataColumn column in dtChargeHeadCols.Columns)
            //    {
            //        if (column.ColumnName == dtChargeIds.Rows[i]["Col"].ToString())
            //        {
            //            drChargesVals[column.ColumnName] = dtChargeIds.Rows[i]["Val"].ToString();
            //        }
            //    }
            //}

            DataTable excelDataFinal = new DataView(dtChargeIds, "1=2", null, DataViewRowState.CurrentRows).ToTable();

            for (int j = 0; j < dtExcelData.Rows.Count; j++)
            {
                for (int i = 0; i < dtChargeIds.Rows.Count; i++)
                {
                    foreach (DataColumn columnExcel in dtExcelData.Columns)
                    {
                        foreach (DataColumn columnPayCharges in dtChargeHeadCols.Columns)
                        {
                            if (columnExcel.ColumnName == columnPayCharges.ColumnName)
                            {
                                if (dtChargeIds.Rows[i]["Col"].ToString() == columnExcel.ColumnName.ToString())
                                {
                                    dtChargeIds.Rows[i]["fnChargeAmount"] = dtExcelData.Rows[j][columnExcel.ColumnName];
                                    dtChargeIds.Rows[i]["Emp_Code"] = dtExcelData.Rows[j]["Emp_Code"];
                                    dtChargeIds.Rows[i]["Trn_WEF"] = dtExcelData.Rows[j]["Trn_WEF"];
                                }
                            }
                        }
                    }
                }
                excelDataFinal.Merge(dtChargeIds);
            }

            //dtChargeHeadCols.Rows.Add(drChargesVals);
            dtExcelData = _PayChargesSubmit(dtExcelData, excelDataFinal);

            DataTable dtUploadedSheets = new DataTable();
            if (Session["UploadedExcel"] != null)
            {
                dtUploadedSheets = (DataTable)Session["UploadedExcel"];
            }
            string linkStr = "";
            HtmlString _htmlLinks = new HtmlString("");
            if (dtExcelData.Rows.Count > 0)
            {
                for (int i = 0; i < dtUploadedSheets.Rows.Count; i++)
                {
                    //UploadLinks 
                    if (dtUploadedSheets.Rows[i]["SheetName"].ToString() == sheetName)
                    {
                        DataTable dtTe = dtExcelData.AsEnumerable()
                                      .GroupBy(r => r["SheetName"])
                                      .Select(g =>
                                      {
                                          var row = dtExcelData.NewRow();

                                          row["SheetName"] = g.Key;
                                          row["~Saved"] = g.Sum(r => r.Field<int?>("~Saved") ?? 0);
                                          row["~Updated"] = g.Sum(r => r.Field<int?>("~Updated") ?? 0);
                                          row["~AlreadyExist"] = g.Sum(r => r.Field<int?>("~AlreadyExist") ?? 0);
                                          row["~Error"] = g.Sum(r => r.Field<int?>("~Error") ?? 0);


                                          return row;
                                      }).CopyToDataTable();
                        for (int j = 0; j < dtTe.Rows.Count; j++)
                        {
                            linkStr = "<a href='javascript:;' onclick='viewUploadedDetail(\"" + sheetName + "\",\"saved\")'><b>" + dtTe.Rows[j]["~Saved"].ToString() + "</b> saved</a><span>&nbsp;&nbsp;</span>";
                            linkStr += "<a href='javascript:;' onclick='viewUploadedDetail(\"" + sheetName + "\",\"updated\")'><b>" + dtTe.Rows[j]["~Updated"].ToString() + "</b> updated</a><span>&nbsp;</span>";
                            linkStr += "<a href='javascript:;' onclick='viewUploadedDetail(\"" + sheetName + "\",\"already\")'><b>" + dtTe.Rows[j]["~AlreadyExist"].ToString() + "</b> AlreadyExist</a><span>&nbsp;</span>";
                            linkStr += "<a href='javascript:;' onclick='viewUploadedDetail(\"" + sheetName + "\",\"error\")'><b>" + dtTe.Rows[j]["~Error"].ToString() + "</b> error</a><span>&nbsp;&nbsp;</span>";

                        }
                        _htmlLinks = new HtmlString(linkStr);

                        dtUploadedSheets.Rows[i]["UploadLinks"] = _htmlLinks;
                    }
                }
            }

            dsExcels.Tables.Add(dtExcelData);
            Session["dsExcels"] = dsExcels;
            return dtUploadedSheets;
        }
        */

        /*
        public DataTable _PayChargesSubmit(DataTable dtExcelData, DataTable dtloop)
        {
            for (int i = 0; i < dtExcelData.Rows.Count; i++)
            {
                string mes = string.Empty;

                connection conObj = new connection();
                SqlTransaction tran;
                if (conObj.con.State == ConnectionState.Closed)
                {
                    conObj.con.Open();
                }

                tran = conObj.con.BeginTransaction();
                SortedList list = new SortedList();
                try
                {

                    list.Add("@ChargeAppliedDate", dateformat(dtExcelData.Rows[i]["Trn_WEF"].ToString()).ToString());
                    list.Add("@SessionId", Session["SessionId"].ToString());
                    list.Add("@EmployeeCode", dtExcelData.Rows[i]["Emp_Code"].ToString());
                    list.Add("@GrossPay", 0);
                    mes = comfun.executeNonQueryWTranOutMes("payChargeAppliedPart1_ExcelSave", list, conObj.con, tran).ToString();
                    string partId = mes;

                    DataRow[] dr = dtloop.Select("Emp_Code= '" + dtExcelData.Rows[i]["Emp_Code"].ToString() + "'");

                    for (int j = 0; j < dr.Count(); j++)
                    {
                        if (dr[j]["fnChargeAmount"].ToString() != "")
                        {
                            list.Clear();
                            list.Add("@ChargeAppliedID", partId);
                            list.Add("@ChargePercentage", "0.00");
                            list.Add("@ChargeAmount", dr[j]["fnChargeAmount"].ToString());
                            list.Add("@SessionId", Session["SessionId"].ToString());
                            list.Add("@EmployeeCode", dr[j]["Emp_Code"].ToString());
                            list.Add("@ChargeAppliedDate", dateformat(dr[j]["Trn_WEF"].ToString()).ToString());
                            list.Add("@PayChargeId", dr[j]["PayChargeId"].ToString());
                            mes = comfun.executeNonQueryWTranOutMes("payChargeAppliedPart2_Excelsave", list, conObj.con, tran).ToString();
                        }
                    }

                    if (dtExcelData.Rows.Count > 0)
                    {
                        dtExcelData.Rows[i]["~Id"] = i + 1;
                        if (mes.ToLower().Contains("saved"))
                        {
                            dtExcelData.Rows[i]["~Saved"] = 1;
                            dtExcelData.Rows[i]["~Updated"] = 0;
                            dtExcelData.Rows[i]["~Error"] = 0;
                            dtExcelData.Rows[i]["~AlreadyExist"] = 0;
                        }

                        dtExcelData.Rows[i]["~Remarks"] = mes.ToLower();
                        dtExcelData.Rows[i]["SheetName"] = "PayTran";
                    }

                    mes = "Record saved successfully";
                    tran.Commit();
                    conObj.con.Close();
                }
                catch (Exception ex)
                {
                    tran.Rollback();
                    conObj.con.Close();
                    mes = comfun.errorMessage("Home  _PayChargesSubmit", ex.Message);
                }

            }
            return dtExcelData;
        }
        */


        /*
         //Excel Import New
         //[customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.rolePayrollTeam)]
         public ActionResult ExcelDataImport()
         {
             if (payfun.sessionRecreate() == "expires")
             {
                 return RedirectToAction("Login", "account");
             }
             if (Request.QueryString["key"] != null)
             {
                 Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
             }
             comfun.saveformname("ExcelDataImport", "/Home/ExcelDataImport", "Excel Import", "ExcelDataImport", "Y", "ExcelDataImport", "ExcelDataImport", "List");
             return View();
         }
         *//*
         //New Code
         //Step 2
         [HttpPost]
         public ActionResult _excelUpload()
         {
             string FileName = "";
             string mes = string.Empty;
             SortedList list = new SortedList();
             try
             {
                 int count = 0;

                 foreach (string file in Request.Files)
                 {
                     var fileContent = Request.Files[file];
                     if (fileContent != null && fileContent.ContentLength > 0)
                     {
                         string date = DateTime.UtcNow.AddMinutes(330).ToString("yyyyMMdd_HHmm");
                         FileName = User.Identity.Name + "_" + fileContent.FileName.Replace(".", "_") + "_" + date;
                         var fileName = FileName + Path.GetExtension(fileContent.FileName);
                         var path = Path.Combine(Server.MapPath("~/uploads"), fileName);
                         fileContent.SaveAs(path);
                         FileName = fileName;
                         count++;
                     }
                 }
                 list.Add("Msg", "File uploaded successfully");
                 list.Add("FileName", FileName);
                 Session["FileName"] = FileName.ToString();
             }
             catch (Exception ex)
             {
                 mes = "Error:" + ex.Message;
             }
             return Json(list, JsonRequestBehavior.AllowGet);
         }
         */
        /*
        public ActionResult _excelListSheetNames()
        {
            DataTable dt = new DataTable();
            try
            {
                string excel = Request.Form["FileName"].ToString();//"d29d5113-cb89-4e36-88cd-779c68f50c41.xls";


                Excelcon.Application xlApp = new Excelcon.Application();


                dt.Columns.Add("SheetName", typeof(System.String));
                dt.Columns.Add("UploadLinks", typeof(System.String));

                DataRow dr;
                string file = Path.Combine(Server.MapPath("~/uploads/") + excel);
                comfun.errorMessage("0", file);
                Excelcon.Workbook excelBook = xlApp.Workbooks.Open(file, 0,
                                  true,
                                  5,
                                  "",
                                  "",
                                  true,
                                  Excelcon.XlPlatform.xlWindows,
                                  "\t",
                                  false,
                                  false,
                                  0,
                                  true,
                                  1,
                                  0);
                // comfun.errorMessage("1", "file opened");
                // ViewData["Error"] = "some Excel files opened on server";

                foreach (Excelcon.Worksheet wSheet in excelBook.Worksheets)
                {
                    // comfun.errorMessage("1", "Foreach");
                    if (wSheet.Visible != Excelcon.XlSheetVisibility.xlSheetHidden)
                    {
                        //comfun.errorMessage("1", wSheet.Name.ToString());
                        dr = dt.NewRow();
                        dr["SheetName"] = wSheet.Name.ToString();
                        dt.Rows.Add(dr);
                    }
                }
                killExcel(excelBook);

                //statement to get the worksheet object by using the sheet id  
                ViewData["Excel"] = excel;
                Session["UploadedExcel"] = dt;
            }
            catch (Exception ex)
            {
                comfun.errorMessage("Excel", ex.Message);
                ViewData["Error"] = "some Excel files opened on server | " + ex.Message;
            }
            return PartialView("_excelListSheetNames", dt);
        } 
        */

        public ActionResult Upload()
        {
            return View();
        }


        //[HttpPost]
        //[ValidateAntiForgeryToken]
        //public ActionResult Upload(HttpPostedFileBase upload)
        //{
        //    if (ModelState.IsValid)
        //    {

        //        if (upload != null && upload.ContentLength > 0)
        //        {
        //            // ExcelDataReader works with the binary Excel file, so it needs a FileStream
        //            // to get started. This is how we avoid dependencies on ACE or Interop:
        //            Stream stream = upload.InputStream;

        //            // We return the interface, so that
        //            IExcelDataReader reader = null;


        //            if (upload.FileName.EndsWith(".xls"))
        //            {
        //                reader = ExcelReaderFactory.CreateBinaryReader(stream);
        //            }
        //            else if (upload.FileName.EndsWith(".xlsx"))
        //            {
        //                reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
        //            }
        //            else
        //            {
        //                ModelState.AddModelError("File", "This file format is not supported");
        //                return View();
        //            }


        //            var conf = new ExcelDataSetConfiguration
        //            {
        //                ConfigureDataTable = _ => new ExcelDataTableConfiguration
        //                {
        //                    UseHeaderRow = true
        //                }
        //            };

        //            DataSet result = reader.AsDataSet(conf);
        //            reader.Close();

        //            return View(result.Tables[0]);
        //        }
        //        else
        //        {
        //            ModelState.AddModelError("File", "Please Upload Your file");
        //        }
        //    }
        //    return View();
        //}

        public JsonResult encryptString()
        {
            return Json(comfun.encryptString(Request.Form["str"].ToString()));
        }

        /*
        private string employee_import(DataTable dtExcelDataEmp, DataTable dtExcelSetting)
        {
            string Employeecode = string.Empty;
            string EmployeeName = string.Empty;

            string EmployeeCategoryCode = string.Empty;
            string BranchCode = string.Empty;
            string DepartmentCode = string.Empty;
            string DesignationCode = string.Empty;
            string Married = string.Empty;
            string GenderCode = string.Empty;
            string FatherHusbandName = string.Empty;
            string DateOfBirth = string.Empty;
            string Address1 = string.Empty;
            string Address2 = string.Empty;
            string Phone = string.Empty;
            string Mobile = string.Empty;
            string Email = string.Empty;
            string AadhaarNo = string.Empty;
            string ReportingManagerCode = string.Empty;
            string DateOfJoining = string.Empty;
            string CostCenterCode = string.Empty;
            string HRManagerCode = string.Empty;
            string LocationCode = string.Empty;
            string UnitCode = string.Empty;
            string CompanyCode = string.Empty;
            string StateName = string.Empty;
            string PFNo = string.Empty;
            string ESINo = string.Empty;
            string CityName = string.Empty;

            string ContractorCompanyCode = string.Empty;
            for (int i = 0; i <= dtExcelDataEmp.Rows.Count - 1; i++)
            {
                Employeecode = dtExcelDataEmp.Rows[i]["Emp_Code"].ToString();
                EmployeeName = dtExcelDataEmp.Rows[i]["Emp_Name"].ToString();
                EmployeeCategoryCode = dtExcelDataEmp.Rows[i]["Type_Code"].ToString();
                BranchCode = dtExcelDataEmp.Rows[i]["Cost_Code"].ToString();
                DepartmentCode = dtExcelDataEmp.Rows[i]["Dept_Code"].ToString();
                DesignationCode = dtExcelDataEmp.Rows[i]["Dsg_Code"].ToString();
                Married = dtExcelDataEmp.Rows[i]["MStatus"].ToString();
                GenderCode = dtExcelDataEmp.Rows[i]["Sex"].ToString();
                FatherHusbandName = dtExcelDataEmp.Rows[i]["FathHusbName"].ToString();
                DateOfBirth = dtExcelDataEmp.Rows[i]["DOB"].ToString();
                Address1 = dtExcelDataEmp.Rows[i]["MAddr1"].ToString();
                Address2 = dtExcelDataEmp.Rows[i]["MAddr2"].ToString();
                ReportingManagerCode = dtExcelDataEmp.Rows[i]["Mngr_Code"].ToString();
                DateOfJoining = dtExcelDataEmp.Rows[i]["DOJ"].ToString();
                HRManagerCode = dtExcelDataEmp.Rows[i]["Mngr1_code"].ToString();
                LocationCode = dtExcelDataEmp.Rows[i]["Loc_Code"].ToString();
                UnitCode = dtExcelDataEmp.Rows[i]["Cost_code"].ToString();
                CompanyCode = dtExcelDataEmp.Rows[i]["Divi_Code"].ToString();
                CityName = dtExcelDataEmp.Rows[i]["MCity"].ToString();
                ContractorCompanyCode = dtExcelDataEmp.Rows[i]["Comp_Code"].ToString();
                CostCenterCode = dtExcelDataEmp.Rows[i]["Cost_Code"].ToString();
                PFNo = dtExcelDataEmp.Rows[i]["PFNo"].ToString();
                ESINo = dtExcelDataEmp.Rows[i]["ESINo"].ToString();
                Mobile = dtExcelDataEmp.Rows[i]["MPHONENO"].ToString();

                if (Employeecode != "")
                {
                    if (payfun.IsExistsEmployee(Employeecode) == false)
                    {
                        string mes = string.Empty;
                        SortedList list = new SortedList();
                        try
                        {
                            list.Add("@CreatedBy", Session["EmpId"].ToString());
                            list.Add("@CreatedDate", comfun.dateISTstr());
                            list.Add("@EmployeeCode", Employeecode);
                            list.Add("@EmployeeName", EmployeeName);
                            list.Add("@BranchCode", BranchCode);
                            list.Add("@DepartmentCode", DepartmentCode);
                            list.Add("@DesignationCode", DesignationCode);
                            list.Add("@Married", Married);
                            list.Add("@GenderCode", GenderCode);
                            list.Add("@EmployeeCategoryCode", EmployeeCategoryCode);
                            list.Add("@FatherHusbandName", FatherHusbandName);
                            //   list.Add("@DateOfBirth", Convert.ToDateTime(DateOfBirth));
                            list.Add("@Address1", Address1);
                            list.Add("@Address2", Address2);
                            list.Add("@ReportingManagerCode", ReportingManagerCode);
                            // list.Add("@DateOfJoining", Convert.ToDateTime(DateOfJoining));
                            list.Add("@HRManagerCode", HRManagerCode);
                            list.Add("@LocationCode ", LocationCode);
                            list.Add("@UnitCode", UnitCode);
                            list.Add("@CompanyCode", CompanyCode);
                            list.Add("@CityName ", CityName);
                            list.Add("@CostCenterCode", CostCenterCode);
                            list.Add("@ContractorCompanyCode", ContractorCompanyCode);
                            list.Add("@PfNo", PFNo);
                            list.Add("@ESINo", ESINo);
                            list.Add("@MobileNo", Mobile);
                            mes = comfun.executeNonQueryWMessage("Employee_import", "", list).ToString();
                            if (!mes.Contains("Error"))
                            {
                                mes = comfun.encryptString(mes);
                            }
                        }
                        catch (Exception ex)
                        {
                            mes = comfun.errorMessage("Home excelimport", "Error:" + ex.Message);
                        }
                    }
                }
            }

            return "";
        }
        */

        public string dateformat(string date)
        {
            try
            {
                string d = Convert.ToDateTime(date).ToString("dd-MMM-yyyy");
                date = d;
            }
            catch (Exception ex)
            {

            }

            return date;
        }
        public ActionResult CompOffApplyShow()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            dt = comfun.fillDataTable("stpVIKSATPMEmployees_CompOffShow", "", list);
            return View(dt);
        }
        public ActionResult CompOffApply()
        {
            //DataTable dt = new DataTable();
            //SqlCommand cmd = new SqlCommand("stpVIKSATPMEmployees_Alllist", con);
            //cmd.CommandType = System.Data.CommandType.StoredProcedure;
            //SqlDataAdapter da = new SqlDataAdapter(cmd);
            //da.Fill(dt);
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            dt = comfun.fillDataTable("stpVIKSATPMEmployees_Alllist", "", list);
            return View(dt);
        }



        public ActionResult _CompOffApply()
        {
            //try
            //{
            //    string Message = string.Empty;
            //    SqlCommand cmd = new SqlCommand("stpEmployees_CompOffSave", con);
            //    cmd.Parameters.AddWithValue("@LoginId", User.Identity.Name);
            //    cmd.Parameters.AddWithValue("@AMDate", Request.Form["AMDate"].ToString());
            //    cmd.Parameters.AddWithValue("@fiEmployeeID", Request.Form["fiEmployeeID"].ToString());
            //    cmd.Parameters.AddWithValue("@fvRemarks", Request.Form["fvRemarks"].ToString());
            //    cmd.Parameters.Add("@Mes", SqlDbType.VarChar, 500);
            //    cmd.Parameters["@Mes"].Direction = ParameterDirection.Output;
            //    cmd.CommandType = System.Data.CommandType.StoredProcedure;
            //    SqlDataAdapter da = new SqlDataAdapter(cmd);
            //    con.Open();
            //    cmd.ExecuteNonQuery();
            //    Message = (string)cmd.Parameters["@Mes"].Value;
            //    con.Close();




            //    return Json(new
            //    {
            //        Status = "true",
            //        Data = Message,

            //    });
            //}
            //catch (Exception ex)
            //{
            //    return View();
            //}

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@LoginId", User.Identity.Name);
                list.Add("@AMDate", Request.Form["AMDate"].ToString());
                list.Add("@fiEmployeeID", Request.Form["fiEmployeeID"].ToString());
                list.Add("@fvRemarks", Request.Form["fvRemarks"].ToString());
                mes = comfun.executeNonQueryWMessage("stpEmployees_CompOffSave", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }

        public ActionResult _CompOffApproval()
        {
            string mes = string.Empty;
            SortedList list = new SortedList();
            list.Add("@loginCode", User.Identity.Name);
            list.Add("@fiEmployeeID", Request.Form["fiEmployeeID"].ToString());
            list.Add("@fdAMDate", Request.Form["fdAMDate"].ToString());
            list.Add("@fvRemarks", Request.Form["fvRemarks"].ToString().Trim());
            DataTable dt = comfun.fillDataTable("ApiEmpMonthlyLeaveByEmpid", "", list);
            return PartialView("_CompOffForApproval", dt);


        }
        public JsonResult _CompoffApprovalStatus()
        {
            String mes = string.Empty;
            SortedList list = new SortedList();
            list.Add("@DetailId", Request.Form["DetailId"].ToString());
            list.Add("@Id", Request.Form["Day"].ToString());
            list.Add("@Status", Request.Form["Status"].ToString().Trim());
            mes = comfun.executeNonQueryWMessage("CompoffApprovalByPm", "", list).ToString();
            return Json(mes);
        }
        public ActionResult _NoticeBoardList()
        {
            DataTable dt = comfun.fillDataTable("stpNoticeBoardDashboardList", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _NoticeBoardReadList()
        {
            SortedList list = new SortedList();

            list.Add("@EmpId", Session["EmpId"].ToString());
            DataTable dt = comfun.fillDataTable("stpNoticeBoardPendingReadList", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }


        public ActionResult _NoticeBoardById()
        {
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["NoticeBoard"].ToString());
            DataTable dt = comfun.fillDataTable("stpNoticeBoardById", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _NoticeBoardReadbyEmployeeSubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();

                list.Add("@fiNoticeBoardId", Request.Form["NoticeBoardId"].ToString());
                list.Add("@fiEmpId", Session["EmpId"].ToString());
                mes = comfun.executeNonQueryWMessage("stpNoticeBoardRead_Accept", "", list).ToString();


            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        #endregion Properties

        [Route("http://112.196.38.246:80/LAPI/V1.0/PACS/Controller/Event/Notifications")]
        public ActionResult test()
        {
            StreamReader reader = new StreamReader(HttpContext.Request.InputStream);
            string requestFromPost = reader.ReadToEnd();
            ViewData["Stream"] = requestFromPost;
            return View();
        }

        public ActionResult readJson()
        {
            //get the Json filepath  
            string file = Server.MapPath("~/content/returns_24102017.json");
            //deserialize JSON from file  
            string json = System.IO.File.ReadAllText(file);
            JavaScriptSerializer ser = new JavaScriptSerializer();
            //DataSet data = JsonConvert.DeserializeObject<DataSet>(json);


            var jsonLinq = JObject.Parse(json);

            // Find the first array using Linq
            var srcArray = jsonLinq.Descendants().Where(d => d is JArray).First();
            var trgArray = new JArray();
            foreach (JObject row in srcArray.Children<JObject>())
            {
                var cleanRow = new JObject();
                foreach (JProperty column in row.Properties())
                {
                    // Only include JValue types
                    if (column.Value is JValue)
                    {
                        cleanRow.Add(column.Name, column.Value);
                    }
                }

                trgArray.Add(cleanRow);
            }
            DataTable dt = JsonConvert.DeserializeObject<DataTable>(trgArray.ToString());
            return View();
        }
        public ActionResult TrackLatLong()
        {
            // comfun.saveformname("TrackLatLong", "/home/TrackLatLong", "Track Lat Long Main Form", "Track Lat Long Main Form", "N");
            DataTable dt = comfun.fillDataTable("ApiEmp_SelectForTracking", "", null);
            return View("TrackLatLong2703", dt);
        }

        //[customAuthorize(Roles = _roles)]
        public JsonResult _latlongJson()
        {
            // comfun.saveformname("_latlongJson", "/home/_latlongJson", "Track Lat Long json", "Track Lat Long json inner function", "N");
            /*DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("lat"));
            dt.Columns.Add(new DataColumn("lng"));

            DataRow row = dt.NewRow();
            row["lat"] = 33.779005;
            row["lng"] = -118.178985;
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["lat"] = 33.879005;
            row["lng"] = -118.098985;
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["lat"] = 33.979005;
            row["lng"] = -118.218985;
            dt.Rows.Add(row);
            
            */

            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@FromDate", Request.Form["Date"].ToString());
            list.Add("@ToDate", Convert.ToDateTime(Convert.ToDateTime(Request.Form["Date"]).AddDays(1)).ToString("dd-MMM-yyyy"));
            DataTable dt = comfun.fillDataTable("ApiEmp_GetLatLong", "", list);

            // string jsonString = string.Empty;
            // jsonString = Newtonsoft.Json.JsonConvert.SerializeObject( dt, Formatting.Indented);
            //ViewData["map"] = PlotGPSPoints(dt);

            DataTable dt2 = new DataTable();
            dt2.Columns.Add(new DataColumn("address"));
            dt2.Columns.Add(new DataColumn("lat"));
            dt2.Columns.Add(new DataColumn("lng"));
            /*DataRow row = dt2.NewRow();
            string lat = string.Empty;
            string lng = string.Empty;
            string latpre = string.Empty;
            string lngpre = string.Empty;
            string status = string.Empty;
            int loopCount = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (!dt.Rows[i]["Lng"].ToString().ToLower().Contains("day"))
                {
                    loopCount = 0;
                    if (Convert.ToDecimal(dt.Rows[i]["Lat"]) > 0)
                    {
                        if (i > 0)
                        {
                            //if ((dt.Rows[i - 1]["lat"].ToString() != dt.Rows[i]["lat"].ToString()) && (dt.Rows[i - 1]["lng"].ToString() != dt.Rows[i]["lng"].ToString()))
                            {
                            Status:
                                status = comfun.getAddressFromLatLong(dt.Rows[i]["Lat"].ToString(), dt.Rows[i]["Lng"].ToString());

                                if (status.Contains("OVER_QUERY_LIMIT"))
                                {
                                    loopCount = loopCount + 1;
                                    if (loopCount < 2)
                                    {
                                        goto Status;
                                    }
                                }
                            }
                        }
                        else
                        {
                        Status:
                            status = comfun.getAddressFromLatLong(dt.Rows[i]["Lat"].ToString(), dt.Rows[i]["Lng"].ToString());

                            if (status.Contains("OVER_QUERY_LIMIT"))
                            {
                                loopCount = loopCount + 1;
                                if (loopCount < 2)
                                {
                                    goto Status;
                                }
                            }
                        }
                        row = dt2.NewRow();
                        row["address"] = status;

                        row["lat"] = dt.Rows[i]["Lat"].ToString();
                        row["lng"] = dt.Rows[i]["Lng"].ToString();
                        dt2.Rows.Add(row);
                    }
                    //calcularRota(dt.Rows[i]["Lat"].ToString(), dt.Rows[i]["Lng"].ToString()); 

                }
            }
            */
            string jsonString = string.Empty;
            jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(dt, Formatting.Indented);
            return Json(jsonString);
        }


        //  [customAuthorize(Roles = _roles)]
        public JsonResult _getLocationTable()
        {
            // comfun.saveformname("_getLocationTable", "/home/_getLocationTable", "Track Lat Long Location table", "Track Lat Long Location table", "N");
            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@FromDate", Request.Form["Date"].ToString());
            list.Add("@ToDate", Convert.ToDateTime(Convert.ToDateTime(Request.Form["Date"]).AddDays(1)).ToString("dd-MMM-yyyy"));
            DataTable dt = comfun.fillDataTable("ApiEmp_GetLatLong", "", list);


            //return sCoord.GetDistanceTo(eCoord);
            string table = "<table class='table'>";
            table += "<tr>";
            table += "<th class='text-center'>Sr.</th>";
            table += "<th class='text-center'>Location</th>";
            table += "<th class='text-center'>KM</th>";
            table += "<th class='text-center'>Time</th>";
            table += "<th class='text-center'>Tag</th>";
            table += "</tr>";


            var latStart = "";
            var lngStart = "";
            var latEnd = "";
            var lngEnd = "";
            string oldAdderss = "";
            var startLoc = new GeoCoordinate();
            var endLoc = new GeoCoordinate();
            var km = (double)0;
            var totalKM = (double)0;
            int loopCount = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                loopCount = 0;
                if (i > 0)
                {
                    latStart = (dt.Rows[i - 1]["Lat"].ToString() == "" ? "0" : dt.Rows[i - 1]["Lat"].ToString());
                    lngStart = (dt.Rows[i - 1]["Lng"].ToString() == "" ? "0" : dt.Rows[i - 1]["Lng"].ToString());
                    latEnd = (dt.Rows[i]["Lat"].ToString() == "" ? "0" : dt.Rows[i]["Lat"].ToString());
                    lngEnd = (dt.Rows[i]["Lng"].ToString() == "" ? "0" : dt.Rows[i]["Lng"].ToString());
                    startLoc = new GeoCoordinate(Convert.ToDouble(latStart), Convert.ToDouble(lngStart));
                    endLoc = new GeoCoordinate(Convert.ToDouble(latEnd), Convert.ToDouble(lngEnd));
                    km = startLoc.GetDistanceTo(endLoc) / 1000;
                    totalKM = totalKM + km;
                }
                table += "<tr>";
                table += "  <td class='text-center'>" + (i + 1).ToString() + ".</td>";


            Start:
                if (!dt.Rows[i]["Lng"].ToString().ToLower().Contains("Day"))
                {

                    oldAdderss = comfun.getAddressFromLatLong(dt.Rows[i]["Lat"].ToString(), dt.Rows[i]["Lng"].ToString());
                    if (oldAdderss == "OVER_QUERY_LIMIT")
                    {
                        loopCount = loopCount + 1;
                        if (loopCount < 4)
                        {
                            Thread.Sleep(1000);
                            goto Start;
                        }
                    }
                    //else
                    //{
                    table += "<td class='text-center'>" + oldAdderss + "</td>";
                    //}
                }
                else
                {
                    table += "<td class='text-center'>" + oldAdderss + "</td>";
                }
                // table += "<td class='text-center'></td>";
                table += "<td class='text-center'>" + km.ToString("0.00") + "</td>";
                table += "<td class='text-center'>" + (dt.Rows[i]["CreatedDate"].ToString() == "" ? "" : Convert.ToDateTime(dt.Rows[i]["CreatedDate"]).ToString("dd-MMM-yyyy HH:mm:ss tt")) + "</td>";
                table += "<td class='text-center'>" + dt.Rows[i]["LocationTag"].ToString() + "</td>";
                //table += "<td class='text-center'><a href='#' onclick='getAttress("+dt.Rows[i]["Lat"].ToString()+","+ dt.Rows[i]["Lng"].ToString() + ")'>Address</a></td>";
                table += "</tr>";
            }
            table += "<tr>";
            table += "  <td class='text-right' colspan='2'>Total KMs</td>";
            table += "  <td class='text-center'>" + totalKM.ToString("0.00") + "</td>";
            table += "  <td class='text-center'></td>";
            table += "</tr>";
            table += "</table>";
            return Json(table);
        }

        //[customAuthorize(Roles = _roles)]
        public ActionResult _latlongTrack()
        {
            // comfun.saveformname("_latlongTrack", "/home/_latlongTrack", "Track Lat Long Partial View", "Track Lat Long Partial View", "N");
            //SortedList list = new SortedList();
            //list.Add("@EmpId", "");
            //list.Add("@DayId", "");
            //DataTable dt = comfun.fillDataTable("", "", list);
            return PartialView("_latlongTrack");
        }
        public ActionResult _SaleOrderStatus()
        {
            SortedList list = new SortedList();

            list.Add("@Status", Request.Form["Tag"]);
            list.Add("@FromDate", Request.Form["FromDate"]);
            list.Add("@ToDate", Request.Form["ToDate"]);

            list.Add("@LoginID", Session["EmpId"]);
            list.Add("@SessionID", Session["SessionId"]);
            DataSet dt = comfun.fillDataSet("stpSaleExecutive_Dashboard", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult _SaleOrderHeadStatus()
        {
            SortedList list = new SortedList();

            list.Add("@Status", Request.Form["Tag"]);
            list.Add("@LoginID", Session["EmpId"]);
            list.Add("@SessionID", Session["SessionId"]);
            list.Add("@FromDate", Request.Form["FromDate"]);
            list.Add("@ToDate", Request.Form["ToDate"]);
            DataSet dt = comfun.fillDataSet("stpSaleHead_Dashboard", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult _SaleSCMStatus()
        {
            SortedList list = new SortedList();

            list.Add("@Status", Request.Form["Tag"]);
            list.Add("@FromDate", Request.Form["FromDate"]);
            list.Add("@ToDate", Request.Form["ToDate"]);
            list.Add("@LoginID", Session["EmpId"]);
            list.Add("@SessionID", Session["SessionId"]);
            DataSet dt = comfun.fillDataSet("stpSCMHead_Dashboard", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }

        public ActionResult _SearchEmpDetails()
        {
            DataTable dt = new DataTable();

            decimal total_records = 0;
            decimal pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            decimal pagesize = Convert.ToDecimal(Request.Form["PageSize"]);
            SortedList list = new SortedList();

            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            list.Add("@Filter", Request.Form["Filter"]);

            //list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("stpSearchEmpDetails", "", list);


            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);


            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "Home", "_SearchEmpDetails", "_SearchEmpDetails");
            HtmlString htm = new HtmlString(paging);
            ViewData["paging"] = htm;
            var json = JsonConvert.SerializeObject(dt);
            return Json(new
            {
                Paging = paging,
                data = json
            });
        }
    }
}