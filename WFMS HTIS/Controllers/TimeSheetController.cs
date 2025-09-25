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
using DocumentFormat.OpenXml.Bibliography;
using Payroll.portal.SessionLogout;

namespace Payroll.portal.Controllers
{
    //[Authorize]
    [SessionExpire]
    public class TimeSheetController : Controller
    {
        protected System.Web.UI.HtmlControls.HtmlInputFile Images;
        #region Payroll Management System
        public const string _roles = payrollFunctions._rolesGlobal;
        commonFunctions comfun = new commonFunctions();
        payrollFunctions payfun = new payrollFunctions();

        public ActionResult Index()
        {

            return View();
        }

        public ActionResult CountriesEdit()
        {
            //comfun.saveformname("CountriesEdit", "/admin/CountriesEdit", "Country Edit", "Country Edit view", "N", "Country Edit", "Country", "Edit");

            string countryid = "0";
            if (Request.QueryString["id"] == null)
            {
                return RedirectToAction("Countries");
            }
            countryid = comfun.decryptString(Request.QueryString["Id"].ToString());
            SortedList list = new SortedList();
            list.Add("@CountryId", countryid);
            DataTable dt = comfun.fillDataTable("Country_SelectWithId", "", list);
            return View(dt);
        }

        [customAuthorize(Roles = payrollFunctions.rolePayrollTeam)]
        public ActionResult PayCharges(int? page)
        {
            //comfun.saveformname("PayCharges", "/admin/PayCharges", "PayCharges", "PayCharges Main form", "Y", "", "PayCharges", "List");
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }

            //dt = comfun.fillDataTable("Charges_Select", "", null);

            return View();
        }

        public ActionResult _city()
        {
            SortedList list = new SortedList();

            list.Add("@StateId", Request.Form["State"].ToString());

            DataTable dt = comfun.fillDataTable("City_SelectByStateId", "", list);
            return PartialView("_CityddlSelect", dt);
        }
        public ActionResult _city_selectddl()
        {
            SortedList list = new SortedList();
            list.Add("@StateId", Request.Form["StateId"].ToString());
            DataTable dt = comfun.fillDataTable("City_SelectByStateId", "", list);
            return PartialView("_city_ddl", dt);
        }



        //  [customAuthorize(Roles = payrollFunctions.rolePayrollTeam)]

        public ActionResult _Employeeddl()
        {
            DataTable dt = comfun.fillDataTable("stpEmployeeddl", "", null);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }



        public ActionResult CompanyList()
        {
            comfun.saveformname("CompanyList", "/admin/CompanyList", "Master/Organization/Company List", "Company List", "N", "CompanyList", "CompanyList", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("CompanyList");
            DataSet ds = comfun.fillDataSet("stpViksatDivisionStateCity", "", null);
            return View("DivisionList", ds);
            //DataTable dt = comfun.fillDataTable("stpviksatDivisionSelect", "", null);
            //return View("DivisionList", dt);
        }
        public ActionResult CompanyAdd()
        {
            //comfun.saveformname("CompanyAdd", "/admin/CompanyAdd", "Company Add", "Company Add", "N", "CompanyAdd", "CompanyAdd", "Add");
            DataSet ds = comfun.fillDataSet("stpViksatDivisionStateCity", "", null);
            return View("DivisionAdd", ds);
        }
        public ActionResult CompanyEdit()
        {
            //comfun.saveformname("CompanyEdit", "/admin/CompanyEdit", "Company Edit", "Company Edit", "N", "CompanyEdit", "CompanyEdit", "Edit");

            string DivisionId = "0";
            if (Request.QueryString["id"] == null)
            {
                return RedirectToAction("CompanyList");
            }
            DivisionId = comfun.decryptString(Request.QueryString["Id"].ToString());

            SortedList list = new SortedList();

            list.Add("@DivisionID", DivisionId);
            DataSet dt = comfun.fillDataSet("stpviksatDivisionSelectWithID", "", list);
            return View("DivisionEdit", dt);
        }
        public ActionResult UnitAdd()
        {
            //comfun.saveformname("UnitAdd", "/admin/UnitAdd", "Unit Add", "Unit view", "N", "UnitAdd", "Unit", "Add");

            DataSet ds = comfun.fillDataSet("stpViksatBranchCompanyStateCity", "", null);
            return View(ds);
        }
        public ActionResult UnitEdit()
        {
            //comfun.saveformname("UnitEdit", "/admin/UnitEdit", "Unit Edit", "Unit Edit", "N", "UnitEdit", "Unit", "Edit");

            string UnitId = "0";
            if (Request.QueryString["id"] == null)
            {
                return RedirectToAction("UnitList");
            }
            UnitId = comfun.decryptString(Request.QueryString["Id"].ToString());

            SortedList list = new SortedList();

            list.Add("@UnitId", UnitId);
            DataSet dt = comfun.fillDataSet("stpViksatUnit_SelectWithId", "", list);
            return View(dt);
        }
        public JsonResult MaritalStatusUpdate()
        {
            comfun.saveformname("MaritalStatusUpdate", "/admin/MaritalStatusUpdate", "Master/Address/Tier", "Tier Status", "N", "Tier", "Tier", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@MaritalStatusId", Request.Form["MaritalStatusId"].ToString());
                list.Add("@MaritalStatusActivation", Request.Form["MaritalStatusActivation"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_MaritalStatus_Active_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }

        public ActionResult AddNewEmpStatus()
        {
            return View();
        }

        //To Save/Update Employee Status//
        public JsonResult EmpStatusSaveUpdateSubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@EmpStatusId", Request.Form["EmpStatusId"].ToString());
                list.Add("@EmpStatusCode", Request.Form["EmpStatusCode"].ToString());
                list.Add("@EmpStatusName", Request.Form["EmpStatusName"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_EmpStatus_SaveUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        public ActionResult UpdateEmpStatus()
        {
            string empstatusid = "0";
            if (Request.QueryString["id"] == null)
            {
                return RedirectToAction("EmpStatusList", "EmpStatusList");
            }
            empstatusid = Request.QueryString["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@EmpStatusId", empstatusid);
            DataTable dt = comfun.fillDataTable("sp_EmpStatusDisplayViaId", "", list);
            return View(dt);
        }

        //To Delete Existing Employee Status//
        public JsonResult DeleteEmpStatus()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@EmpStatusId", Request.Form["EmpStatusId"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_EmpStatusDelete", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes, JsonRequestBehavior.AllowGet);
        }

        //To Change Active/Deactive Employee Status//
        public JsonResult EmpActivationStatusUpdate()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@EmpStatusId", Request.Form["EmpStatusId"].ToString());
                list.Add("@EmpactivationStatus", Request.Form["EmpActivationStatus"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_EmpStatus_Active_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }


        //////////////////////////////
        ///Employee Type ((DONE))
        /////////////////////////////
        ///

        public ActionResult AddNewEmpType()
        {
            return View();
        }

        //To Save/Update Employee Type//
        public JsonResult EmpTypeSaveUpdateSubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@EmpTypeId", Request.Form["EmpTypeId"].ToString());
                list.Add("@EmpTypeCode", Request.Form["EmpTypeCode"].ToString());
                list.Add("@EmpTypeName", Request.Form["EmpTypeName"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_EmpType_SaveUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }

        ////To Display Employee Type List//
        //public ActionResult EmpTypeList()
        //{
        //    DataTable dt = comfun.fillDataTable("sp_EmpTypeListDisplay", "", null);
        //    return View(dt);
        //}

        ////To Get Employee Type Details Via Id//
        //public ActionResult UpdateEmpType()
        //{
        //    string emptypeid = "0";
        //    if (Request.QueryString["id"] == null)
        //    {
        //        return RedirectToAction("EmpTypeList", "EmpTypeList");
        //    }
        //    emptypeid = Request.QueryString["Id"].ToString();
        //    SortedList list = new SortedList();
        //    list.Add("@EmpTypeId", emptypeid);
        //    DataTable dt = comfun.fillDataTable("sp_EmpTypeDisplayViaId", "", list);
        //    return View(dt);
        //}

        ////To Delete Existing Employee Type//
        //public JsonResult DeleteEmpType()
        //{
        //    string mes = string.Empty;
        //    try
        //    {
        //        SortedList list = new SortedList();
        //        list.Add("@EmpTypeId", Request.Form["EmpTypeId"].ToString());
        //        mes = comfun.executeNonQueryWMessage("sp_EmpTypeDelete", "", list).ToString();
        //    }
        //    catch (Exception ex)
        //    {
        //        mes = ex.Message;
        //    }
        //    return Json(mes, JsonRequestBehavior.AllowGet);
        //}

        ////To Change Active/Deactive Employee Type//
        //public JsonResult EmpTypeActivationStatusUpdate()
        //{
        //    string mes = string.Empty;
        //    try
        //    {
        //        SortedList list = new SortedList();
        //        list.Add("@EmpTypeId", Request.Form["EmpTypeId"].ToString());
        //        list.Add("@EmpTypeStatus", Request.Form["EmpTypeActivationStatus"].ToString());
        //        mes = comfun.executeNonQueryWMessage("sp_EmpType_Active_Deactive", "", list).ToString();
        //    }
        //    catch (Exception ex)
        //    {
        //        mes = ex.Message;
        //    }
        //    return Json(mes);
        //}


        public ActionResult Conveyance()
        {
            //comfun.saveformname("Conveyance", "/admin/Conveyance", "Conveyance Add", "Conveyance Add", "N", "Conveyance", "Conveyance", "Add");
            DataSet ds = comfun.fillDataSet("stpVIKSATEmployeeslist_PMAll", "", null);
            return View(ds);
        }
        public ActionResult _branchdata()
        {
            //SqlDataAdapter da = new SqlDataAdapter("Branch_Select", con);
            DataTable dt = new DataTable();
            //da.Fill(dt);
            SortedList list = new SortedList();
            dt = comfun.fillDataTable("Branch_Select", "", list);
            return PartialView("_branchdatas", dt);
        }
        public JsonResult _ConveyanceMasterAddSubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@AMDate", Request.Form["AMDate"].ToString());
                list.Add("@fiEmployeeId", Request.Form["fiEmployeeId"].ToString());
                list.Add("@fvFromLoaction", Request.Form["fvFromLoaction"].ToString());
                list.Add("@fvStatus", Request.Form["fvStatus"].ToString());
                list.Add("@fvToLoaction", Request.Form["fvToLoaction"].ToString());
                list.Add("@BranchId", Request.Form["BranchId"].ToString());
                list.Add("@fvAmount", Request.Form["fvAmount"].ToString());
                mes = comfun.executeNonQueryWMessage("Conveyance_Save", "", list).ToString();
            }
            catch (Exception ex)
            {

                mes = comfun.errorMessage("Admin _Conveyance", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult AddNewPolicyCategory()
        {
            return View();
        }

        //To Save/Update Policy Category//
        public JsonResult PolicyCategorySaveUpdateSubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@PolicyCategoryId", Request.Form["PolicyCategoryId"].ToString());
                list.Add("@PolicyCategoryCode", Request.Form["PolicyCategoryCode"].ToString());
                list.Add("@PolicyCategoryName", Request.Form["PolicyCategoryName"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_PolicyCategory_SaveUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }

        //To Display Policy Category List//
        public ActionResult PolicyCategoryaList()
        {
            DataTable dt = comfun.fillDataTable("sp_PolicyCategoryListDisplay", "", null);
            return View(dt);
        }

        //To Get Policy Category Details Via Id//
        public ActionResult UpdatePolicyCategory()
        {
            string policycategoryid = "0";
            if (Request.QueryString["id"] == null)
            {
                return RedirectToAction("PolicyCategoryList", "PolicyCategoryList");
            }
            policycategoryid = Request.QueryString["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@PolicyCategoryId", policycategoryid);
            DataTable dt = comfun.fillDataTable("sp_PolicyCategoryDisplayViaId", "", list);
            return View(dt);
        }

        //To Delete Existing Policy Category//
        public JsonResult DeletePolicyCategory()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@PolicyCategoryId", Request.Form["ZoneId"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_PolicyCategory_Delete", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes, JsonRequestBehavior.AllowGet);
        }

        //To Change Active/Deactive Policy Category//
        public JsonResult PolicyCategoryActivationStatusUpdate()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@PolicyCategoryId", Request.Form["PolicyCategoryId"].ToString());
                list.Add("@PolicyCategoryStatus", Request.Form["PolicyCategoryStatus"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_PolicyCategory_Active_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }

        //////////////////////////////
        ///Loan Type ((DONE))
        /////////////////////////////
        ///

        //public ActionResult AddNewLoanType()
        //{
        //    return View();
        //}

        ////To Save/Update Loan Type//
        //public JsonResult LoanTypeSaveUpdateSubmit()
        //{
        //    string mes = string.Empty;
        //    try
        //    {
        //        SortedList list = new SortedList();
        //        list.Add("@LoanTypeId", Request.Form["LoanTypeId"].ToString());
        //        list.Add("@LoanTypeCode", Request.Form["LoanTypeCode"].ToString());
        //        list.Add("@LoanTypeName", Request.Form["LoanTypeName"].ToString());
        //        mes = comfun.executeNonQueryWMessage("sp_LoanType_SaveUpdate", "", list).ToString();
        //    }
        //    catch (Exception ex)
        //    {
        //        mes = ex.Message;
        //    }
        //    return Json(mes);
        //}

        ////To Display Loan Type List//
        ////public ActionResult LoanTypeList()
        ////{
        ////    DataTable dt = comfun.fillDataTable("sp_LoanTypeListDisplay", "", null);
        ////    return View(dt);
        ////}
        ////To Get Loan Type Details Via Id//
        //public ActionResult UpdateLoanType()
        //{
        //    string loantypeid = "0";
        //    if (Request.QueryString["id"] == null)
        //    {
        //        return RedirectToAction("LoanTypeList", "LoanTypeList");
        //    }
        //    loantypeid = Request.QueryString["Id"].ToString();
        //    SortedList list = new SortedList();
        //    list.Add("@LoanTypeId", loantypeid);
        //    DataTable dt = comfun.fillDataTable("sp_LoanTypeDisplayViaId", "", list);
        //    return View(dt);
        //}
        ////To Delete Existing Loan Type//
        //public JsonResult DeleteLoanType()
        //{
        //    string mes = string.Empty;
        //    try
        //    {
        //        SortedList list = new SortedList();
        //        list.Add("@LoanTypeId", Request.Form["LoanTypeId"].ToString());
        //        mes = comfun.executeNonQueryWMessage("sp_LoanType_Delete", "", list).ToString();
        //    }
        //    catch (Exception ex)
        //    {
        //        mes = ex.Message;
        //    }
        //    return Json(mes, JsonRequestBehavior.AllowGet);
        //}
        ////To Change Active/Deactive Loan Type//
        //public JsonResult LoanTypeActivationStatusUpdate()
        //{
        //    string mes = string.Empty;
        //    try
        //    {
        //        SortedList list = new SortedList();
        //        list.Add("@LoanTypeId", Request.Form["LoanTypeId"].ToString());
        //        list.Add("@LoanTypeStatus", Request.Form["LoanTypeStatus"].ToString());
        //        mes = comfun.executeNonQueryWMessage("sp_LoanType_Active_Deactive", "", list).ToString();
        //    }
        //    catch (Exception ex)
        //    {
        //        mes = ex.Message;
        //    }
        //    return Json(mes);
        //}


        public ActionResult HelpDesk()
        {
            return View();
        }
        public ActionResult GetEmployee()
        {
            //DataTable dt = new DataTable();
            //SqlCommand cmd = new SqlCommand("stpVIKSATPMEmployees_Alllist", con);
            //cmd.CommandType = System.Data.CommandType.StoredProcedure;
            //SqlDataAdapter da = new SqlDataAdapter(cmd);
            //da.Fill(dt);
            //return PartialView("ddlEmployee", dt);

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            dt = comfun.fillDataTable("stpVIKSATPMEmployees_Alllist", "", null);
            return PartialView("ddlEmployee", dt);

        }
        public ActionResult GridQuery()
        {
            //SqlDataAdapter da = new SqlDataAdapter("stpVIKSATPMEmployees_helpDeskShow", con);
            //DataTable dt = new DataTable();
            //da.Fill(dt);
            //return PartialView("_GridHelpDesk", dt);

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            dt = comfun.fillDataTable("stpVIKSATPMEmployees_helpDeskShow", "", null);
            return PartialView("_GridHelpDesk", dt);
        }
        public JsonResult _QuerySave()
        {
            //try
            //{
            //    string Message = string.Empty;
            //    //string SessionId = string.Empty;
            //    //SessionId = Session["LoginID"].ToString();
            //    DataTable dt = new DataTable();

            //    SqlCommand cmd = new SqlCommand("stpVIKSATPMEmployees_helpDesk", con);
            //    //cmd.Parameters.AddWithValue("@SessionId", SessionId);
            //    cmd.Parameters.AddWithValue("@AMDate", Request.Form["AMDate"].ToString());
            //    cmd.Parameters.AddWithValue("@Title", Request.Form["Title"].ToString());
            //    cmd.Parameters.AddWithValue("@fvQuery", Request.Form["fvQuery"].ToString());
            //    cmd.Parameters.AddWithValue("@fiEmployeeID", Request.Form["fiEmployeeID"].ToString());
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
                list.Add("@AMDate", Request.Form["AMDate"].ToString());
                list.Add("@Title", Request.Form["Title"].ToString());
                list.Add("@fvQuery", Request.Form["fvQuery"].ToString());
                list.Add("@fiEmployeeID", Request.Form["fiEmployeeID"].ToString());
                mes = comfun.executeNonQueryWMessage("stpVIKSATPMEmployees_helpDesk", "", list).ToString();
            }
            catch (Exception ex)
            {

                mes = comfun.errorMessage("Admin helpdesk", ex.Message);
            }
            return Json(mes);

        }

        //////////////////////////////
        ///Clearance Group ((DONE))
        /////////////////////////////
        ///
        public ActionResult AddNewCG()
        {
            return View();
        }
        //To Save/Update Clearance Group//
        public JsonResult CGSaveUpdateSubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@CGId", Request.Form["CGId"].ToString());
                list.Add("@CGCode", Request.Form["CGCode"].ToString());
                list.Add("@CGName", Request.Form["CGName"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_CG_SaveUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        //To Display Clearance Group List//
        public ActionResult CGLista()
        {
            DataTable dt = comfun.fillDataTable("sp_CGListDisplay", "", null);
            return View(dt);
        }
        //To Get Clearance Group Details Via Id//
        public ActionResult UpdateCG()
        {
            string cgid = "0";
            if (Request.QueryString["id"] == null)
            {
                return RedirectToAction("CGList", "CGList");
            }
            cgid = Request.QueryString["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@CGId", cgid);
            DataTable dt = comfun.fillDataTable("sp_CGDisplayViaId", "", list);
            return View(dt);
        }
        //To Delete Existing Clearance Group//
        public JsonResult DeleteCG()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@CGId", Request.Form["CGId"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_CG_Delete", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes, JsonRequestBehavior.AllowGet);
        }
        //To Change Active/Deactive Clearance Group//
        public JsonResult CGActivationStatusUpdate()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@CGId", Request.Form["CGId"].ToString());
                list.Add("@CGStatus", Request.Form["CGStatus"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_CG_Active_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        /// <summary>
        /// Group and Clearance Head Mapping Details  ((PENDING))
        /// </summary>
        /// <returns></returns>
        public ActionResult AddNewCGM()
        {
            DataSet ds = comfun.fillDataSet("sp_CGM_DDL_Data", "", null);
            return View(ds);
        }
        //To Display Employee List based on Selected Department Id//
        public ActionResult _EmpListDisplayViaDepartmentId()
        {
            SortedList list = new SortedList();
            list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
            DataTable dt = comfun.fillDataTable("sp_EmpListViaDepartmentId", "", list);
            return PartialView("_EmpListDisplayViaDepartmentId", dt);
        }

        #endregion
        #region Munish


        public ActionResult _Pageno()
        {
            DataTable dt = comfun.fillDataTable("stpPageMaster", null, null);

            return PartialView("_PageNo", dt);
        }

        public ActionResult Country()
        {

            comfun.saveformname("Country", "/TimeSheet/Country", "Master/Address/Country", "Country  Main form", "Y", "Country", "Country", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("Country");
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            DataTable dt = comfun.fillDataTable("Country_Select", "", null);
            return View(dt);
        }

        public ActionResult _countryList()
        {
            comfun.saveformname("_countryList", "/TimeSheet/_countryList", "Master/Address/Country", "Country  list", "N", "Country", "Country", "List", 4);
            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 50;
            if (Request.Form["Page"] != null)
            {
                pagesize = Convert.ToDecimal(Request.Form["Page"]);
            }
            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }
            SortedList list = new SortedList();
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            DataTable dt = comfun.fillDataTable("Country_Select", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public JsonResult _CountrySubmit()
        {
            comfun.saveformname("_CountrySubmit", "/TimeSheet/_CountrySubmit", "Master/Address/Country", "Country  add", "N", "Country", "Country", "Add", 2);
            comfun.saveformname("_CountrySubmit", "/TimeSheet/_CountrySubmit", "Master/Address/Country", "Country  edit", "N", "Country", "Country", "Edit", 3);

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {

                list.Add("@CountryName", Request.Form["CountryName"].ToString());
                list.Add("@CountryCode", Request.Form["CountryCode"].ToString());
                list.Add("@CountryId", Request.Form["CountryId"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                list.Add("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                mes = comfun.executeNonQueryWMessage("Country_Save", "", list).ToString();
                if (mes.Contains("Error"))
                {
                    list.Clear();
                    list.Add("Error", mes);
                }
                else
                {
                    list.Clear();
                    list.Add("Error", "");
                    list.Add("CountryId", mes);
                    list.Add("CountryName", Request.Form["CountryName"].ToString());
                    list.Add("CountryCode", Request.Form["CountryCode"].ToString());
                }
            }
            catch (Exception ex)
            {
                list.Clear();
                list.Add("Error", mes);
            }
            return Json(list);
        }
        public JsonResult _CountryStatusUpdate()
        {
            comfun.saveformname("_CountryStatusUpdate", "/TimeSheet/_CountryStatusUpdate", "Master/Address/Country", "Country  Status", "N", "Country", "Country", "Status", 5);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpCountryActivationStatusUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }

        public ActionResult States()
        {
            comfun.saveformname("State", "/TimeSheet/States", "Master/Address/States", "State  Main Form", "Y", "State", "State", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("State");
            DataTable dt = comfun.fillDataTable("CountryState_Select", "", null);
            return View(dt);
        }
        public ActionResult _state_selectddl()
        {
            SortedList list = new SortedList();
            list.Add("@CountryId", Request.Form["CountryId"].ToString());
            DataTable dt = comfun.fillDataTable("State_SelectByCountryId", "", list);
            return PartialView("_state_ddl", dt);
        }
        public JsonResult _StateStatusUpdate()
        {
            comfun.saveformname("_StateStatusUpdate", "/TimeSheet/_StateStatusUpdate", "Master/Address/States", "State  Status", "N", "State", "State", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpStateStatusUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }

        public ActionResult _stateListJson()
        {
            comfun.saveformname("_stateListJson", "/TimeSheet/_stateListJson", "Master/Address/States", "State List", "N", "State", "State", "List", 4);
            SortedList list = new SortedList();
            decimal pageno = 1;
            decimal pagesize = 50;
            if (Request.Form["Page"] != null)
            {
                pagesize = Convert.ToDecimal(Request.Form["Page"]);
            }
            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }
            list.Add("@CountryId", Request.Form["CountryId"].ToString());
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);

            DataTable dt = comfun.fillDataTable("stpviksatStateSelectByCountryId", "", list);
            ViewData["CountryId"] = Request.Form["CountryId"].ToString();
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }




        public JsonResult _StateSubmit()
        {
            comfun.saveformname("_StateSubmit", "/TimeSheet/_StateSubmit", "Master/Address/States", "States Add", "N", "State", "State", "Add", 2);
            comfun.saveformname("_StateSubmit", "/TimeSheet/_StateSubmit", "Master/Address/States", "States Edit", "N", "State", "State", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@StateName", Request.Form["StateName"].ToString());
                list.Add("@StateId", Request.Form["StateId"].ToString());
                list.Add("@StateCode", Request.Form["StateCode"].ToString());
                list.Add("@CountryId", Request.Form["CountryId"].ToString());
                list.Add("@CreatedBy", "1");
                list.Add("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                mes = comfun.executeNonQueryWMessage("State_Save", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_StateSubmit", ex.Message);
            }
            return Json(mes);
        }

        /// <summary>
        /// City Master
        /// </summary>
        /// <returns></returns>

        public ActionResult Zone()
        {
            comfun.saveformname("Zone", "/admin/Zone", "Master/Address/Zone", "Zone  Main form", "Y", "Zone", "Zone", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("Zone");
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            return View();
        }
        public ActionResult _zones()
        {
            comfun.saveformname("_zones", "/admin/_zones", "Master/Address/Zone", "Zone List", "N", "Zone", "Zone", "List", 4);
            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 10;
            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }
            SortedList list = new SortedList();
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            DataTable dt = comfun.fillDataTable("sp_ZoneListDisplay", "", list);
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);
            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "admin", "_zones", "_zones");
            HtmlString htm = new HtmlString(paging);
            ViewData["paging"] = htm;
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }

        public JsonResult _ZoneAddSubmit()
        {

            comfun.saveformname("_ZoneAddSubmit", "/admin/_ZoneEditSubmit", "Master/Address/Zone", "Zone  Add", "N", "Zone", "Zone", "Add", 2);
            comfun.saveformname("_ZoneAddSubmit", "/admin/_ZoneEditSubmit", "Master/Address/Zone", "Zone  Edit", "N", "Zone", "Zone", "Edit", 3);
            string mes = string.Empty;
            try
            {


                SortedList list = new SortedList();
                list.Add("@ZoneId", Request.Form["ZoneId"].ToString());
                list.Add("@ZoneCode", Request.Form["ZoneCode"].ToString());
                list.Add("@ZoneName", Request.Form["ZoneName"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_Zone_SaveUpdate", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = "Error: " + comfun.errorMessage("_ZoneAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        public JsonResult _ZoneActivationStatusUpdate()
        {
            comfun.saveformname("_ZoneActivationStatusUpdate", "/admin/_ZoneActivationStatusUpdate", "Master/Address/Zone", "Zone  Status", "N", "Zone", "Zone", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@ZoneId", Request.Form["Id"].ToString());
                list.Add("@ZoneStatus", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_Zone_Active_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }
        public ActionResult _ZoneEdit()
        {
            SortedList list = new SortedList();
            list.Add("@ZoneId", Request.Form["Ids"].ToString());
            DataTable dt = comfun.fillDataTable("sp_ZoneDisplayViaId", "", list);
            return PartialView("_ZoneEdit", dt);
        }
        public JsonResult _ZoneEditSubmit()
        {

            comfun.saveformname("_ZoneEditSubmit", "/admin/_ZoneEditSubmit", "Master/Address/Zone", "Zone Add", "N", "Zone", "Zone", "Add", 2);
            comfun.saveformname("_ZoneEditSubmit", "/admin/_ZoneEditSubmit", "Master/Address/Zone", "Zone Edit", "N", "Zone", "Zone", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@ZoneId", Request.Form["ZoneId"].ToString());
                list.Add("@ZoneCode", Request.Form["ZoneCode"].ToString());
                list.Add("@ZoneName", Request.Form["ZoneName"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_Zone_SaveUpdate", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        public ActionResult Banks()
        {
            comfun.saveformname("BankMaster", "/TimeSheet/Banks", "Master/Bank Detail/Bank", "Bank Master form", "Y", "BankMaster", "BankMaster", "View", 1);
            //  ViewData["AccessRights"] = payfun.getAccessRights("BankMaster");
            return View();
        }
        public ActionResult _BankList()
        {
            comfun.saveformname("_BankList", "/TimeSheet/Banks", "Master/Bank Detail/Bank", "Bank Master List", "N", "BankMaster", "BankMaster", "List", 2);
            DataTable dt = comfun.fillDataTable("Banks_List", "", null);
            //return PartialView("_BankList", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _BankAdd()
        {
            // comfun.saveformname("BanksAdd", "/admin/BanksAdd", "Bank Add", "Bank Add view", "N", "BanksAdd", "Bank", "Add");
            return PartialView("_BankAdd");
        }
        public JsonResult _BanksAddSubmit()
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}
            comfun.saveformname("_BanksAddSubmit", "/TimeSheet/_BanksAddSubmit", "Master/Bank Detail/Bank", "Bank Master Add", "N", "BankMaster", "BankMaster", "Add", 2);
            comfun.saveformname("_BankEditSubmit", "/TimeSheet/_BankEditSubmit", "Master/Bank Detail/Bank", "Bank Master Edit", "N", "BankMaster", "BankMaster", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@BankName", Request.Form["BankName"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                // list.Add("@")
                mes = comfun.executeNonQueryWMessage("Banks_Save", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _BanksAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        public JsonResult _BankStatusUpdate()
        {
            comfun.saveformname("_BankStatusUpdate", "/TimeSheet/_BankStatusUpdate", "Master/Bank Detail/Bank", "Bank Master Status", "N", "BankMaster", "BankMaster", "Status", 5);

            string mes = string.Empty;
            try
            {

                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpBankStatusUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        public ActionResult _BanksEdit()
        {
            // comfun.saveformname("BanksEdit", "/admin/BanksEdit", "Bank Edit", "Bank Edit", "N", "BanksEdit", "Bank", "Edit");

            //string bankid = "0";
            //if (Request.QueryString["id"] == null)
            //{
            //    return RedirectToAction("Banks");
            //}
            // bankid = comfun.decryptString(Request.QueryString["Id"].ToString());
            SortedList list = new SortedList();
            list.Add("@BankId", Request.Form["Id"].ToString());
            DataTable dt = comfun.fillDataTable("Bank_SelectWithId", "", list);
            return PartialView("_BankEdit", dt);
        }
        public JsonResult _BankEditSubmit()
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@BankName", Request.Form["BankName"].ToString());
                list.Add("@BankId", Request.Form["BankId"].ToString());
                mes = comfun.executeNonQueryWMessage("Bank_Edit", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _BankEditSubmit", ex.Message);
            }
            return Json(mes);
        }

        public ActionResult empcategories()
        {

            return View();
        }
        public ActionResult _empcategoriesList()
        {
            DataTable dt = comfun.fillDataTable("EmpCategory_Select", "", null);
            return PartialView("_empcategoriesList", dt);
        }
        public ActionResult _employeecategoryAdd()
        {
            return PartialView("_employeeCategoryAdd");
        }
        public JsonResult _EmpCategoryAddSubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@EmpCategoryName", Request.Form["EmployeeCategoryName"].ToString());
                mes = comfun.executeNonQueryWMessage("EmpCategory_Save", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_EmpCategoryAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult _EmployeeCategoryStatusUpdate()
        {
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpUpdateStatusEmployeeCategory", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        public ActionResult _employeecategoryEdit()
        {
            string empcategoryid = "0";
            //if (Request.QueryString["id"] == null)
            //{
            //    return RedirectToAction("empcategories");
            //}
            empcategoryid = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@EmpCategoryId", empcategoryid);
            DataTable dt = comfun.fillDataTable("EmpCategory_SelectWithId", "", list);
            return PartialView("_employeecategoryEdit", dt);
        }
        public JsonResult _EmployeeCategoryEditSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();

                list.Add("@EmpCategoryName", Request.Form["EmployeeCategoryName"].ToString());
                list.Add("@EmpCategoryId", Request.Form["EmployeeCategoryId"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("EmpCategory_Edit", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_EmployeeCategoryEditSubmit", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult Level()
        {
            comfun.saveformname("Level", "/TimeSheet/Level", "Master/General/Level", "Level  Main form", "Y", "Level", "Level", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("Level");
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            //DataTable dt = comfun.fillDataTable("Country_Select", "", null);
            return View();
        }

        public ActionResult _LevelList()
        {
            comfun.saveformname("_LevelList", "/TimeSheet/_LevelList", "Master/General/Level", "Level  list", "N", "Level", "Level", "List", 4);


            SortedList list = new SortedList();

            DataTable dt = comfun.fillDataTable("stpHtisLevelsList", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public JsonResult _LevelSubmit()
        {
            comfun.saveformname("_LevelSubmit", "/TimeSheet/_LevelSubmit", "Master/General/Level", "Level  add", "N", "Level", "Level", "Add", 2);
            comfun.saveformname("_LevelSubmit", "/TimeSheet/_LevelSubmit", "Master/General/Level", "Level  edit", "N", "Level", "Level", "Edit", 3);

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@Title", Request.Form["Title"].ToString());
                list.Add("@ScheduleTime", "0");
                list.Add("@LeveCode", Request.Form["LevelCode"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                list.Add("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                mes = comfun.executeNonQueryWMessage("stpLevel_Save", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Add", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult _LevelStatusUpdate()
        {
            comfun.saveformname("_LevelStatusUpdate", "/TimeSheet/_LevelStatusUpdate", "Master/General/Level", "Level  Status", "N", "Level", "Level", "Status", 4);

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpUpdateStatusLevel", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        //public ActionResult department()
        //{
        //    comfun.saveformname("department", "/TimeSheet/department", "Master/General/Department", "Department Main form", "Y", "Department", "Department", "View", 1);
        //    ViewData["AccessRights"] = payfun.getAccessRights("department");

        //    return View();
        //}
        //public ActionResult _departmentList()
        //{
        //    // comfun.saveformname("_departmentList", "/admin/_departmentList", "Master/General/Department", "Department  List", "N", "Department", "Department", "List");
        //    comfun.saveformname("_departmentList", "/TimeSheet/_departmentList", "Master/General/Department", "Department List", "N", "Department", "Department", "List", 4);
        //    decimal total_records = 0;
        //    decimal pageno = 1;
        //    decimal pagesize = 50;
        //    if (Request.Form["Page"] != null)
        //    {
        //        pagesize = Convert.ToDecimal(Request.Form["Page"]);
        //    }
        //    if (Request.Form["PageNo"] != null)
        //    {
        //        pageno = Convert.ToDecimal(Request.Form["PageNo"]);
        //    }
        //    SortedList list = new SortedList();
        //    list.Add("@PageNo", pageno);
        //    list.Add("@PageSize", pagesize);
        //    DataTable dt = comfun.fillDataTable("Department_Select", "", list);
        //    if (dt.Rows.Count > 0)
        //        total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

        //    string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "admin", "_departmentList", "_departmentList");

        //    HtmlString htm = new HtmlString(paging);
        //    ViewData["paging"] = htm;
        //    var json = JsonConvert.SerializeObject(dt);
        //    return Json(json);
        //}
        ////  [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        //public ActionResult _departmentAdd()
        //{

        //    //if (payfun.sessionRecreate() == "expires")
        //    //{
        //    //    return RedirectToAction("Login", "account");
        //    //}
        //    //comfun.saveformname("departmentAdd", "/admin/departmentAdd", "Department Add", "Department add view", "N", "departmentAdd", "Department", "Add");
        //    return PartialView("_departmentAdd");
        //}
        //// [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        //public ActionResult _departmentEdit()
        //{
        //    //if (payfun.sessionRecreate() == "expires")
        //    //{
        //    //    return RedirectToAction("Login", "account");
        //    //}
        //    //comfun.saveformname("departmentEdit", "/admin/departmentEdit", "Department Edit", "Department edit view", "N", "departmentEdit", "Department", "Edit");
        //    string departmentId = "0";
        //    //if (Request.QueryString["id"] == null)
        //    //{
        //    //    return RedirectToAction("department");
        //    //}
        //    departmentId = Request.Form["Id"].ToString();
        //    SortedList list = new SortedList();
        //    list.Add("@DepartmentId", departmentId);

        //    DataTable dt = comfun.fillDataTable("Department_SelectWithId", "", list);
        //    var json1 = JsonConvert.SerializeObject(dt);
        //    return Json(json1);

        //}
        //public JsonResult _departmentAddSubmit()
        //{
        //    string mes = string.Empty;
        //    try
        //    {
        //        comfun.saveformname("_departmentEditSubmit", "/TimeSheet/_departmentEditSubmit", "Master/General/Department", "Department Edit", "N", "Department", "Department", "Edit", 3);
        //        comfun.saveformname("_departmentAddSubmit", "/TimeSheet/_departmentAddSubmit", "Master/General/Department", "Department Add", "N", "Department", "Department", "Add", 2);
        //        //comfun.saveformname("_departmentAddSubmit", "/admin/_departmentAddSubmit", "Master/General/Department", "Department  Add", "N", "Department", "Department", "Add");
        //        SortedList list = new SortedList();
        //        list.Add("@DepartmentName", Request.Form["DepartmentName"].ToString());
        //        list.Add("@DepartmentCode", Request.Form["DepartmentCode"].ToString());
        //        list.Add("@FunnelId", Request.Form["Funnel"].ToString());
        //        list.Add("@EmpID", Request.Form["EmpId"].ToString());
        //        list.Add("@CreatedBy", "1");
        //        list.Add("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
        //        mes = comfun.executeNonQueryWMessage("Department_save", "", list).ToString();
        //    }
        //    catch (Exception ex)
        //    {
        //        mes = comfun.errorMessage("Add", ex.Message);
        //    }
        //    return Json(mes);
        //}
        //public JsonResult _UpdateStatusDepartment()
        //{
        //    comfun.saveformname("_UpdateStatusDepartment", "/TimeSheet/_UpdateStatusDepartment", "Master/General/Department", "Department Status", "N", "Department", "Department", "Status", 5);

        //    string mes = string.Empty;
        //    SortedList list = new SortedList();
        //    try
        //    {
        //        list.Add("@Id", Request.Form["Id"].ToString());
        //        list.Add("@IsActive", Request.Form["IsActive"].ToString());
        //       mes= comfun.executeNonQueryWMessage("stpUpdateStatusDepartment", "", list).ToString();
        //    }
        //    catch (Exception ex)
        //    {

        //        mes = ex.Message;
        //    }

        //    return Json(mes);
        //}
        //public JsonResult _departmentEditSubmit()
        //{
        //   // comfun.saveformname("_departmentEditSubmit", "/admin/_departmentEditSubmit", "Master/General/Department", "Department  Add", "N", "Department", "Department", "Edit");
        //    string mes = string.Empty;
        //    try
        //    {

        //        SortedList list = new SortedList();
        //        list.Add("@DepartmentName", Request.Form["DepartmentName"].ToString());
        //        list.Add("@DepartmentCode", Request.Form["DepartmentCode"].ToString());
        //        list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
        //        list.Add("@FunnelId", Request.Form["Funnel"].ToString());
        //        list.Add("@EmpID", Request.Form["EmpId"].ToString());
        //        mes = comfun.executeNonQueryWMessage("Department_Edit", "", list).ToString();
        //    }
        //    catch (Exception ex)
        //    {
        //        mes = comfun.errorMessage("Edit", ex.Message);
        //    }
        //    return Json(mes);
        //}
        public ActionResult Designation()
        {
            comfun.saveformname("Designation", "/TimeSheet/Designation", "Master/General/Designation", "Designation Main form", "Y", "Designation", "Designation", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("Designation");
            return View();
        }
        public ActionResult _designationList()
        {
            comfun.saveformname("_designationList", "/TimeSheet/_designationList", "Master/General/Designation", "Designation List", "N", "Designation", "Designation", "List", 4);
            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 50;
            if (Request.Form["Page"] != null)
            {
                pagesize = Convert.ToDecimal(Request.Form["Page"]);
            }
            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }
            SortedList list = new SortedList();
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            DataTable dt = comfun.fillDataTable("Designation_Select", "", list);
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "admin", "_designationList", "_designationList");

            HtmlString htm = new HtmlString(paging);
            ViewData["paging"] = htm;
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        //[customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult _DesignationAdd()
        {

            // comfun.saveformname("DesignationAdd", "/admin/DesignationAdd", "Designation Add", "Designation add view", "N", "DesignationAdd", "Designation", "Add");
            DataTable dt = comfun.fillDataTable("stpViksatPayGradeSelect", "", null);
            return PartialView("_DesignationAdd", dt);
        }
        public ActionResult _DesignationStatusUpdate()
        {
            comfun.saveformname("_DesignationStatusUpdate", "/TimeSheet/_DesignationStatusUpdate", "Master/General/Designation", "Designation Status", "N", "Designation", "Designation", "Status", 5);

            // comfun.saveformname("DesignationAdd", "/admin/DesignationAdd", "Designation Add", "Designation add view", "N", "DesignationAdd", "Designation", "Add");

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpUpdateStatusDesignation", "", list).ToString();
            }
            catch (Exception ex)
            {

                mes = ex.Message;
            }

            return Json(mes);
        }
        public JsonResult _designationAddSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            comfun.saveformname("_designationEditSubmit", "/TimeSheet/_designationEditSubmit", "Master/General/Designation", "Designation Edit", "N", "Designation", "Designation", "Edit", 3);
            comfun.saveformname("_designationAddSubmit", "/TimeSheet/_designationAddSubmit", "Master/General/Designation", "Designation Add", "N", "Designation", "Designation", "Add", 2);
            try
            {
                SortedList list = new SortedList();

                //  list.Add("@fiGradeId", Request.Form["GradeID"].ToString());
                list.Add("@DesignationCode", Request.Form["DesignationCode"].ToString());
                list.Add("@DesignationName", Request.Form["DesignationName"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                list.Add("@CreatedDate", comfun.dateISTstr());
                mes = comfun.executeNonQueryWMessage("Designation_Save", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_designationAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        // [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult _designationEdit()
        {

            //comfun.saveformname("designationEdit", "/admin/designationEdit", "Designation Edit", "Designation Edit view", "N", "DesignationEdit", "Designation", "Edit");
            string designationid = "0";
            //if (Request.QueryString["id"] == null)
            //{
            //    return RedirectToAction("Designation");
            //}
            designationid = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@DesignationId", designationid);
            DataSet ds = comfun.fillDataSet("Designation_SelectWithId", "", list);
            return PartialView(ds);
        }
        public JsonResult _designationEditSubmit()
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}
            //comfun.saveformname("_designationEditSubmit", "/admin/_designationEditSubmit", "General/Designation", "Designation Edit", "N", "Designation", "Designation", "Edit");
            comfun.saveformname("_designationEditSubmit", "/TimeSheet/_designationEditSubmit", "Master/General/Designation", "Designation Edit", "N", "Designation", "Designation", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                //list.Add("@GradeID", Request.Form["GradeID"].ToString());
                list.Add("@DesignationName", Request.Form["DesignationName"].ToString());
                list.Add("@fvDesignationCode", Request.Form["DesignationCode"].ToString());
                list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
                //list.Add("@Isactive", Request.Form["Isactive"].ToString());
                mes = comfun.executeNonQueryWMessage("Designation_Edit", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_designationEditSubmit", ex.Message);
            }
            return Json(mes);
        }


        public ActionResult CompanyNatureList()
        {
            comfun.saveformname("CompanyNature", "/TimeSheet/CompanyNatureList", "Master/Organization/Company Nature", "Company Nature Main form", "Y", "CompanyNature", "CompanyNature", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("CompanyNature");
            //comfun.saveformname("Country", "/admin/Country", "Country", "Country  form", "Y", "", "Country", "List");
            //if (Request.QueryString["key"] != null)
            //{
            //    Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            //}

            return View();
        }
        public ActionResult _companyNatureList()
        {
            comfun.saveformname("_companyNatureList", "/TimeSheet/_companyNatureList", "Master/Organization/Company Nature", "Company Nature List", "N", "CompanyNature", "CompanyNature", "List", 4);
            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 50;
            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }
            SortedList list = new SortedList();
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            DataTable dt = comfun.fillDataTable("sp_CompanyNatureListDisplay", "", list);
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);
            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "admin", "_companyNatureList", "_companyNatureList");
            HtmlString htm = new HtmlString(paging);
            ViewData["paging"] = htm;
            //return PartialView(dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        //[customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        //public ActionResult _CompanyNatureAdd()
        //{// comfun.saveformname("CountriesAdd", "/admin/CountriesAdd", "Country Add", "Country view", "N", "Country Add", "Country", "Add");
        //    return PartialView("_companyNatureAdd");
        //}
        //public JsonResult _CompanyNatureAddSubmit()
        //{
        //    //if (payfun.sessionRecreate() == "expires")
        //    //{
        //    //    return Json("Session expires");
        //    //}
        //    string mes = string.Empty;
        //    try
        //    {

        //        SortedList list = new SortedList();
        //        list.Add("@CompanyNature", Request.Form["CompanyNature"].ToString());
        //        mes = comfun.executeNonQueryWMessage("sp_CompanyNature_Save", "", list).ToString();

        //    }
        //    catch (Exception ex)
        //    {
        //        mes = comfun.errorMessage("_CompnayNatureAddSubmit", ex.Message);
        //    }
        //    return Json(mes);




        //}
        public JsonResult _CompnayNatureStatusUpdate()
        {



            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpCountryActivationStatusUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }
        public ActionResult _CompanyNatureEdit()
        {
            //if (Request.QueryString["id"] == null)
            //{
            //    return RedirectToAction("Countries");
            //}
            // countryid = comfun.decryptString(Request.QueryString["Id"].ToString());
            SortedList list = new SortedList();
            list.Add("@CompanyNatureId", Request.Form["Ids"].ToString());
            DataTable dt = comfun.fillDataTable("sp_CompanyNatureListDisplayViaId", "", list);
            return PartialView("_companyNatureEdit", dt);

        }
        public JsonResult _CompnayNatureEditSubmit()
        {
            //{
            //    if (payfun.sessionRecreate() == "expires")
            //    {
            //        return Json("Session expires");
            //    }
            comfun.saveformname("_CompnayNatureEditSubmit", "/TimeSheet/_CompnayNatureEditSubmit", "Master/Organization/Company Nature", "Company Nature Add", "N", "CompanyNature", "CompanyNature", "Add", 2);
            comfun.saveformname("_CompnayNatureEditSubmit", "/TimeSheet/_CompnayNatureEditSubmit", "Master/Organization/Company Nature", "Company Nature Edit", "N", "CompanyNature", "CompanyNature", "Edit", 3);
            string mes = string.Empty;
            try
            {

                SortedList list = new SortedList();
                list.Add("@CompanyNatureId", Request.Form["CompanyNatureId"].ToString());
                list.Add("@CompanyNature", Request.Form["CompanyNature"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_CompanyNature_SaveUpdate", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }


        public ActionResult ConveyanceSettingList()
        {
            comfun.saveformname("ConveyanceSettingList", "/TimeSheet/ConveyanceSettingList", "Master/Conveyance&nbsp;/Auto setting", "Conveyance Setting form", "Y", "ConveyanceSetting", "ConveyanceSetting", "View", 1);
            //  ViewData["AccessRights"] = payfun.getAccessRights("ConveyanceSettingList");
            return View();
        }
        public ActionResult _conveyanceSettingList()
        {
            comfun.saveformname("ConveyanceSettingList", "/TimeSheet/_conveyanceSettingList", "Master/Conveyance&nbsp;/Auto setting", "Conveyance Setting List", "N", "ConveyanceSetting", "ConveyanceSetting", "List", 4);

            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 50;
            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }
            SortedList list = new SortedList();
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            DataTable dt = comfun.fillDataTable("sp_CASListDisplay", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _ConveyanceSettingAdd()
        {// comfun.saveformname("CountriesAdd", "/admin/CountriesAdd", "Country Add", "Country view", "N", "Country Add", "Country", "Add");
            return PartialView("_conveyanceSettingAdd");
        }
        public JsonResult _ConveyanceSettingAddSubmit()
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}
            comfun.saveformname("ConveyanceSettingList", "/TimeSheet/_ConveyanceSettingAddSubmit", "Master/Conveyance&nbsp;/Auto setting", "Conveyance Setting Add", "N", "ConveyanceSetting", "ConveyanceSetting", "Add", 2);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@ConveyanceAutoSettingId", Request.Form["CASId"].ToString());
                list.Add("@FromTime", Request.Form["FromTime"].ToString());
                list.Add("@ToTime", Request.Form["ToTime"].ToString());
                list.Add("@Interval", Request.Form["Interval"].ToString());
                list.Add("@DeviceInterval", Request.Form["DeviceInterval"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_CAS_SaveUpdate", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_CompnayNatureAddSubmit", ex.Message);
            }
            return Json(mes);





        }
        public JsonResult _ConveyanceSettingStatusUpdate()
        {
            comfun.saveformname("_ConveyanceSettingStatusUpdate", "/TimeSheet/_ConveyanceSettingStatusUpdate", "Master/Conveyance&nbsp;/Auto setting", "Conveyance Setting Status", "N", "ConveyanceSetting", "ConveyanceSetting", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_CAS_Active_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }
        public ActionResult _ConveyanceSettingEdit()
        {
            //if (Request.QueryString["id"] == null)
            //{
            //    return RedirectToAction("Countries");
            //}
            // countryid = comfun.decryptString(Request.QueryString["Id"].ToString());
            SortedList list = new SortedList();
            string cgid = "0";
            cgid = Request.QueryString["Id"].ToString();
            list.Add("@CASId", cgid);
            DataTable dt = comfun.fillDataTable("sp_CASDisplayViaId", "", list);
            return PartialView("_ConveyanceSettingEdit", dt);

        }
        public JsonResult _ConveyanceSettingEditSubmit()
        {
            //{
            //    if (payfun.sessionRecreate() == "expires")
            //    {
            //        return Json("Session expires");
            //    }
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@ConveyanceGroupId", Request.Form["ConveyanceGroupId"].ToString());
                list.Add("@ConveyanceGroupCode", Request.Form["ConveyanceGroupCode"].ToString());
                list.Add("@ConveyanceGroupName", Request.Form["ConveyanceGroupName"].ToString());
                list.Add("@PricePerKM", Request.Form["ConveyanceGroupPricePerKM"].ToString());
                list.Add("@MonthlyLimit", Request.Form["ConveyanceGroupMonthlyLimit"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_ConveyanceGroup_SaveUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }


        public ActionResult AddNewConveyanceAutoSetting()
        {
            return View();
        }
        //To Save/Update Conveyance Auto Setting//
        public JsonResult ConveyanceAutoSettingSaveUpdateSubmit()
        {
            comfun.saveformname("ConveyanceSettingList", "/TimeSheet/ConveyanceAutoSettingSaveUpdateSubmit", "Master/Conveyance&nbsp;/Auto setting", "Conveyance Setting Add", "N", "ConveyanceSetting", "ConveyanceSetting", "Add", 2);
            comfun.saveformname("ConveyanceSettingList", "/TimeSheet/ConveyanceAutoSettingSaveUpdateSubmit", "Master/Conveyance&nbsp;/Auto setting", "Conveyance Setting Edit", "N", "ConveyanceSetting", "ConveyanceSetting", "Edit", 3);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@ConveyanceAutoSettingId", Request.Form["CASId"].ToString());
                list.Add("@FromTime", Request.Form["FromTime"].ToString());
                list.Add("@ToTime", Request.Form["ToTime"].ToString());
                list.Add("@Interval", Request.Form["Interval"].ToString());
                list.Add("@DeviceInterval", Request.Form["DeviceInterval"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_CAS_SaveUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        //To Display Conveyance Auto Setting List//
        public ActionResult ConveyanceAutoSettingList()
        {
            DataTable dt = comfun.fillDataTable("sp_CASListDisplay", "", null);
            return View(dt);
        }
        //To Get Conveyance Auto Setting Details Via Id//
        public ActionResult UpdateConveyanceAutoSetting()
        {
            string casid = "0";
            if (Request.QueryString["id"] == null)
            {
                return RedirectToAction("ConveyanceAutoSettingList", "ConveyanceAutoSettingList");
            }
            casid = Request.QueryString["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@CASId", casid);
            DataTable dt = comfun.fillDataTable("sp_CASDisplayViaId", "", list);
            return View(dt);
        }
        //To Change Active/Deactive Conveyance Auto Setting//
        public JsonResult ConveyanceAutoSettingActivationStatusUpdate()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@CASId", Request.Form["CASId"].ToString());
                list.Add("@CASStatus", Request.Form["CASStatus"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_CAS_Active_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }


        public JsonResult _CityStatusUpdate()
        {
            comfun.saveformname("_CityStatusUpdate", "/TimeSheet/_CityStatusUpdate", "Master/Address/City", "City  Status", "N", "City", "City", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpCityStatusUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }

        public ActionResult Branch()
        {
            comfun.saveformname("Branch", "/TimeSheet/Branch", "Master/Organization/Branch&nbsp;", "Branch form", "Y", "Branch", "Branch", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("Branch");
            SortedList list = new SortedList();
            list.Add("@Action", "Add");
            DataSet ds = comfun.fillDataSet("stpViksatBranchCompanyStateCity", "", list);
            return View(ds);
        }
        public JsonResult _branchList()
        {
            comfun.saveformname("_branchList", "/TimeSheet/_branchList", "Master/Organization/Branch&nbsp;", "Branch List", "N", "Branch", "Branch", "List", 4);
            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 50;
            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }
            SortedList list = new SortedList();
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            DataTable dt = comfun.fillDataTable("Branch_Select", "", list);
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "admin", "_branchList", "_branchList");

            HtmlString htm = new HtmlString(paging);
            ViewData["paging"] = htm;
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        // [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult _BranchAdd()
        {
            // comfun.saveformname("BranchAdd", "/admin/BranchAdd", "Branch Add", "Branch add view", "N", "BranchAdd", "Branch", "Add");
            SortedList list = new SortedList();
            list.Add("@Action", "Add");
            DataSet ds = comfun.fillDataSet("stpViksatBranchCompanyStateCity", "", list);
            return PartialView("_CompanyBranchAdd", ds);
        }
        public JsonResult _BranchAddSubmit()
        {
            comfun.saveformname("_BranchAddSubmit", "/TimeSheet/_BranchAddSubmit", "Master/Organization/Branch&nbsp;", "Branch Add", "N", "Branch", "Branch", "Add", 2);
            comfun.saveformname("_BranchAddSubmit", "/TimeSheet/_BranchAddSubmit", "Master/Organization/Branch&nbsp;", "Branch Edit", "N", "Branch", "Branch", "Edit", 3);
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();

                list.Add("@fvBranchName", Request.Form["fvBranchName"].ToString());
                list.Add("@fvBranchCode", Request.Form["BranchCode"].ToString());
                list.Add("@fvAddress1", Request.Form["fvAddress1"].ToString());
                list.Add("@fiStateId", Request.Form["fiStateId"].ToString());
                list.Add("@CreatedBy", Session["EmpId"]);
                list.Add("@CreatedDate", comfun.dateISTstr());
                list.Add("@fiDivisionID", Request.Form["fiDivisionID"].ToString());
                list.Add("@fiCityId", Request.Form["fiCityId"].ToString());
                list.Add("@fvPostalCode", Request.Form["PinCode"].ToString());
                list.Add("@fvPhoneNo", Request.Form["fvPhoneNo"].ToString());
                list.Add("@fvMobileNo", Request.Form["fvMobileNo"].ToString());
                list.Add("@fvFaxNo", Request.Form["Fax"].ToString());
                list.Add("@fvEmail", Request.Form["fvEmail"].ToString());
                list.Add("@fiCountryId", Request.Form["Country"].ToString());
                list.Add("@ESINo", Request.Form["ESINo"].ToString());
                list.Add("@fvPanNo", Request.Form["fvPanNo"].ToString());
                list.Add("@fvPFCode", Request.Form["fvPFCode"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                list.Add("@GSTNo", Request.Form["GSTNo"].ToString());
                list.Add("@fiZoneTimeId", Request.Form["ZoneTime"].ToString());
                list.Add("@fiEmployeeId", Request.Form["Employee"].ToString());
                list.Add("@fvEmail1", Request.Form["Email2"].ToString());
                list.Add("@fvTollFree", Request.Form["TollFree"].ToString());
                list.Add("@fiZoneId", Request.Form["Zone"].ToString());
                list.Add("@fvMobileNo1", Request.Form["MobileNo2"].ToString());
                mes = comfun.executeNonQueryWMessage("Branch_Save", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_BranchAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult _CompanyBranchStatusUpdate()
        {
            comfun.saveformname("_CompanyBranchStatusUpdate", "/TimeSheet/_CompanyBranchStatusUpdate", "Master/Organization/Branch&nbsp;", "Branch Status", "N", "Branch", "Branch", "Status", 5);

            // comfun.saveformname("DesignationAdd", "/admin/DesignationAdd", "Designation Add", "Designation add view", "N", "DesignationAdd", "Designation", "Add");

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpCompanyBranchStatus", "", list).ToString();
            }
            catch (Exception ex)
            {

                mes = ex.Message;
            }

            return Json(mes);
        }
        //    [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public JsonResult _CompanyBranchEdit1()
        {
            //  comfun.saveformname("CompanyEdit", "/admin/CompanyEdit", "Company Edit", "Company Edit", "N", "CompanyEdit", "CompanyEdit", "Edit");

            string branchid = "0";
            //if (Request.QueryString["id"] == null)
            //{
            //    return RedirectToAction("branches");
            //}
            branchid = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@BranchId", branchid);
            list.Add("@Action", "Add");
            DataSet ds = comfun.fillDataSet("Branch_SelectWithId", "", list);

            var json = JsonConvert.SerializeObject(ds.Tables[0]);
            return Json(json);
        }
        public ActionResult _CompanyBranchEdit()
        {
            // comfun.saveformname("BranchEdit", "/admin/BranchEdit", "Branch Edit", "Branch Edit view", "N", "BranchEdit", "Branch", "Edit");

            string branchid = "0";
            //if (Request.QueryString["id"] == null)
            //{
            //    return RedirectToAction("branches");
            //}
            branchid = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@BranchId", branchid);
            list.Add("@Action", "Edit");
            DataSet ds = comfun.fillDataSet("Branch_SelectWithId", "", list);
            return PartialView("_CompanyBranchEdit", ds);
        }
        public JsonResult _BranchEditSubmit()
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@fiBranchId", Request.Form["BranchId"].ToString());
                list.Add("@fvBranchName", Request.Form["fvBranchName"].ToString());
                list.Add("@fvBranchCode", Request.Form["BranchCode"].ToString());
                list.Add("@fvAddress1", Request.Form["fvAddress1"].ToString());
                list.Add("@fiStateId", Request.Form["fiStateId"].ToString());
                list.Add("@fiDivisionID", Request.Form["fiDivisionID"].ToString());
                list.Add("@fiCityId", Request.Form["fiCityId"].ToString());
                list.Add("@fvPhoneNo", Request.Form["fvPhoneNo"].ToString());
                list.Add("@fvMobileNo", Request.Form["fvMobileNo"].ToString());
                list.Add("@fvEmail", Request.Form["fvEmail"].ToString());
                list.Add("@fvPanNo", Request.Form["fvPanNo"].ToString());
                list.Add("@fvPFCode", Request.Form["fvPFCode"].ToString());
                //list.Add("@fvBranchName", Request.Form["BranchName"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                list.Add("@GSTNo", Request.Form["GSTNo"].ToString());
                list.Add("@fiCountryId ", Request.Form["Country"].ToString());
                list.Add("@fvPostalCode", Request.Form["PinCode"].ToString());
                list.Add("@fiZoneTimeId", Request.Form["ZoneTime"].ToString());
                list.Add("@fiEmployeeId", Request.Form["Employee"].ToString());
                list.Add("@fvEmail1    ", Request.Form["Email2"].ToString());
                list.Add("@fvTollFree  ", Request.Form["TollFree"].ToString());
                list.Add("@fiZoneId    ", Request.Form["Zone"].ToString());
                list.Add("@fvMobileNo1 ", Request.Form["MobileNo2"].ToString());
                list.Add("@fvFaxNo", Request.Form["Fax"].ToString());

                mes = comfun.executeNonQueryWMessage("Branch_Edit", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_BranchEditSubmit", ex.Message);
            }
            return Json(mes);
        }
        //public ActionResult _Employeeddl()
        //{
        //    DataTable dt = comfun.fillDataTable("stpEmployeeddl", "", null);
        //    var json1 = JsonConvert.SerializeObject(dt);
        //    return Json(json1);
        //}
        public ActionResult DivisionList()
        {
            //if (Request.QueryString["key"] != null)
            //{
            //    Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            //}
            ////Division renamed to company
            //    comfun.saveformname("CompanyList", "/admin/CompanyList", "Company List", "Company Main form", "Y", "", "Company", "List");
            DataSet ds = comfun.fillDataSet("stpViksatDivisionStateCity", "", null);
            return View(ds);
        }
        public ActionResult _DivisionList()
        {
            comfun.saveformname("_DivisionList", "/TimeSheet/_DivisionList", "Master/Organization/Company List", "Company Main form", "Y", "CompanyList", "CompanyList", "List", 4);
            SortedList list = new SortedList();
            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 50;

            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }
            if (Request.Form["PageNo"] != null)
            {
                pagesize = Convert.ToDecimal(Request.Form["Page"]);
            }

            //paging_get_downline1
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);

            DataTable dt = comfun.fillDataTable("stpviksatDivisionSelect", "", null);
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "admin", "_DivisionList", "_DivisionList");

            HtmlString htm = new HtmlString(paging);
            ViewData["paging"] = htm;
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
            //return PartialView(dt);

        }
        // [customAuthorize(Roles = _roles)]
        public ActionResult _DivisionAdd()
        {
            //   comfun.saveformname("CompanyAdd", "/admin/CompanyAdd", "Company Add", "Company Add", "N", "CompanyAdd", "CompanyAdd", "Add");
            DataSet ds = comfun.fillDataSet("stpViksatDivisionStateCity", "", null);
            return PartialView("_DivisionAdd", ds);
        }
        public JsonResult _DivisionAddSubmit()
        {
            comfun.saveformname("_DivisionAddSubmit", "/TimeSheet/_DivisionAddSubmit", "Master/Organization/Company List", "Company Add", "N", "CompanyList", "CompanyList", "Add", 2);
            comfun.saveformname("_DivisionAddSubmit", "/TimeSheet/_DivisionAddSubmit", "Master/Organization/Company List", "Company Edit", "N", "CompanyList", "CompanyList", "Edit", 3);
            //  comfun.saveformname("_DivisionAddSubmit", "/admin/_DivisionAddSubmit", "Division Add Submit", "Division Add Submit", "N", "DivisionAdd", "Division", "");
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {

                list.Add("@fvDivisionName", Request.Form["fvDivisionName"].ToString());
                list.Add("@fvAddress1", Request.Form["fvAddress1"].ToString());
                list.Add("@fiStateId", Request.Form["fiStateId"].ToString());
                list.Add("@fiCityId", Request.Form["fiCityId"].ToString());
                list.Add("@fvEmail", Request.Form["fvEmail"].ToString());




                list.Add("@fvPanNo", Request.Form["fvPanNo"].ToString());
                list.Add("@fvPFCode", Request.Form["fvPFCode"].ToString());

                list.Add("@GSTNO", Request.Form["GSTNo"].ToString());
                list.Add("@fvPhoneNo", Request.Form["fvPhoneNo"].ToString());
                list.Add("@fvMobileNo", Request.Form["fvMobileNo"].ToString());
                list.Add("@IsActive", "Y");
                list.Add("@fvWebSite", Request.Form["fvWebSite"].ToString());
                list.Add("@companyNatureId", Request.Form["CompanyNature"].ToString());
                list.Add("@Email2", Request.Form["Email2"].ToString());
                list.Add("@Mobile2", Request.Form["Mobile2"].ToString());
                list.Add("@TollFree", Request.Form["TollFree"].ToString());
                list.Add("@CountryId", Request.Form["CountryId"].ToString());
                list.Add("@FaxNo", Request.Form["FaxNo"].ToString());
                list.Add("@PrefixName", Request.Form["PrefixName"].ToString());
                if (Request.Form["DivisionID"].ToString() == "0")
                {
                    list.Add("@fvDivisionCode", Request.Form["fvDivisionCode"].ToString());
                    list.Add("@CreatedBy", "1");
                    list.Add("@CreatedDate", comfun.dateISTstr());
                    mes = comfun.executeNonQueryWMessage("stpviksatDivisionSave", "", list).ToString();
                }
                else
                {
                    list.Add("@fvCompanyCode", Request.Form["fvDivisionCode"].ToString());
                    list.Add("@fiDivisionId", Request.Form["DivisionID"].ToString());
                    mes = comfun.executeNonQueryWMessage("stpviksatDivisionUpdate", "", list).ToString();
                }
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _DivisionAddSubmit", ex.Message);
            }

            return Json(mes);
        }
        public JsonResult _CompanyStatusUpdate()
        {
            comfun.saveformname("_CompanyStatusUpdate", "/TimeSheet/_CompanyStatusUpdate", "Master/Organization/Company Status", "Company Status", "N", "CompanyList", "CompanyList", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpDivisionStatusUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);



        }
        public JsonResult _DivisionEdit1()
        {
            //  comfun.saveformname("CompanyEdit", "/admin/CompanyEdit", "Company Edit", "Company Edit", "N", "CompanyEdit", "CompanyEdit", "Edit");

            string DivisionId = "0";

            DivisionId = Request.Form["Id"].ToString();

            SortedList list = new SortedList();

            list.Add("@DivisionID", DivisionId);
            DataSet ds = comfun.fillDataSet("stpviksatDivisionSelectWithID", "", list);
            var json = JsonConvert.SerializeObject(ds.Tables[0]);
            return Json(json);
        }

        //   [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult _DivisionEdit()
        {
            //  comfun.saveformname("CompanyEdit", "/admin/CompanyEdit", "Company Edit", "Company Edit", "N", "CompanyEdit", "CompanyEdit", "Edit");

            string DivisionId = "0";
            //if (Request.QueryString["id"] == null)
            //{
            //    return RedirectToAction("CompanyList");
            //}
            DivisionId = Request.Form["Id"].ToString();

            SortedList list = new SortedList();

            list.Add("@DivisionID", DivisionId);
            DataSet dt = comfun.fillDataSet("stpviksatDivisionSelectWithID", "", list);
            return PartialView("_DivisionEdit", dt);
        }
        public JsonResult _DivisionEditSubmit()
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@fiDivisionId", Request.Form["DivisionID"].ToString());
                list.Add("@fvDivisionName", Request.Form["Division"].ToString());
                list.Add("@fvAddress1", Request.Form["Address"].ToString());
                list.Add("@fiStateId", Request.Form["State"].ToString());
                list.Add("@fiCityId", Request.Form["City"].ToString());
                list.Add("@fvEmail", Request.Form["Email"].ToString());
                list.Add("@fvPanNo", Request.Form["PanNo"].ToString());
                list.Add("@fvPFCode", Request.Form["PFCode"].ToString());
                list.Add("@fvCompanyCode", Request.Form["DivisionCode"].ToString());
                list.Add("@IsActive", Request.Form["Isactive"].ToString());
                list.Add("@fvPhoneNo", Request.Form["PhoneNo"].ToString());
                list.Add("@fvMobileNo", Request.Form["MobileNo"].ToString());
                list.Add("@GSTNO", Request.Form["GSTNO"].ToString());
                list.Add("@fvWebSite", Request.Form["Website"].ToString());
                list.Add("@companyNatureId", Request.Form["CompanyNature"].ToString());
                list.Add("@Email2", Request.Form["Email2"].ToString());
                list.Add("@Mobile2", Request.Form["Mobile2"].ToString());
                list.Add("@TollFree", Request.Form["TollFree"].ToString());
                list.Add("@CountryId", Request.Form["CountryId"].ToString());
                list.Add("@FaxNo", Request.Form["FaxNo"].ToString());
                list.Add("@PrefixName", Request.Form["PrefixName"].ToString());
                mes = comfun.executeNonQueryWMessage("stpviksatDivisionUpdate", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _DivisionEditSubmit", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult UnitList()
        {
            return View();
        }

        public ActionResult _UnitList()
        {
            SortedList list = new SortedList();
            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 50;

            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }

            //paging_get_downline1
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);

            DataTable dt = comfun.fillDataTable("StpViksatUnit_Select", "", list);
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "admin", "_UnitList", "_UnitList");

            HtmlString htm = new HtmlString(paging);
            ViewData["paging"] = htm;
            return PartialView("_UnitList", dt);
        }
        public JsonResult _UnitStatusUpdate()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpUnitStatusUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);



        }
        public ActionResult _UnitAdd()
        {
            // comfun.saveformname("UnitAdd", "/admin/UnitAdd", "Unit Add", "Unit view", "N", "UnitAdd", "Unit", "Add");

            DataSet ds = comfun.fillDataSet("stpViksatBranchCompanyStateCity", "", null);
            return PartialView("_UnitAdd", ds);
        }
        public JsonResult _UnitAddSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                list.Add("@CreatedDate", comfun.dateISTstr());
                list.Add("@fvUnitName", Request.Form["fvBranchName"].ToString());
                list.Add("@fvAddress1", Request.Form["fvAddress1"].ToString());
                list.Add("@fiStateId", Request.Form["fiStateId"].ToString());
                list.Add("@fiCompanyID", Request.Form["fiDivisionID"].ToString());
                list.Add("@fiCityId", Request.Form["fiCityId"].ToString());
                //  list.Add("@fvPostalCode",Request.Form["PostalCode"].ToString());
                list.Add("@fvPhoneNo", Request.Form["fvPhoneNo"].ToString());
                list.Add("@fvMobileNo", Request.Form["fvMobileNo"].ToString());
                // list.Add("@fvFaxNo",Request.Form["FaxNo"].ToString());
                list.Add("@fvEmail", Request.Form["fvEmail"].ToString());
                list.Add("@fvUnitCode", Request.Form["fvUnitCode"].ToString());
                list.Add("@ESINo", Request.Form["ESINo"].ToString());
                list.Add("@fvPanNo", Request.Form["fvPanNo"].ToString());
                list.Add("@fvPFCode", Request.Form["fvPFCode"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                list.Add("@GSTNo", Request.Form["GSTNo"].ToString());
                mes = comfun.executeNonQueryWMessage("stpViksatUnit_Save", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _UnitAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        //[customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult _UnitEdit()
        {
            // comfun.saveformname("UnitEdit", "/admin/UnitEdit", "Unit Edit", "Unit Edit", "N", "UnitEdit", "Unit", "Edit");

            string UnitId = "0";
            //if (Request.QueryString["id"] == null)
            //{
            //    return RedirectToAction("UnitList");
            //}
            UnitId = Request.Form["Id"].ToString();

            SortedList list = new SortedList();

            list.Add("@UnitId", UnitId);
            DataSet dt = comfun.fillDataSet("stpViksatUnit_SelectWithId", "", list);
            return View("_UnitEdit", dt);
        }
        public JsonResult _UnitEditSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@fiUnitId", Request.Form["UnitId"].ToString());
                list.Add("@fvUnitName", Request.Form["UnitName"].ToString());
                list.Add("@fvUnitCode", Request.Form["UnitCode"].ToString());
                list.Add("@fvAddress1", Request.Form["Address"].ToString());
                list.Add("@fiStateId", Request.Form["State"].ToString());
                list.Add("@fiCompanyId", Request.Form["CompanyID"].ToString());
                list.Add("@fiCityId", Request.Form["City"].ToString());
                list.Add("@fvPhoneNo", Request.Form["PhoneNo"].ToString());
                list.Add("@fvMobileNo", Request.Form["MobileNo"].ToString());
                list.Add("@fvEmail", Request.Form["Email"].ToString());
                list.Add("@fvPanNo", Request.Form["PanNo"].ToString());
                list.Add("@fvPFCode", Request.Form["PFCode"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                list.Add("@GSTNo", Request.Form["GSTNo"].ToString());
                mes = comfun.executeNonQueryWMessage("stpviksatUnit_Edit", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _UnitEditSubmit", ex.Message);
            }

            return Json(mes);
        }


        //[customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult GradeList()
        {
            comfun.saveformname("GradeList", "/TimeSheet/GradeList", "Master/General/GradeList", "Grade List form", "Y", "GradeList", "GradeList", "View", 1);
            // ViewData["AccessRights"] = payfun.getAccessRights("GradeList");
            return View();
        }

        public ActionResult _GradeList()
        {
            comfun.saveformname("_GradeList", "/TimeSheet/_GradeList", "Master/General/GradeList", "Grade List List", "N", "GradeList", "GradeList", "List", 4);
            SortedList list = new SortedList();
            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 50;

            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }

            //paging_get_downline1
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            DataTable dt = comfun.fillDataTable("stpViksatPayGradeSelect", "", list);
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "admin", "_GradeList", "_GradeList");

            HtmlString htm = new HtmlString(paging);
            ViewData["paging"] = htm;
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        //[customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult _GradeAdd()
        {
            return PartialView("_GradeAdd");
        }
        public JsonResult _GradesAddSubmit()
        {
            comfun.saveformname("_GradesAddSubmit", "/TimeSheet/_GradesAddSubmit", "Master/General/GradeList", "Grade List Add", "N", "GradeList", "GradeList", "Add", 2);
            comfun.saveformname("_GradeEditSubmit", "/TimeSheet/_GradeEditSubmit", "Master/General/GradeList", "Grade List Edit", "N", "GradeList", "GradeList", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();

                list.Add("@Grade", Request.Form["Grade"].ToString());
                list.Add("@GradeCode", Request.Form["GradeCode"].ToString());
                list.Add("@IsActive", "Y");
                list.Add("@CreatedBy", "1");
                list.Add("@CreatedDate", comfun.dateISTstr());
                mes = comfun.executeNonQueryWMessage("stpViksatpayGradeinsert", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = "Error: " + comfun.errorMessage("Admin _GradesAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult _GradeEdit()
        {

            string GradeID = "0";
            GradeID = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@GradeID", GradeID);
            DataTable dt = comfun.fillDataTable("stpViksatPayGradeSelectWithGradeID", "", list);
            return PartialView(dt);
        }
        public JsonResult _GradeEditSubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@GradeID", Request.Form["GradeId"].ToString());
                list.Add("@GradeName", Request.Form["Grade"].ToString());
                list.Add("@GradeCode", Request.Form["GradeCode"].ToString());
                list.Add("@IsActive", 'Y');
                mes = comfun.executeNonQueryWMessage("stpViksatpayGradeupdate", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = "Error: " + comfun.errorMessage("Admin _GradeEditSubmit", ex.Message);
            }
            return Json(mes);
        }
        public JsonResult _GradeStatusUpdate()
        {
            comfun.saveformname("_GradeStatusUpdate", "/TimeSheet/_GradeStatusUpdate", "Master/General/GradeList", "Grade List Status", "Y", "GradeList", "GradeList", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpGradeStatusUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
            }
            return Json(mes);



        }

        //To Display Employee Type List//
        public ActionResult EmptypeList()
        {
            comfun.saveformname("EmployeeType", "/TimeSheet/EmptypeList", "Master/General/Employee Type", "Employee Type form", "Y", "EmployeeType", "EmployeeType", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("EmployeeType");
            return View();
        }
        public ActionResult _EmptypeList()
        {
            comfun.saveformname("_EmptypeList", "/TimeSheet/_EmptypeList", "Master/General/Employee Type", "Employee Type List", "N", "EmployeeType", "EmployeeType", "List", 4);
            SortedList list = new SortedList();
            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 50;

            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }

            //paging_get_downline1
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            DataTable dt = comfun.fillDataTable("sp_EmpTypeListDisplay", "", list);
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "admin", "_EmployeetypeList", "_EmployeetypeList");

            HtmlString htm = new HtmlString(paging);
            ViewData["paging"] = htm;
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
            //return PartialView("_EmployeetypeList", dt);
        }
        public ActionResult _EmptypeAdd()
        {
            return PartialView("_EmptypeAdd");
        }
        public JsonResult _EmptypeAddSubmit()
        {
            comfun.saveformname("_EmptypeAddSubmit", "/TimeSheet/_EmptypeAddSubmit", "Master/General/Employee Type", "Employee Type Add", "N", "EmployeeType", "EmployeeType", "Add", 2);
            comfun.saveformname("_EmptypeEditSubmit", "/TimeSheet/_EmptypeEditSubmit", "Master/General/Employee Type", "Employee Type Edit", "N", "EmployeeType", "EmployeeType", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@EmpTypeId", Request.Form["EmpTypeId"].ToString());
                list.Add("@EmpTypeCode", Request.Form["EmpTypeCode"].ToString());
                list.Add("@EmpTypeName", Request.Form["EmpTypeName"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_EmpType_SaveUpdate", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _GradesAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult _EmptypeEdit()
        {

            string emptypeid = "0";
            emptypeid = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@EmpTypeId", emptypeid);
            DataTable dt = comfun.fillDataTable("sp_EmpTypeDisplayViaId", "", list);

            return PartialView("_EmptypeEdit", dt);
        }
        public JsonResult _EmptypeEditSubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@EmpTypeId", Request.Form["EmpTypeId"].ToString());
                list.Add("@EmpTypeCode", Request.Form["EmpTypeCode"].ToString());
                list.Add("@EmpTypeName", Request.Form["EmpTypeName"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_EmpType_SaveUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin ", ex.Message);
            }
            return Json(mes);
        }
        public JsonResult _EmptypeStatusUpdate()
        {
            comfun.saveformname("_EmptypeStatusUpdate", "/TimeSheet/_EmptypeStatusUpdate", "Master/General/Employee Type", "Employee Type Status", "N", "EmployeeType", "EmployeeType", "Status", 5);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_EmpType_Active_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }


        public ActionResult EmpStatusList()
        {
            comfun.saveformname("EmployeeStatus", "/TimeSheet/EmpStatusList", "Master/General/Employee Status", "Employee Status form", "Y", "EmployeeStatus", "EmployeeStatus", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("EmployeeStatus");
            return View();
        }
        public ActionResult _EmpstatusList()
        {
            comfun.saveformname("_EmpstatusList", "/TimeSheet/_EmpstatusList", "Master/General/Employee Status", "Employee Status List", "N", "EmployeeStatus", "EmployeeStatus", "List", 4);
            SortedList list = new SortedList();
            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 50;

            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }

            //paging_get_downline1
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            DataTable dt = comfun.fillDataTable("sp_EmpStatusListDisplay", "", list);
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "admin", "_EmpstatusList", "_EmpstatusList");

            HtmlString htm = new HtmlString(paging);
            ViewData["paging"] = htm;
            //return PartialView("_EmpstatusList", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _EmpstatusAdd()
        {
            return PartialView("_EmpstatusAdd");
        }
        public JsonResult _EmpstatusAddSubmit()
        {
            comfun.saveformname("_EmpstatusAddSubmit", "/TimeSheet/_EmpstatusAddSubmit", "Master/General/Employee Status", "Employee Status Add", "N", "EmployeeStatus", "EmployeeStatus", "Add", 2);
            comfun.saveformname("_EmpstatusEditSubmit", "/TimeSheet/_EmpstatusEditSubmit", "Master/General/Employee Status", "Employee Status Edit", "N", "EmployeeStatus", "EmployeeStatus", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@EmpStatusId", 0);
                list.Add("@EmpStatusCode", Request.Form["EmpStatusCode"].ToString());
                list.Add("@EmpStatusName", Request.Form["EmpStatusName"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_EmpStatus_SaveUpdate", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _GradesAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult _EmpstatusEdit()
        {

            string empstatusid = "0";
            empstatusid = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@EmpStatusId", empstatusid);
            DataTable dt = comfun.fillDataTable("sp_EmpStatusDisplayViaId", "", list);
            return PartialView("_EmpstatusEdit", dt);
        }
        public JsonResult _EmpstatusEditSubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@EmpStatusId", Request.Form["EmpStatusId"].ToString());
                list.Add("@EmpStatusCode", Request.Form["EmpStatusCode"].ToString());
                list.Add("@EmpStatusName", Request.Form["EmpStatusName"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_EmpStatus_SaveUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin ", ex.Message);
            }
            return Json(mes);
        }
        public JsonResult _EmpStatusUpdate()
        {
            comfun.saveformname("_EmpStatusUpdate", "/TimeSheet/_EmpStatusUpdate", "Master/General/Employee Status", "Employee Status", "N", "EmployeeStatus", "EmployeeStatus", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_EmpStatus_Active_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }



        //To Display Marital Status List//
        public ActionResult MaritalStatusList()
        {
            comfun.saveformname("MaritalStatus", "/TimeSheet/MaritalStatusList", "Master/General/Marital Status", "Marital Status form", "Y", "MaritalStatus", "MaritalStatus", "View", 1);
            //   ViewData["AccessRights"] = payfun.getAccessRights("MaritalStatus");
            return View();
        }
        public ActionResult _MaritalstatusList()
        {
            comfun.saveformname("_MaritalstatusList", "/TimeSheet/_MaritalstatusList", "Master/General/Marital Status", "Marital Status List", "N", "MaritalStatus", "MaritalStatus", "List", 4);
            SortedList list = new SortedList();
            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 50;

            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }

            //paging_get_downline1
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            DataTable dt = comfun.fillDataTable("sp_MaritalStatusListDisplay", "", list);
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "admin", "_MaritaltatusList", "_MaritaltatusList");

            HtmlString htm = new HtmlString(paging);
            ViewData["paging"] = htm;
            //return PartialView("_MaritalstatusList", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _MaritalstatusAdd()
        {
            return PartialView("_MaritalstatusAdd");
        }
        public JsonResult _MaritalstatusAddSubmit()
        {
            comfun.saveformname("_MaritalstatusAddSubmit", "/TimeSheet/_MaritalstatusAddSubmit", "Master/General/Marital Status", "Marital Status Add", "N", "MaritalStatus", "MaritalStatus", "Add", 2);
            comfun.saveformname("_MaritalstatusEdit", "/TimeSheet/_MaritalstatusEdit", "Master/General/Marital Status", "Marital Status Edit", "N", "MaritalStatus", "MaritalStatus", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@MaritalStatusId", Request.Form["MaritalStatusId"].ToString());
                list.Add("@MaritalStatus", Request.Form["MaritalStatus"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_MaritalStatus_SaveUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _GradesAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult _MaritalstatusEdit()
        {

            string statusid = "0";
            statusid = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@MaritalStatusId", statusid);
            DataTable dt = comfun.fillDataTable("sp_MaritalStatusDisplayViaId", "", list);
            return PartialView("_MaritalstatusEdit", dt);
        }
        public JsonResult _MaritalstatusEditSubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@MaritalStatusId", Request.Form["MaritalStatusId"].ToString());
                list.Add("@MaritalStatus", Request.Form["MaritalStatus"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_MaritalStatus_SaveUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin ", ex.Message);
            }
            return Json(mes);
        }
        public JsonResult _MaritalStatusUpdate()
        {
            comfun.saveformname("_MaritalStatusUpdate", "/TimeSheet/_MaritalStatusUpdate", "Master/General/Marital Status", "Marital Status ", "N", "MaritalStatus", "MaritalStatus", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_MaritalStatus_Active_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }


        //To Display Policy Category List//
        public ActionResult PolicyCategoryList()
        {
            comfun.saveformname("PolicyCategory", "/TimeSheet/PolicyCategoryList", "Master/General/Policy Category", "Policy Category form", "Y", "PolicyCategory", "PolicyCategory", "View", 1);
            //   ViewData["AccessRights"] = payfun.getAccessRights("PolicyCategory");
            return View();
        }
        public ActionResult _PolicyCategoryList()
        {
            comfun.saveformname("_PolicyCategoryList", "/TimeSheet/_PolicyCategoryList", "Master/General/Policy Category", "Policy Category form", "N", "PolicyCategory", "PolicyCategory", "List", 4);
            SortedList list = new SortedList();
            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 50;

            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }

            //paging_get_downline1
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            DataTable dt = comfun.fillDataTable("sp_PolicyCategoryListDisplay", "", list);
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "admin", "_PolicyCategoryList", "_PolicyCategoryList");

            HtmlString htm = new HtmlString(paging);
            ViewData["paging"] = htm;
            //return PartialView("_PolicyCategoryList", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _PolicyCategoryAdd()
        {
            return PartialView("_PolicyCategoryAdd");
        }
        public JsonResult _PolicyCategoryAddSubmit()
        {
            comfun.saveformname("_PolicyCategoryAddSubmit", "/TimeSheet/_PolicyCategoryAddSubmit", "Master/General/Policy Category", "Policy Category Add", "N", "PolicyCategory", "PolicyCategory", "Add", 2);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@PolicyCategoryId", Request.Form["PolicyCategoryId"].ToString());
                list.Add("@PolicyCategoryCode", Request.Form["PolicyCategoryCode"].ToString());
                list.Add("@PolicyCategoryName", Request.Form["PolicyCategoryName"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_PolicyCategory_SaveUpdate", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _GradesAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult _PolicyCategoryEdit()
        {

            string id = "0";
            id = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@PolicyCategoryId", id);
            DataTable dt = comfun.fillDataTable("sp_PolicyCategoryDisplayViaId", "", list);
            return PartialView("_PolicyCategoryEdit", dt);
        }
        public JsonResult _PolicyCategoryEditSubmit()
        {
            comfun.saveformname("_PolicyCategoryEditSubmit", "/TimeSheet/_PolicyCategoryEditSubmit", "Master/General/Policy Category", "Policy Category Edit", "N", "PolicyCategory", "PolicyCategory", "Edit", 3);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@PolicyCategoryId", Request.Form["PolicyCategoryId"].ToString());
                list.Add("@PolicyCategoryCode", Request.Form["PolicyCategoryCode"].ToString());
                list.Add("@PolicyCategoryName", Request.Form["PolicyCategoryName"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_PolicyCategory_SaveUpdate", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _GradesAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        public JsonResult PolicyCategoryStatusUpdate()
        {
            comfun.saveformname("PolicyCategoryStatusUpdate", "/TimeSheet/PolicyCategoryStatusUpdate", "Master/General/Policy Category", "Policy Category Status", "N", "PolicyCategory", "PolicyCategory", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_PolicyCategory_Active_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }





        public ActionResult SendEmail()
        {
            //DataSet ds = comfun.fillDataSet("sp_DeviceDetails_ddl", "", null);
            return View("SendEmail");
        }


        //Temp Shift Assign//
        public ActionResult ShiftAssign()
        {
            comfun.saveformname("ShiftAssign", "/TimeSheet/ShiftAssign", "Shift Management/Shift Assign", "Shift Assign form", "Y", "ShiftAssign", "ShiftAssign", "View", 1);
            //  ViewData["AccessRights"] = payfun.getAccessRights("ShiftAssign");
            DataSet ds = new DataSet();
            ds = comfun.fillDataSet("stpShiftAssignDepartmentddl", "", null);
            return View(ds);
        }
        public ActionResult _Shiftddl()
        {
            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("stpShiftSelect", "", null);
            return PartialView("_Shiftddl", dt);
        }
        public ActionResult _ShiftGroupddl()
        {
            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("stpShiftGroupSelect", "", null);
            return PartialView("_ShiftGroupddl", dt);
        }

        public ActionResult _ShiftAssignJsonList()
        {
            comfun.saveformname("_ShiftAssignJsonList", "/TimeSheet/_ShiftAssignJsonList", "Shift Management/Shift Assign", "Shift Assign List", "N", "ShiftAssign", "ShiftAssign", "List", 4);

            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("stpShiftAssign_list", "", null);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);

        }
        //public ActionResult OutDoorAttendanceList()
        //{

        //    return View();
        //}
        public ActionResult _ShiftAssignList()
        {
            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("stpShiftAssign_list", "", null);
            //return PartialView("_OutDoorAttendancelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _ShiftAssignEdit()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("stpHtisShiftAssignEdit", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _ShiftAssignEditSubmit()
        {
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@Id", Request.Form["Id"].ToString()); ;
                list.Add("@FromDate", Request.Form["FromDate"].ToString());
                list.Add("@ToDate", Request.Form["ToDate"].ToString());
                mes = comfun.executeNonQueryWMessage("stpHtisShiftAssign_Update", "", list).ToString();
            }
            catch (Exception ex)
            {

                throw;
            }

            return Json(mes);
        }
        public ActionResult _ShiftAssignAdd()
        {
            DataSet ds = new DataSet();
            ds = comfun.fillDataSet("stpOutdoorDepartmentddl", "", null);
            return PartialView("_ShiftAssignAdd", ds);
        }
        public ActionResult _ShiftAssignEmployeeList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
            dt = comfun.fillDataTable("stpShiftAssignEmployeeList", "", list);
            return PartialView("_ShiftAssignEmployeeList", dt);
        }
        public ActionResult _ShiftAssigndesignationwiseEmployeeList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
            dt = comfun.fillDataTable("stpShiftAssignListDesignationwise", "", list);
            return PartialView("_ShiftAssignEmployeeList", dt);
        }
        public ActionResult _ShiftBranchwiseEmployeeList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@BranchId", Request.Form["BranchId"].ToString());
            dt = comfun.fillDataTable("stpShiftAssignListBranchwise", "", list);
            return PartialView("_ShiftAssignEmployeeList", dt);
        }
        public ActionResult _UnAssignedList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@BranchId", Request.Form["BranchId"].ToString());
            dt = comfun.fillDataTable("stpUnAssignedShiftList", "", null);
            return PartialView("_ShiftAssignEmployeeList", dt);
        }

        public JsonResult _ShiftAssignSubmit(string EmployeeList, string Id, string FromDate, string ToDate, string ShiftId, string GroupId)
        {
            comfun.saveformname("_ShiftAssignSubmit", "/TimeSheet/_ShiftAssignSubmit", "Shift Management/Shift Assign", "Shift Assign  add", "N", "ShiftAssign", "ShiftAssign", "Add", 2);

            string mes = string.Empty;
            try
            {
                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject(EmployeeList);
                DataTable table = new DataTable();
                table.Columns.Add("EmpId", typeof(int));

                foreach (var item in jsonData)
                {
                    DataRow dr = table.NewRow();
                    dr["EmpId"] = item.EmpId;

                    table.Rows.Add(dr);
                }


                SqlCommand com = new SqlCommand();
                connection conObj = new connection();

                com.Connection = conObj.con;

                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "stpShiftAssign_Accept";
                com.Parameters.AddWithValue("@ID", Id);
                com.Parameters.AddWithValue("@FromDate", FromDate);
                com.Parameters.AddWithValue("@ToDate", ToDate);

                com.Parameters.AddWithValue("@ShiftId", ShiftId);
                com.Parameters.AddWithValue("@GroupId", GroupId);
                //com.Parameters.AddWithValue("@OutTime", OutTime);
                //com.Parameters.AddWithValue("@Remarks", Remarks);
                // com.Parameters.AddWithValue("@fdWeeklyoffDate", attendanceDate);
                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@ShiftEmp";
                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table;
                com.Parameters.Add(parameter);

                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);


                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();

                com.ExecuteNonQuery();

                conObj.con.Close();

                mes = sp.Value.ToString();


            }
            catch (Exception ex)
            {
                // mes = comfun.errorMessage("Attendance _WeeklyoffSubmit", "Error: " + ex.Message);
                mes = "Error: " + ex.Message;



            }
            return Json(mes);
        }
        //public ActionResult HODAttendanceList()
        //{

        //    return View();

        //}

        //Temp Shift//
        //ShiftAssign//
        //public ActionResult ShiftAssign()
        //{
        //    comfun.saveformname("ShiftAssign", "/admin/ShiftAssign", "ShiftAssign", "ShiftAssign  form", "Y", "", "ShiftAssign", "List");
        //    if (Request.QueryString["key"] != null)
        //    {
        //        Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
        //    }

        //    return View();
        //}

        //public ActionResult _ShiftAssignJsonList()
        //{
        //    decimal total_records = 0;
        //    decimal pageno = 1;
        //    decimal pagesize = 50;
        //    if (Request.Form["Page"] != null)
        //    {
        //        pagesize = Convert.ToDecimal(Request.Form["Page"]);
        //    }
        //    if (Request.Form["PageNo"] != null)
        //    {
        //        pageno = Convert.ToDecimal(Request.Form["PageNo"]);
        //    }
        //    SortedList list = new SortedList();
        //    list.Add("@PageNo", pageno);
        //    list.Add("@PageSize", pagesize);
        //    DataTable dt = comfun.fillDataTable("stpShiftAssign_list", "", null);
        //    var json1 = JsonConvert.SerializeObject(dt);
        //    return Json(json1);
        //}
        //public ActionResult _ShiftAssignList()
        //{
        //    decimal total_records = 0;
        //    decimal pageno = 1;
        //    decimal pagesize = 50;
        //    if (Request.Form["Page"] != null)
        //    {
        //        pagesize = Convert.ToDecimal(Request.Form["Page"]);
        //    }
        //    if (Request.Form["PageNo"] != null)
        //    {
        //        pageno = Convert.ToDecimal(Request.Form["PageNo"]);
        //    }
        //    SortedList list = new SortedList();
        //    list.Add("@PageNo", pageno);
        //    list.Add("@PageSize", pagesize);
        //    DataTable dt = comfun.fillDataTable("stpShiftAssign_list", "", list);
        //    return PartialView(dt);
        //}
        ////[customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        //public ActionResult _ShiftAssignSubmitView()
        //{
        //    // comfun.saveformname("CountriesAdd", "/admin/CountriesAdd", "Country Add", "Country view", "N", "Country Add", "Country", "Add");
        //    ViewData["Id"] = "0";
        //    ViewData["EmployeeID"] = "";
        //    ViewData["fromDate"] = "";
        //    ViewData["toDate"] = "";
        //    ViewData["shiftID"] = "";
        //    ViewData["GroupID"] = "";
        //    if (Request.Form["CountryId"].ToString() != "0")
        //    {
        //        ViewData["Id"] = Request.Form["Id"].ToString();
        //        ViewData["EmployeeID"] = Request.Form["EmployeeID"].ToString();
        //        ViewData["fromDate"] = Request.Form["fromDate"].ToString();
        //        ViewData["toDate"] = Request.Form["toDate"].ToString();
        //        ViewData["shiftID"] = Request.Form["shiftID"].ToString();
        //        ViewData["GroupID"] = Request.Form["GroupID"].ToString();
        //    }
        //    return PartialView("_ShiftAssignSubmit");
        //}
        //public ActionResult _OnBoardcandidateList()
        //{
        //    DataTable dt = comfun.fillDataTable("Htis_OnBoardcandidateSelect", "", null);
        //    return PartialView("_OnBoardcandidateList", dt);
        //}
        //public JsonResult _ShiftAssignSubmit()
        //{
        //    string mes = string.Empty;
        //    try
        //    {
        //        SortedList list = new SortedList();
        //        list.Add("@Id", Request.Form["Id"].ToString());
        //        list.Add("@EmployeeID", Request.Form["EmployeeID"].ToString());
        //        list.Add("@fromDate", Request.Form["fromDate"].ToString());
        //        list.Add("@toDate", Request.Form["toDate"].ToString());
        //        list.Add("@shiftID", Request.Form["shiftID"].ToString());
        //        list.Add("@GroupID", Request.Form["GroupID"].ToString());
        //        //list.Add("@CreatedBy", "1");
        //        //list.Add("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
        //        mes = comfun.executeNonQueryWMessage("stpShiftAssign_Accept", "", list).ToString();
        //    }
        //    catch (Exception ex)
        //    {
        //        mes = "Error: " + ex.Message;
        //    }
        //    return Json(mes);
        //}
        //public JsonResult _ShiftAssignUpdate()
        //{
        //    string mes = string.Empty;
        //    try
        //    {
        //        SortedList list = new SortedList();
        //        list.Add("@Id", Request.Form["Id"].ToString());
        //        list.Add("@IsActive", Request.Form["IsActive"].ToString());
        //        mes = comfun.executeNonQueryWMessage("stpCountryActivationStatusUpdate", "", list).ToString();
        //    }
        //    catch (Exception ex)
        //    {
        //        mes = ex.Message;
        //    }
        //    return Json(mes);

        //}

        //ShiftAssign//
        public ActionResult ShiftGroup()
        {
            comfun.saveformname("ShiftGroup", "/TimeSheet/ShiftGroup", "Shift Management/Shift Group", "Shift Group form", "Y", "ShiftGroup", "ShiftGroup", "View", 1);
            //    ViewData["AccessRights"] = payfun.getAccessRights("ShiftGroup");

            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            SortedList list = new SortedList();
            list.Add("@PageNo", 1);
            list.Add("@PageSize", 10);
            DataTable dt = comfun.fillDataTable("stpSHIFT_Select", "", null);
            return View(dt);
        }
        public ActionResult _ShiftGroupJsonList()
        {
            comfun.saveformname("_ShiftGroupJsonList", "/TimeSheet/_ShiftGroupJsonList", "Shift Management/Shift Group", "Shift Group List", "N", "ShiftGroup", "ShiftGroup", "List", 4);

            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 50;
            if (Request.Form["Page"] != null)
            {
                pagesize = Convert.ToDecimal(Request.Form["Page"]);
            }
            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }
            SortedList list = new SortedList();
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            DataTable dt = comfun.fillDataTable("stpShiftGroup_list", "", null);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult _ShiftGroupList()
        {
            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 50;
            if (Request.Form["Page"] != null)
            {
                pagesize = Convert.ToDecimal(Request.Form["Page"]);
            }
            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }
            SortedList list = new SortedList();

            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            DataTable dt = comfun.fillDataTable("stpShiftGroup_list", "", null);
            return PartialView(dt);
        }
        //[customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]

        public ActionResult _ShiftGroupSubmitView()
        {
            DataTable dt = comfun.fillDataTable("stpSHIFT_Select", "", null);
            // comfun.saveformname("CountriesAdd", "/admin/CountriesAdd", "Country Add", "Country view", "N", "Country Add", "Country", "Add");
            ViewData["Id"] = "0";
            ViewData["GroupName"] = "";
            ViewData["GroupShortName"] = "";
            //ViewData["Shiftids"] = "";
            ViewData["Shiftids"] = Request.Form["Shiftids"].ToString();
            if (Request.Form["Id"].ToString() != "0")
            {
                ViewData["Id"] = Request.Form["Id"].ToString();
                ViewData["GroupName"] = Request.Form["GroupName"].ToString();
                ViewData["GroupShortName"] = Request.Form["GroupShortName"].ToString();
                ViewData["Shiftids"] = Request.Form["Shiftids"].ToString();
            }
            return PartialView("_ShiftGroupSubmit", dt);
        }

        public JsonResult _ShiftGroupSubmit()
        {
            comfun.saveformname("_ShiftGroupSubmit", "/TimeSheet/_ShiftGroupSubmit", "Shift Management/Shift Group", "Shift Group  add", "N", "ShiftGroup", "ShiftGroup", "Add", 2);
            comfun.saveformname("_ShiftGroupSubmit", "/TimeSheet/_ShiftGroupSubmit", "Shift Management/Shift Group", "Shift Group  Edit", "N", "ShiftGroup", "ShiftGroup", "Edit", 3);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@GroupName", Request.Form["GroupName"].ToString());
                list.Add("@GroupShortName", Request.Form["GroupShortName"].ToString());
                list.Add("@Shiftids", Request.Form["Shiftids"].ToString());
                //list.Add("@CreatedBy", "1");
                //list.Add("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                mes = comfun.executeNonQueryWMessage("stpShiftGroup_Accept", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
            }
            return Json(mes);
        }
        public ActionResult ddlShiftselect()
        {
            DataTable dt = comfun.fillDataTable("stpSHIFT_Select", "", null);
            return PartialView("_ddlShiftselect", dt);
        }
        public JsonResult _ShiftGroupStatusUpdate()
        {
            comfun.saveformname("_ShiftGroupStatusUpdate", "/TimeSheet/_ShiftGroupStatusUpdate", "Shift Management/Shift Group", "Shift Group  Status", "N", "ShiftGroup", "ShiftGroup", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpShiftGroupStatus", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }
        public ActionResult _ShiftGroupEdit()
        {
            string LeaveId = "0";
            LeaveId = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@Id", LeaveId);
            DataTable dt = comfun.fillDataTable("stpShiftGroupDisplayviaID", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult ShiftMasterDisplay()
        {
            comfun.saveformname("ShiftMaster", "/TimeSheet/ShiftMasterDisplay", "Shift Management/Shift Master", "Shift Master form", "Y", "ShiftMaster", "ShiftMaster", "View", 1);
            //ViewData["AccessRights"] = payfun.getAccessRights("ShiftMaster");
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            return View();
        }
        public ActionResult _ShiftMasterJsonList()
        {
            comfun.saveformname("ShiftMaster", "/TimeSheet/_ShiftMasterJsonList", "Shift Management/ShiftMaster", "Shift Master List", "N", "ShiftMaster", "ShiftMaster", "List", 4);

            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 50;
            if (Request.Form["Page"] != null)
            {
                pagesize = Convert.ToDecimal(Request.Form["Page"]);
            }
            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }
            SortedList list = new SortedList();
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            DataTable dt = comfun.fillDataTable("stpShiftMaster_list", "", null);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        //public ActionResult _ShiftMasterList()
        //{

        //    decimal total_records = 0;
        //    decimal pageno = 1;
        //    decimal pagesize = 50;
        //    if (Request.Form["Page"] != null)
        //    {
        //        pagesize = Convert.ToDecimal(Request.Form["Page"]);
        //    }
        //    if (Request.Form["PageNo"] != null)
        //    {
        //        pageno = Convert.ToDecimal(Request.Form["PageNo"]);
        //    }
        //    SortedList list = new SortedList();
        //    list.Add("@PageNo", pageno);
        //    list.Add("@PageSize", pagesize);
        //    DataTable dt = comfun.fillDataTable("stpShiftMaster_list", "", null);
        //    return PartialView(dt);
        //}

        //public ActionResult _shiftMasterSubmitView()
        //{
        //    // comfun.saveformname("CountriesAdd", "/admin/CountriesAdd", "Country Add", "Country view", "N", "Country Add", "Country", "Add");
        //    ViewData["Id"] = "0";
        //    ViewData["Shiftname"] = "";
        //    ViewData["ShiftShortname"] = "";
        //    ViewData["Fromtime"] = "";
        //    ViewData["Totime"] = "";
        //    ViewData["Breakfrom"] = "";
        //    ViewData["BreakTo"] = "";
        //    ViewData["InfromMin"] = "";
        //    ViewData["INToMin"] = "";
        //    ViewData["OutFromMin"] = "";
        //    ViewData["OutToMin"] = "";
        //    ViewData["IsNightShift"] = "";
        //    ViewData["IsDefault"] = "";
        //    ViewData["IsActive"] = "";
        //    if (Request.Form["ID"].ToString() != "0")
        //    {
        //        ViewData["Id"] = Request.Form["Id"].ToString();
        //        ViewData["Shiftname"] = Request.Form["Shiftname"].ToString();
        //        ViewData["ShiftShortname"] = Request.Form["ShiftShortname"].ToString();
        //        ViewData["Fromtime"] = Request.Form["Fromtime"].ToString();
        //        ViewData["Totime"] = Request.Form["Totime"].ToString();
        //        ViewData["Breakfrom"] = Request.Form["Breakfrom"].ToString();
        //        ViewData["BreakTo"] = Request.Form["BreakTo"].ToString();
        //        ViewData["InfromMin"] = Request.Form["InfromMin"].ToString();
        //        ViewData["INToMin"] = Request.Form["INToMin"].ToString();
        //        ViewData["OutFromMin"] = Request.Form["OutFromMin"].ToString();
        //        ViewData["OutToMin"] = Request.Form["OutToMin"].ToString();
        //        ViewData["IsNightShift"] = Request.Form["IsNightShift"].ToString();
        //        ViewData["IsDefault"] = Request.Form["IsDefault"].ToString();
        //        ViewData["IsActive"] = Request.Form["IsActive"].ToString();

        //    }
        //    return PartialView("_ShiftMasterSubmit");
        //}

        public JsonResult _ShiftMasterSubmit()
        {
            comfun.saveformname("_ShiftMasterSubmit", "/TimeSheet/_ShiftMasterSubmit", "Shift Management/ShiftMaster", "Shift Master  add", "N", "ShiftMaster", "ShiftMaster", "Add", 2);
            comfun.saveformname("_ShiftMasterSubmit", "/TimeSheet/_ShiftMasterSubmit", "Shift Management/ShiftMaster", "Shift Master  edit", "N", "ShiftMaster", "ShiftMaster", "Edit", 3);


            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@Shiftname", Request.Form["Shiftname"].ToString());
                list.Add("@ShiftShortname", Request.Form["ShiftShortname"].ToString());
                list.Add("@Fromtime", Request.Form["Fromtime"].ToString());
                list.Add("@Totime", Request.Form["Totime"].ToString());
                list.Add("@Breakfrom", Request.Form["Breakfrom"].ToString());
                list.Add("@BreakTo", Request.Form["BreakTo"].ToString());
                list.Add("@InfromMin", Request.Form["InfromMin"].ToString());
                list.Add("@INToMin", Request.Form["INToMin"].ToString());
                list.Add("@OutFromMin", Request.Form["OutFromMin"].ToString());
                list.Add("@OutToMin", Request.Form["OutToMin"].ToString());
                list.Add("@IsNightShift", Request.Form["IsNightShift"].ToString());
                list.Add("@IsDefault", Request.Form["IsDefault"].ToString());
                // list.Add("@IsActive", Request.Form["IsActive"].ToString());
                //list.Add("@CreatedBy", "1");
                //list.Add("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                mes = comfun.executeNonQueryWMessage("stpShiftMaster_Accept", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
            }
            return Json(mes);
        }

        public JsonResult _ShiftMasterStatusUpdate()
        {
            comfun.saveformname("_ShiftMasterStatusUpdate", "/TimeSheet/_ShiftMasterStatusUpdate", "Shift Management/Shift Master", "Shift Master form", "N", "ShiftMaster", "ShiftMaster", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_Shift_Active_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }



        public ActionResult _ShiftMasterEdit()
        {
            //comfun.saveformname("ShiftMasterEdit", "/admin/ShiftMasterEdit", "ShiftMaster Add", "ShiftMaster Edit", "N", "ShiftMasterEdit", "ShiftMaster", "Edit");

            string shiftid = "0";

            shiftid = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@Id", shiftid);
            DataTable dt = comfun.fillDataTable("ShiftName_SelectWidId", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult PartnerDetailddl()
        {
            SortedList list = new SortedList();

            DataTable dt = comfun.fillDataTable("sp_Partnerddl", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult RateCardDetailddl()
        {
            SortedList list = new SortedList();

            DataTable dt = comfun.fillDataTable("sp_RateCardddl", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _Countryddl()
        {
            SortedList list = new SortedList();

            DataTable dt = comfun.fillDataTable("stpCountryddl", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _Stateddl()
        {
            SortedList list = new SortedList();
            list.Add("@CountryId", Request.Form["CountryId"].ToString());
            DataTable dt = comfun.fillDataTable("stpStateddl", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _Cityddl()

        {
            SortedList list = new SortedList();
            list.Add("@StateId", Request.Form["StateId"].ToString());
            DataTable dt = comfun.fillDataTable("stpCitybyStateIdsddl", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult UserRole()
        {
            // comfun.saveformname("UserRole", "/users/UserRole", "Roles - Main Form", "Roles - Assign roles to user and add addon activities to User", "Y");

            DataTable dt = comfun.fillDataTable("stpRoleddl", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _Designation()
        {
            // comfun.saveformname("UserRole", "/users/UserRole", "Roles - Main Form", "Roles - Assign roles to user and add addon activities to User", "Y");

            SortedList list = new SortedList();
            list.Add("@PageNo", 1);
            list.Add("@PageSize", 1000);
            DataTable dt = comfun.fillDataTable("Designation_Select", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _Gender()
        {
            // comfun.saveformname("UserRole", "/users/UserRole", "Roles - Main Form", "Roles - Assign roles to user and add addon activities to User", "Y");


            DataTable dt = comfun.fillDataTable("stpGenderddl", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _Department()
        {
            // comfun.saveformname("UserRole", "/users/UserRole", "Roles - Main Form", "Roles - Assign roles to user and add addon activities to User", "Y");

            SortedList list = new SortedList();
            list.Add("@PageNo", 1);
            list.Add("@PageSize", 10000);
            DataTable dt = comfun.fillDataTable("Departments_Select", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult PartnerDetailList()
        {
            comfun.saveformname("PartnerDetailList", "/TimeSheet/PartnerDetailList", "Master/General/Partner&nbsp;", "Partner View", "Y", "PartnerDetailList", "PartnerDetailList", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("PartnerDetailList");
            return View();
        }
        public ActionResult _PartnerDetailList()
        {
            //comfun.saveformname("ShiftMasterEdit", "/admin/ShiftMasterEdit", "ShiftMaster Add", "ShiftMaster Edit", "N", "ShiftMasterEdit", "ShiftMaster", "Edit");
            comfun.saveformname("_PartnerDetailList", "/TimeSheet/_PartnerDetailList", "Master/General/Partner&nbsp;", "Partner List", "N", "PartnerDetailList", "PartnerDetailList", "List", 4);
            SortedList list = new SortedList();

            DataTable dt = comfun.fillDataTable("sp_PartnerList", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _PartnerDetailSubmit()
        {
            comfun.saveformname("_PartnerDetailSubmit", "/TimeSheet/_PartnerDetailSubmit", "Master/General/Partner&nbsp;", "Partner Master  add", "N", "PartnerDetailList", "PartnerDetailList", "Add", 2);
            comfun.saveformname("_PartnerDetailSubmit", "/TimeSheet/_PartnerDetailSubmit", "Master/General/Partner&nbsp;", "Partner Master  edit", "N", "PartnerDetailList", "PartnerDetailList", "Edit", 3);


            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@PartnerId", Request.Form["PartnerId"].ToString());
                list.Add("@PartnerName", Request.Form["PartnerName"].ToString());
                list.Add("@PartnerCode", Request.Form["PartnerCode"].ToString());
                list.Add("@PartnerAddress", Request.Form["PartnerAddress"].ToString());
                list.Add("@Password", Request.Form["Password"].ToString());
                list.Add("@Isaccess", Request.Form["Isaccess"].ToString());
                list.Add("@UserName", Request.Form["UserName"].ToString());
                list.Add("@PONo", Request.Form["PONo"].ToString());
                list.Add("@Email", Request.Form["Email"].ToString());
                list.Add("@CountryId", Request.Form["CountryId"].ToString());
                list.Add("@StateId", Request.Form["StateId"].ToString());
                list.Add("@CityId", Request.Form["CityId"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                list.Add("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                mes = comfun.executeNonQueryWMessage("Partner_save", "", list).ToString();
                if (Request.Form["Isaccess"].ToString() == "Y")
                {
                    if (!mes.Contains("Error"))
                    {
                        if (!mes.Contains("NoEmail"))
                        {
                            payfun.SendEmail(Request.Form["PartnerName"].ToString(), Request.Form["Email"].ToString(), "User Name Password", mes);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
            }
            return Json(mes);
        }

        public JsonResult _PartnerDetailStatusUpdate()
        {
            comfun.saveformname("_PartnerShiftMasterStatusUpdate", "/TimeSheet/_PartnerShiftMasterStatusUpdate", "Master/General/Partner&nbsp;", "Partner Master form", "N", "PartnerDetailList", "PartnerDetailList", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpUpdateStatusPartner", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }
        public ActionResult _PartnerDetailEdit()
        {
            //comfun.saveformname("ShiftMasterEdit", "/admin/ShiftMasterEdit", "ShiftMaster Add", "ShiftMaster Edit", "N", "ShiftMasterEdit", "ShiftMaster", "Edit");

            string PartnerId = "0";

            PartnerId = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@Id", PartnerId);
            DataTable dt = comfun.fillDataTable("stpPartnerDetailEdit", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _MyTeamddl()
        {
            //comfun.saveformname("ShiftMasterEdit", "/admin/ShiftMasterEdit", "ShiftMaster Add", "ShiftMaster Edit", "N", "ShiftMasterEdit", "ShiftMaster", "Edit");



            DataTable dt = comfun.fillDataTable("StpMyTeamSelect", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }



        public ActionResult _Levelddl()
        {
            //comfun.saveformname("ShiftMasterEdit", "/admin/ShiftMasterEdit", "ShiftMaster Add", "ShiftMaster Edit", "N", "ShiftMasterEdit", "ShiftMaster", "Edit");




            DataTable dt = comfun.fillDataTable("stpHtisLevelsList", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult RateCard()
        {
            comfun.saveformname("RateCard", "/TimeSheet/RateCard", "Master/General/RateCardMaster", "Rate Card Master form", "Y", "RateCard", "RateCard", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("RateCard");
            return View();
        }
        public ActionResult _RateCardList()
        {
            //comfun.saveformname("ShiftMasterEdit", "/admin/ShiftMasterEdit", "ShiftMaster Add", "ShiftMaster Edit", "N", "ShiftMasterEdit", "ShiftMaster", "Edit");
            comfun.saveformname("_RateCardList", "/TimeSheet/_RateCardList", "Master/General/RateCardMaster", "Rate Card Master form", "N", "RateCard", "RateCard", "List", 4);

            SortedList list = new SortedList();

            DataTable dt = comfun.fillDataTable("stpRateCardList", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _RateDetailStatus()
        {
            comfun.saveformname("_RateDetailStatus", "/TimeSheet/_RateDetailStatus", "Master/General/RateCardMaster", "Rate Card Master form", "N", "RateCard", "RateCard", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpUpdateStatusRateCard", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }
        public JsonResult _RateCardSubmit()
        {
            comfun.saveformname("_RateCardSubmit", "/TimeSheet/_RateCardSubmit", "Master/General/RateCardMaster", "Rate Card add", "N", "RateCard", "RateCard", "Add", 2);
            comfun.saveformname("_RateCardSubmit", "/TimeSheet/_RateCardSubmit", "Master/General/RateCardMaster", "Rate Card  edit", "N", "RateCard", "RateCard", "Edit", 3);


            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@CardName", Request.Form["CardName"].ToString());
                list.Add("@WEF", Request.Form["WEF"].ToString());
                list.Add("@PartnerID", Request.Form["PartnerID"].ToString());
                list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
                list.Add("@LevelId", Request.Form["LevelId"].ToString());
                list.Add("@Rate", Request.Form["Rate"].ToString());

                // list.Add("@IsActive", Request.Form["IsActive"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                list.Add("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                mes = comfun.executeNonQueryWMessage("stpRateCard_Accept", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
            }
            return Json(mes);
        }
        public ActionResult _RateDetailEdit()
        {
            //comfun.saveformname("ShiftMasterEdit", "/admin/ShiftMasterEdit", "ShiftMaster Add", "ShiftMaster Edit", "N", "ShiftMasterEdit", "ShiftMaster", "Edit");


            string RateId = "0";
            RateId = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@Id", RateId);
            DataTable dt = comfun.fillDataTable("stpRateCardEdit", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult Holiday()
        {
            comfun.saveformname("PublicHoliday", "/TimeSheet/Holiday", "Master/General/Public Holiday", "Public Holiday form", "Y", "PublicHoliday", "PublicHoliday", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("PublicHoliday");
            DataSet dt = new DataSet();
            dt = comfun.fillDataSet("StateBranch_Select", "", null);
            return View(dt);
        }
        public ActionResult _Select_BranchByState()
        {
            SortedList list = new SortedList();
            list.Add("@StateId", Request.Form["StateId"].ToString());
            DataTable dt = comfun.fillDataTable("stpSelectBranchByState", "", list);
            return PartialView("_Select_BranchByState", dt);
        }
        public ActionResult _HolidayList()
        {
            comfun.saveformname("_HolidayList", "/TimeSheet/_HolidayList", "Master/General/Public Holiday", "Public Holiday List", "N", "PublicHoliday", "PublicHoliday", "List", 4);
            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("StpviksatHolidaylist", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _HolidayAdd()
        {
            DataSet dt = new DataSet();
            dt = comfun.fillDataSet("StateBranch_Select", "", null);

            return PartialView(dt);
        }
        public JsonResult _HolidayAddSumbit()
        {
            comfun.saveformname("_HolidayAddSumbit", "/TimeSheet/_HolidayAddSumbit", "Master/General/Public Holiday", "Public Holiday Add", "N", "PublicHoliday", "PublicHoliday", "Add", 2);
            comfun.saveformname("_HolidayEditSumbit", "/TimeSheet/_HolidayEditSumbit", "Master/General/Public Holiday", "Public Holiday Edit", "N", "PublicHoliday", "PublicHoliday", "Edit", 3);

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@HolidayName", Request.Form["Holiday"].ToString());
                list.Add("@HolidayDate", Request.Form["FromDate"].ToString());
                list.Add("@StateId", Request.Form["State"].ToString());
                list.Add("@CostCenterId", Request.Form["Branch"].ToString());
                list.Add("@Todate", Request.Form["ToDate"].ToString());
                list.Add("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                list.Add("@IsPublic", Request.Form["IsPublic"].ToString());
                list.Add("@CreatedBy", Session["EmpId"]);
                mes = comfun.executeNonQueryWMessage("Holidays_Save", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
                throw;
            }

            return Json(mes);


        }
        public ActionResult _HolidayEdit()
        {
            SortedList list = new SortedList();
            list.Add("@HolidayId", Request.Form["Id"].ToString());
            DataSet ds = comfun.fillDataSet("Holiday_SelectwithId", "", list);
            var json1 = JsonConvert.SerializeObject(ds.Tables[0]);
            return Json(json1);

        }
        public JsonResult _HolidayEditSumbit()
        {
            comfun.saveformname("_HolidayAddSumbit", "/TimeSheet/_HolidayAddSumbit", "Master/General/Public Holiday", "Public Holiday Add", "N", "PublicHoliday", "PublicHoliday", "Add", 2);
            comfun.saveformname("_HolidayEditSumbit", "/TimeSheet/_HolidayEditSumbit", "Master/General/Public Holiday", "Public Holiday Edit", "N", "PublicHoliday", "PublicHoliday", "Edit", 3);

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@HolidayId", Request.Form["HolidayId"].ToString());
                //list.Add("@HolidayName", Request.Form["Holiday"].ToString());
                //list.Add("@HolidayDate", Convert.ToDateTime(Request.Form["FromDate"].ToString()));
                //list.Add("@StateId", Request.Form["State"].ToString());
                //list.Add("@CostCenterId", Request.Form["Branch"].ToString());
                //list.Add("@Todate",Convert.ToDateTime(Request.Form["ToDate"].ToString()));

                list.Add("@HolidayName", Request.Form["Holiday"].ToString());
                list.Add("@HolidayDate", Request.Form["FromDate"].ToString());
                list.Add("@StateId", Request.Form["State"].ToString());
                list.Add("@CostCenterId", Request.Form["Branch"].ToString());
                list.Add("@Todate", Request.Form["ToDate"].ToString());
                list.Add("@IsPublic", Request.Form["IsPublic"].ToString());

                mes = comfun.executeNonQueryWMessage("Holidays_Edit", "", list).ToString();
            }

            catch (Exception ex)
            {
                mes = ex.Message;
                throw;
            }

            return Json(mes);


        }

        public JsonResult _HolidayStatusUpdate()
        {
            comfun.saveformname("_HolidayStatusUpdate", "/TimeSheet/_HolidayStatusUpdate", "Master/General/Public Holiday", "Public Holiday Status", "N", "PublicHoliday", "PublicHoliday", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpHolidayStatusUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }

        #endregion
        public ActionResult RoleList()
        {
            comfun.saveformname("Role", "/TimeSheet/RoleList", "User Management/RoleList", "Role Master Main Form", "Y", "Role", "Role", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("Role");

            return View();
        }

        public JsonResult _roleSave()
        {
            comfun.saveformname("_roleSave", "/TimeSheet/_roleSave", "User Management/RoleList", "Role Master Add", "N", "Role", "Role", "Add", 2);
            comfun.saveformname("_roleSave", "/TimeSheet/_roleSave", "User Management/RoleList", "Role Master Edit", "N", "Role", "Role", "Edit", 3);

            string mes = string.Empty;
            SortedList list = new SortedList();
            list.Add("@RoleId", Request.Form["Id"].ToString());
            list.Add("@RoleName", Request.Form["RoleName"].ToString());
            mes = comfun.executeNonQueryWMessage("sp_RoleUpdate", "", list).ToString();
            return Json(mes);
        }
        public JsonResult _roleshow()
        {
            comfun.saveformname("_roleshow", "/TimeSheet/_roleshow", "User Management/RoleList", "Role Master List", "N", "Role", "Role", "List", 4);

            System.Data.DataTable dt = new DataTable();
            dt = comfun.fillDataTable("UserRoles_Select", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _RoleActivationStatusUpdate()
        {
            comfun.saveformname("_RoleActivationStatusUpdate", "/TimeSheet/_RoleActivationStatusUpdate", "User Management/RoleList", "Role Master Status", "N", "Role", "Role", "Status", 5);

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@RoleId", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_RoleStatusUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                // mes = ex.Message();
                throw;
            }


            return Json(mes);
        }
        public ActionResult _RoleEdit()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("stpActiveRolesById", "", list);

            return PartialView("_roleEdit", dt);

        }
        public JsonResult _RoleEditSubmit()
        {
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@RoleID", Request.Form["RoleId"].ToString());
                list.Add("@RoleName", Request.Form["RoleName"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_RoleUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {

                throw;
            }

            return Json(mes);
        }
        public ActionResult RolePermission()
        {

            DataSet ds = new DataSet();
            ds = comfun.fillDataSet("stpModuleRole", "", null);
            // ViewBag.DOM_TreeViewMenu = PopulateMenuDataTable();
            return View(ds);
        }
        //----------------------------------------------tree View ----------------------------------------------------

        public ActionResult _MenuTreeJS()
        {
            SortedList list = new SortedList();
            list.Add("@RoleId", Request.Form["RoleId"].ToString());
            System.Data.DataTable dt = comfun.fillDataTable("MenuTree2", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        //public string GetTreeData()
        //{
        //    List<FlatObject> flatObjects = new List<FlatObject>
        //    {
        //        new FlatObject("Category",1,0),
        //        new FlatObject("SubCategory1",2,1),
        //        new FlatObject("SubCategory2",3,1),
        //        new FlatObject("Item1",4,2),
        //        new FlatObject("Item2",5,2),
        //        new FlatObject("Item3",6,2),
        //        new FlatObject("Item4",7,3),
        //        new FlatObject("Item5",8,3),
        //        new FlatObject("Item6",9,3)
        //    };
        //    var recursiveObjects = FillRecursive(flatObjects, 0);
        //    string myjsonmodel = new JavaScriptSerializer().Serialize(recursiveObjects);
        //    return myjsonmodel;
        //}

        //private static List<RecursiveObject> FillRecursive(List<FlatObject> flatObjects, Int64 parentId)
        //{
        //    List<RecursiveObject> recursiveObjects = new List<RecursiveObject>();
        //    foreach (var item in flatObjects.Where(x => x.ParentId.Equals(parentId)))
        //    {
        //        recursiveObjects.Add(new RecursiveObject
        //        {
        //            data = item.data,
        //            id = item.Id,
        //            attr = new FlatTreeAttribute { id = item.Id.ToString(), selected = false },
        //            children = FillRecursive(flatObjects, item.Id)
        //        });
        //    }
        //    return recursiveObjects;
        //}


        public ActionResult _MenuTreeByRoleJS()
        {
            SortedList list = new SortedList();
            list.Add("@RoleId", Request.Form["RoleId"].ToString());
            System.Data.DataTable dt = comfun.fillDataTable("MenuTree3", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }


        public ActionResult _UserFormPermissionSave()
        {
            string mes = "";
            try
            {
                SortedList list = new SortedList();
                list.Add("@RoleId", Request.Form["RoleId"].ToString());
                list.Add("@FormIds", Request.Form["FormIds"].ToString());
                mes = comfun.executeNonQueryWMessage("UserActivityRoleDefault_save", "", list).ToString();
                //
            }
            catch (Exception ex)
            {
                mes = "Error:" + ex.Message;
            }
            return Json(mes);
        }
        public ActionResult FreezeAttendance()
        {
            comfun.saveformname("FreezeAttendance", "/TimeSheet/FreezeAttendance", "Attendance Management/FreezeAttendance&nbsp;", "Freeze Attendance form", "Y", "FreezeAttendance", "FreezeAttendance", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("FreezeAttendance");
            SortedList list = new SortedList();
            list.Add("@PageNo", 1);
            list.Add("@PageSize", 1000);
            DataTable dt = comfun.fillDataTable("Department_Select", "", list);
            return View(dt);
        }
        public JsonResult _FreezeAttendanceList()
        {
            comfun.saveformname("_FreezeAttendanceList", "/TimeSheet/_FreezeAttendanceList", "Attendance Management/FreezeAttendance&nbsp;", "Freeze Attendance List", "N", "FreezeAttendance", "FreezeAttendance", "List", 4);
            SortedList list = new SortedList();
            string fromDate = Request.Form["StartDate"].ToString();
            int startYear = Convert.ToDateTime(fromDate).Year;
            int startMonth = Convert.ToDateTime(fromDate).Month;

            DateTime startDate = new DateTime(startYear, startMonth, 1);
            DateTime endDate = startDate.AddMonths(1).AddDays(-1);
            list.Add("@StartDate", startDate.ToString("dd/MMM/yyyy"));
            list.Add("@EndDate", endDate.ToString("dd/MMM/yyyy"));
            list.Add("@BranchID", Request.Form["BranchId"].ToString());
            list.Add("@DepartmentID", Request.Form["DepartmentId"].ToString());
            list.Add("@CircleId", Request.Form["CircleId"].ToString());
            list.Add("@CategoryId", Request.Form["CategoryId"].ToString());
            list.Add("@MainCategoryId", Request.Form["MainCategoryId"].ToString());
            DataTable dt = comfun.fillDataTable("stpEmployeeAttendance_MonthlyLockList", "", list);
            //return PartialView("_HelpdeskGroupMappingList", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _EmployeeFreeze()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@FromDate", Request.Form["FromDate"].ToString());
            dt = comfun.fillDataTable("stpEmployeeFreezeddl", null, list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        //public ActionResult _helpdeskGroupMappingBranch()
        //{
        //    SortedList list = new SortedList();
        //    list.Add("@PageNo", 1);
        //    list.Add("@PageSize", 1000);
        //    DataTable dt = comfun.fillDataTable("Branch_Select", "", list);
        //    var json = JsonConvert.SerializeObject(dt);
        //    return Json(json);
        //}

        public ActionResult _helpDeskDesignation()
        {
            SortedList list = new SortedList();
            list.Add("@PageNo", 1);
            list.Add("@PageSize", 1000);
            DataTable dt = comfun.fillDataTable("Designation_Select", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _FreezeEmployee()
        {
            SortedList list = new SortedList();
            string fromDate = Request.Form["FromDate"].ToString();
            int startYear = Convert.ToDateTime(fromDate).Year;
            int startMonth = Convert.ToDateTime(fromDate).Month;
            string Type = Request.Form["Type"].ToString();
            DateTime startDate = new DateTime(startYear, startMonth, 1);
            list.Add("@StartDate", startDate.ToString("dd/MMM/yyyy"));
            DateTime endDate = startDate.AddMonths(1).AddDays(-1);
            list.Add("@EndDate", endDate.ToString("dd/MMM/yyyy"));
            list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
            list.Add("@BranchId", Request.Form["BranchId"].ToString());
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@CategoryId", Request.Form["CategoryId"].ToString());
            list.Add("@CircleId", Request.Form["CircleId"].ToString());
            list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
            list.Add("@LoginID", Session["EmpId"].ToString());
            DataTable dt = new DataTable();
            if (Type == "Freeze")
            {
                dt = comfun.fillDataTable("stpEmployeeAttendance_MonthlyLock", "", list);

            }
            if (Type == "UnFreeze")
            {
                dt = comfun.fillDataTable("stpEmployeeAttendance_MonthlyUnlock", "", list);

            }
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _UpdateHelpDeskMapping()
        {
            SortedList list = new SortedList();
            list.Add("@HelpdeskGroupId", Request.Form["HelpdeskGroupId"].ToString());
            DataSet ds = comfun.fillDataSet("sp_HDGDisplayViaId", "", list);
            var json1 = JsonConvert.SerializeObject(ds.Tables[0]);
            return Json(json1);
        }
        public ActionResult _HelpDeskMappedEmployees()
        {
            SortedList list = new SortedList();
            list.Add("@HelpdeskGroupID", Request.Form["HelpdeskGroupId"].ToString());
            DataTable dt = comfun.fillDataTable("sp_HelpdeskGroupMappedEmpList", "", list);
            //return PartialView("_HelpDeskMappedEmployees", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _EmpListViaDepartmentDisplay()
        {
            SortedList list = new SortedList();
            list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
            DataTable dt = comfun.fillDataTable("sp_EmpListViaDepartmentId", "", list);
            List<Employees> empList = new List<Employees>();
            empList = (from DataRow dr in dt.Rows
                       select new Employees()
                       {
                           EmployeeId = dr["EmployeeID"].ToString(),
                           EmployeeCode = dr["EmployeeCode"].ToString(),
                           EmployeeEmail = dr["Email"].ToString(),
                           EmployeeName = dr["EmployeeName"].ToString(),
                       }).ToList();
            return Json(empList, JsonRequestBehavior.AllowGet);
        }
        public ActionResult _EmpListViaDesignationDisplay()
        {
            SortedList list = new SortedList();
            list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
            DataTable dt = comfun.fillDataTable("sp_EmpListViaDesignationId", "", list);
            List<Employees> empList = new List<Employees>();
            empList = (from DataRow dr in dt.Rows
                       select new Employees()
                       {
                           EmployeeId = dr["EmployeeID"].ToString(),
                           EmployeeCode = dr["EmployeeCode"].ToString(),
                           EmployeeName = dr["EmployeeName"].ToString(),
                           EmployeeEmail = dr["Email"].ToString(),
                           EmployeeImage = dr["EmployeeImage"].ToString(),
                           EmployeeDepartment = dr["EmployeeDepartment"].ToString(),
                           EmployeeDesignation = dr["EmployeeDesignation"].ToString(),
                           EmployeeBranch = dr["EmployeeBranch"].ToString(),
                       }).ToList();
            return Json(empList, JsonRequestBehavior.AllowGet);
        }

        public ActionResult _EmpListViaBranchDisplay()
        {
            SortedList list = new SortedList();
            list.Add("@BranchId", Request.Form["BranchId"].ToString());
            DataTable dt = comfun.fillDataTable("sp_EmpListViaBranchId", "", list);
            List<Employees> empList = new List<Employees>();
            empList = (from DataRow dr in dt.Rows
                       select new Employees()
                       {
                           EmployeeId = dr["EmployeeID"].ToString(),
                           EmployeeCode = dr["EmployeeCode"].ToString(),
                           EmployeeName = dr["EmployeeName"].ToString(),
                           EmployeeEmail = dr["Email"].ToString(),
                           EmployeeImage = dr["EmployeeImage"].ToString(),
                           EmployeeDepartment = dr["EmployeeDepartment"].ToString(),
                           EmployeeDesignation = dr["EmployeeDesignation"].ToString(),
                           EmployeeBranch = dr["EmployeeBranch"].ToString(),
                       }).ToList();
            return Json(empList, JsonRequestBehavior.AllowGet);
        }

        public JsonResult _FreezeSaveUpdateSubmit()
        {
            comfun.saveformname("_FreezeSaveUpdateSubmit", "/TimeSheet/_FreezeSaveUpdateSubmit", "Attendance Management/FreezeAttendance&nbsp;", "Freeze Attendance Add", "N", "FreezeAttendance", "FreezeAttendance", "Add", 2);
            comfun.saveformname("_FreezeSaveUpdateSubmit", "/TimeSheet/_FreezeSaveUpdateSubmit", "Attendance Management/FreezeAttendance&nbsp;", "Freeze Attendance Edit", "N", "FreezeAttendance", "FreezeAttendance", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                string fromDate = Request.Form["FromDate"].ToString();
                int startYear = Convert.ToDateTime(fromDate).Year;
                int startMonth = Convert.ToDateTime(fromDate).Month;

                DateTime startDate = new DateTime(startYear, startMonth, 1);
                DateTime endDate = startDate.AddMonths(1).AddDays(-1);

                list.Add("@MonthYearId", startMonth.ToString() + startYear.ToString());
                list.Add("@LockDate", endDate.ToString("dd-MMM-yyyy"));


                list.Add("@CreatedDate", comfun.dateISTstr());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                list.Add("@Type", Request.Form["Type"].ToString());
                list.Add("@EmpId", Request.Form["Emps"].ToString());
                list.Add("@EmpIdexclude", Request.Form["Emps1"].ToString());

                mes = comfun.executeNonQueryWMessage("EmployeeAttendance_FreezeSubmit", "", list).ToString();


            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }

        public JsonResult _UnFreezeSaveUpdateSubmit()
        {
            comfun.saveformname("_FreezeSaveUpdateSubmit", "/TimeSheet/_FreezeSaveUpdateSubmit", "Attendance Management/FreezeAttendance&nbsp;", "Freeze Attendance Add", "N", "FreezeAttendance", "FreezeAttendance", "Add", 2);
            comfun.saveformname("_FreezeSaveUpdateSubmit", "/TimeSheet/_FreezeSaveUpdateSubmit", "Attendance Management/FreezeAttendance&nbsp;", "Freeze Attendance Edit", "N", "FreezeAttendance", "FreezeAttendance", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                string fromDate = Request.Form["FromDate"].ToString();
                int startYear = Convert.ToDateTime(fromDate).Year;
                int startMonth = Convert.ToDateTime(fromDate).Month;

                DateTime startDate = new DateTime(startYear, startMonth, 1);
                DateTime endDate = startDate.AddMonths(1).AddDays(-1);

                list.Add("@MonthYearId", startMonth.ToString() + startYear.ToString());
                list.Add("@LockDate", endDate.ToString("dd-MMM-yyyy"));


                list.Add("@CreatedDate", comfun.dateISTstr());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                list.Add("@Type", "UnFreeze");
                list.Add("@EmpId", Request.Form["Emps"].ToString());
                list.Add("@EmpIdexclude", Request.Form["Emps1"].ToString());
                mes = comfun.executeNonQueryWMessage("EmployeeAttendance_FreezeSubmit", "", list).ToString();


            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        public ActionResult EmpTermination()
        {
            comfun.saveformname("EmpTermination", "/TimeSheet/EmpTermination", "Emp Management/EmpTermination", "Emp Termination form", "Y", "EmpTermination", "EmpTermination", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("EmpTermination");
            return View();
        }
        public JsonResult _TerminatedEmpList()
        {
            comfun.saveformname("_TerminatedEmpList", "/TimeSheet/_TerminatedEmpList", "Emp Management/EmpTermination", "Emp Termination form", "N", "EmpTermination", "EmpTermination", "List", 4);

            SortedList list = new SortedList();
            //  list.Add("@FromDate", Convert.ToDateTime(Request.Form["CurrentMonth"]).ToString("dd-MMM-yyyy"));
            DataTable dt = comfun.fillDataTable("sp_TerminatedEmpListNew", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _EmpListviaDepartment()
        {
            SortedList list = new SortedList();
            list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
            list.Add("@BranchId", Request.Form["BranchId"].ToString());
            DataTable dt = comfun.fillDataTable("sp_EmpListViaDep_Branch_DesId", "", list);
            //DataTable dt = comfun.fillDataTable("stpOutDoorEmployeeList", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _EmpDetailsviaEmpId()
        {
            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            DataTable dt = comfun.fillDataTable("sp_EmpDetailsviaEmpId", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _EmpStatusTerminationList()
        {
            DataTable dt = comfun.fillDataTable("sp_EmpStatusListDisplay", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult EmpTerminationSubmit()
        {
            comfun.saveformname("EmpTerminationSubmit", "/TimeSheet/EmpTerminationSubmit", "Emp Management/EmpTermination", "Emp Termination form", "N", "EmpTermination", "EmpTermination", "Add", 2);
            comfun.saveformname("EmpTerminationSubmit", "/TimeSheet/EmpTerminationSubmit", "Emp Management/EmpTermination", "Emp Termination form", "N", "EmpTermination", "EmpTermination", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@TerminationId", 0);
                list.Add("@EmpId", Request.Form["EmpId"].ToString());
                list.Add("@TerminationDate", Request.Form["LastWorkingDate"].ToString());
                list.Add("@ResignationDate", Request.Form["ResignationDate"].ToString());
                list.Add("@TerminationReason", Request.Form["TerminationReason"].ToString());
                list.Add("@LoggedInUserId", Session["EmpId"]);
                list.Add("@StatusId", Request.Form["StatusId"].ToString());
                // list.Add("@StausId",Request.Form["StatusId"])
                mes = comfun.executeNonQueryWMessage("sp_EmpTermination_Save", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
            }
            return Json(mes);

        }
        public JsonResult _CompanyBranchList()
        {
            DataTable dt = comfun.fillDataTable("sp_CompanyBranchList", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _DepartmentTerminationList()
        {
            SortedList list = new SortedList();
            DataSet ds = comfun.fillDataSet("stpOutdoorDepartmentddl", "", null);
            var json = JsonConvert.SerializeObject(ds.Tables[0]);
            return Json(json);
        }
        public ActionResult CreateUser()
        {
            comfun.saveformname("CreateUser", "/TimeSheet/CreateUser", "Emp Management/CreateUser", "Create User form", "Y", "CreateUser", "CreateUser", "View", 1);
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
            list.Add("@CustomerId", Session["Customer"]);
            dt = comfun.fillDataTable("stpCustomerforCreateUser", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _CreateUserList()
        {
            comfun.saveformname("_CreateUserList", "/TimeSheet/_CreateUserList", "Emp Management/CreateUser", "Create User List", "N", "CreateUser", "CreateUser", "List", 4);

            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("stpEmployeeUser_list", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _CreateUserSubmit()
        {
            comfun.saveformname("_CreateUserSubmit", "/TimeSheet/_CreateUserSubmit", "Emp Management/Create User", "Create User  Add", "N", "CreateUser", "CreateUser", "Add", 2);
            comfun.saveformname("_CreateUserSubmit", "/TimeSheet/_CreateUserSubmit", "Emp Management/Create User", "Create User  Edit", "N", "CreateUser", "CreateUser", "Edit", 3);

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                SqlConnection conLogin = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString);
                Random rnd = new Random();
                string password = rnd.Next(1111, 9999).ToString();
                list.Add("@DomainId", "0");
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@FirstName", Request.Form["FirstName"].ToString());
                list.Add("@LastName", Request.Form["LastName"].ToString());
                list.Add("@RoleId", Request.Form["RoleId"].ToString());
                list.Add("@Email", Request.Form["Email"].ToString());
                list.Add("@CustomerId", Request.Form["CustomerId"].ToString());
                list.Add("@Mobile", Request.Form["Mobile"].ToString());
                list.Add("@Password", password);
                mes = comfun.executeNonQueryWMessage("stpwfmsSignUp_AddDomainUser", "", list, conLogin).ToString();
                if (!mes.Contains("Error"))
                {
                    if (!mes.Contains("NoEmail"))
                    {
                        payfun.SendEmail(Request.Form["FirstName"].ToString() + " " + Request.Form["LastName"].ToString(), Request.Form["Email"].ToString(), "User Name Password", mes);
                    }
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
            comfun.saveformname("_CreateUserStatusUpdate", "/TimeSheet/_CreateUserStatusUpdate", "Emp Management/CreateUser", "Create User Status", "N", "CreateUser", "CreateUser", "Status", 5);

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
        /* Emplyee HPSEDC */

        public ActionResult EmpCityselect()
        {

            SortedList list = new SortedList();
            list.Add("@DistrictId", Request.Form["fiDistrictId"].ToString());
            DataTable dt = comfun.fillDataTable("City_SelectByDistrictId", "", list);
            return PartialView("_city_ddl", dt);

        }
        public ActionResult EmployeeList()
        {
            comfun.saveformname("Employee", "/master/EmployeeList", "Emp Management/Employee", "Employee form", "Y", "Employee", "Employee", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("Employee");
            if (payfun.sessionRecreate() == "expires")
            {
                // comfun.SendEmail("Error-  Expire", "jagvir.singh@horizontelecom.in", "redirecting from expire home", "");
                return RedirectToAction("login", "account");
            }
            DataSet ds = comfun.fillDataSet("stpEmployeedll", "", null);
            return View(ds);
        }
        public ActionResult _AddEmpSubmit()
        {
            comfun.saveformname("_AddEmpSubmit", "/master/_AddEmpSubmit", "Emp Management/Employee", "Employee Add", "N", "Employee", "Employee", "Add", 2);
            comfun.saveformname("_AddEmpSubmit", "/master/_AddEmpSubmit", "Emp Management/Employee", "Employee Edit", "N", "Employee", "Employee", "Edit", 3);

            string mes = string.Empty;
            SortedList list = new SortedList();

            try
            {
                if (payfun.sessionRecreate() == "expires")
                {
                    return PartialView("_sessionExpired");
                }
                string Emp_image = "";
                string _comPath = "";
                if (System.Web.HttpContext.Current.Request.Files.AllKeys.Any())
                {
                    var EmpImage = System.Web.HttpContext.Current.Request.Files["EmpImage"];
                    if (EmpImage.ContentLength > 0)
                    {
                        var fileName = Path.GetFileName(EmpImage.FileName);
                        var _ext = Path.GetExtension(EmpImage.FileName);
                        Emp_image = Guid.NewGuid().ToString();
                        _comPath = Server.MapPath("/EmpImg/") + Emp_image + _ext;
                        //  var _comPath1 = Server.MapPath("D://Webserver//demoHtis//EmpImg//") + Emp_image + _ext;

                        Emp_image = Emp_image + _ext;

                        ViewBag.Msg = _comPath;
                        var path = _comPath;

                        // Saving Image in Original Mode
                        EmpImage.SaveAs(path);
                        // resizing image
                        MemoryStream ms = new MemoryStream();
                        WebImage img = new WebImage(_comPath);

                        if (img.Width > 800)
                            img.Resize(800, 800);
                        img.Save(_comPath);
                        // end resize
                    }
                }

                list.Add("@DomainId", Request.Cookies["DomainId"].Value);
                list.Add("@fvEmployeeCode", Request.Form["fvEmployeeCode"].ToString());
                list.Add("@fiBoiMatCode", Request.Form["fiBoiMatCode"].ToString());
                list.Add("@fiCoveyancemodeId", Request.Form["fiCoveyancemodeId"].ToString());
                list.Add("@fiConveyanceManualId", Request.Form["fiConveyanceManualId"].ToString());
                list.Add("@FirstName", Request.Form["FirstName"].ToString());
                list.Add("@LastName", Request.Form["LastName"].ToString());
                list.Add("@fvMobile", Request.Form["fvMobile"].ToString());
                list.Add("@fvOfficialEmail", Request.Form["fvOfficialEmail"].ToString());
                list.Add("@fiMaritalStatusId", Request.Form["fiMaritalStatusId"].ToString());
                list.Add("@fiGenderId", Request.Form["fiGenderId"].ToString());
                list.Add("@fiCountryId", "1");
                list.Add("@fiStateId", Request.Form["fiStateId"].ToString());
                list.Add("@fiCityId", Request.Form["fiCityId"].ToString());
                list.Add("@fiBranchID", Request.Form["fiBranchID"].ToString());
                list.Add("@fiDepartmentID", Request.Form["fiDepartmentID"].ToString());
                list.Add("@fiDesignationID", Request.Form["fiDesignationID"].ToString());
                list.Add("@fiCategoryTypeId", Request.Form["fiCategoryTypeId"].ToString());
                list.Add("@fiGradeId", Request.Form["fiGradeId"].ToString());
                list.Add("@fiEmpCategoryId", Request.Form["fiEmpCategoryId"].ToString());
                list.Add("@fvCardNo", Request.Form["fvCardNo"].ToString());
                list.Add("@fdDateofcomfirmation", Request.Form["fdDateofcomfirmation"].ToString());
                list.Add("@fdDateofJoining", Request.Form["fdDateofJoining"].ToString());
                list.Add("@fiRoleId", Request.Form["fiRoleId"].ToString());
                list.Add("@fiStatusID", Request.Form["fiStatusID"].ToString());
                list.Add("@fvAadhaarNo", Request.Form["fvAadhaarNo"].ToString());
                list.Add("@fiBloodGroupId", Request.Form["fiBloodGroupId"].ToString());
                list.Add("@fvFatherHusbandName", Request.Form["fvFatherHusbandName"].ToString());
                list.Add("@fvMotherName", Request.Form["fvMotherName"].ToString());
                list.Add("@fvPersonalEmail", Request.Form["fvPersonalEmail"].ToString());
                list.Add("@fvEmergencyContact", Request.Form["fvEmergencyContact"].ToString());
                list.Add("@fvEmergencyPhoneno", Request.Form["fvEmergencyPhoneno"].ToString());
                list.Add("@fvEmergencyRelation", Request.Form["fvEmergencyRelation"].ToString());
                list.Add("@fvIsAdmin", Request.Form["fvIsAdmin"].ToString());
                list.Add("@fvIsHOD", Request.Form["fvIsHOD"].ToString());
                list.Add("@fvIsMobile", Request.Form["fvIsMobile"].ToString());
                list.Add("@fdDateofmarriage", Request.Form["fdDateofmarriage"].ToString());
                list.Add("@fvMedicalcardNo", Request.Form["fvMedicalcardNo"].ToString());
                list.Add("@fvAddress1", Request.Form["fvAddress1"].ToString());
                list.Add("@fvAddressP1", Request.Form["fvAddressP1"].ToString());
                list.Add("@fvIsActive", Request.Form["fvIsActive"].ToString());
                list.Add("@fdDateofBirth", Request.Form["fdDateofBirth"].ToString());
                list.Add("@fvBirthPlace", Request.Form["fvBirthPlace"].ToString());
                list.Add("@Nominee1", Request.Form["Nominee1"].ToString());
                list.Add("@Nominee2", Request.Form["Nominee2"].ToString());
                list.Add("@fvPANNo", Request.Form["fvPANNo"].ToString());
                list.Add("@fvUANNo", Request.Form["fvUANNo"].ToString());
                list.Add("@fvESINo", Request.Form["fvESINo"].ToString());
                list.Add("@PPIN", Request.Form["PPIN"].ToString());
                list.Add("@fvPhone", Request.Form["fvPhone"].ToString());
                list.Add("@fibankbranchid", Request.Form["fiBankBranchId"].ToString());
                list.Add("@fibankid", Request.Form["fiBankId"].ToString());
                list.Add("@fvBankAccountNo", Request.Form["fvBankAccountNo"].ToString());
                list.Add("@MarkAttendance", Request.Form["MarkAttendance"].ToString());
                list.Add("@AutoAttendance", Request.Form["AutoAttendance"].ToString());
                list.Add("@fvIsOTAllow", Request.Form["IsOverTime"].ToString());
                list.Add("@fvIsCalSalaryOnWorkingDays", Request.Form["calonWorking"].ToString());
                list.Add("@AttenanceImage", Request.Form["AttenanceImage"].ToString());
                list.Add("@fiReportingManager", Request.Form["fiReportingManager"].ToString());
                list.Add("@fvPFNo", Request.Form["fvPFNo"].ToString());
                list.Add("@ContractorId", Request.Form["ContractorId"].ToString());
                list.Add("@CircleId", Request.Form["Circle"].ToString());
                list.Add("@LocationId", Request.Form["Location"].ToString());
                list.Add("@SubLocationId", Request.Form["SubLocationId"].ToString());
                list.Add("@fiCustomerId", Request.Form["Customer"].ToString());
                list.Add("@InterCity", Request.Form["InterCity"].ToString());
                list.Add("@fvIFSCCode", Request.Form["Ifsccode"].ToString());
                list.Add("@fiEntityId", Request.Form["EntityId"].ToString());
                list.Add("@fvPanchayatName", Request.Form["PanchayatName"].ToString());
                list.Add("@fiDistrictId", Request.Form["fiDistrictId"].ToString());
                list.Add("@EmpImage", Emp_image);
                list.Add("@EmpImagePath", _comPath);
                if (Request.Form["Id"].ToString() == "0")
                {
                    mes = comfun.executeNonQueryWMessage("stpEmployee_Accept1", "", list).ToString();

                }
                else
                {

                    list.Add("@Id", Request.Form["Id"].ToString());

                    mes = comfun.executeNonQueryWMessage("stpEmployee_Edit1", "", list).ToString();

                }
            }
            catch (Exception ex)
            {
                mes = "Error:" + ex.Message;
            }
            return Json(mes);

        }
        public ActionResult _EmployeeList()
        {
            comfun.saveformname("_EmployeeList", "/master/_EmployeeList", "Emp Management/Employee", "Employee List", "N", "Employee", "Employee", "List", 4);
            //DataTable dt = new DataTable();
            //dt = comfun.fillDataTable("stpEmployeeList", "", null); 
            //var json = JsonConvert.SerializeObject(dt);
            //return Json(json);

            decimal total_records = 0;
            decimal pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            decimal pagesize = Convert.ToDecimal(Request.Form["PageSize"]);
            SortedList list = new SortedList();
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            list.Add("@EmpName", Request.Form["EmpName"].ToString());
            list.Add("@CustomerId", Session["Customer"]);

            //  list.Add("@Empcode", Request.Form["EmpName"].ToString());
            DataTable dt = comfun.fillDataTable("stpEmployeeList", "", list);

            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            //string paging = comfun.create_paging(total_records, pagesize, pageno, "home", "Index");
            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "master", "_Employeelist", "_Employeelist");
            HtmlString htm = new HtmlString(paging);
            ViewData["paging"] = htm;
            var json = JsonConvert.SerializeObject(dt);
            return Json(new
            {
                Paging = paging,
                data = json
            });
            //return View("IndexBlog", ds);
        }
        public ActionResult IFSCCode()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@txt", Request.Form["Ifsccode"].ToString());
            dt = comfun.fillDataTable("autoSearch_BankBranch", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        //public ActionResult _Customerddl()
        //{

        //    DataTable dt = new DataTable();
        //    SortedList list = new SortedList();

        //    dt = comfun.fillDataTable("stpCustomerSelect", "", null);
        //    //return PartialView("_Employeelist", dt);
        //    var json = JsonConvert.SerializeObject(dt);
        //    return Json(json);

        //}


        public ActionResult _CustomerReportingList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@CategoryId", Request.Form["CategoryId"].ToString());
            dt = comfun.fillDataTable("stpCustomerReportingManagerddl", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }


        public ActionResult _employeeEdit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                // comfun.SendEmail("Error-  Expire", "jagvir.singh@horizontelecom.in", "redirecting from expire home", "");
                return RedirectToAction("login", "account");
            }
            DataSet dt = new DataSet();
            string Id = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@Id", Id);
            dt = comfun.fillDataSet("stpEmployeeSelectbyId", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
            //return PartialView("_EditEmployee", dt);
        }
        public ActionResult _EmpCategory()
        {
            DataTable dt = new DataTable();


            dt = comfun.fillDataTable("sp_EmpCategoryListDisplay", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
            //return PartialView("_EditEmployee", dt);
        }

        public ActionResult _EditEmpSubmit()
        {
            comfun.saveformname("_EditEmpSubmit", "/master/_EditEmpSubmit", "Emp Management/Employee", "Employee Add", "N", "Employee", "Employee", "Add", 3);
            comfun.saveformname("_EditEmpSubmit", "/master/_EditEmpSubmit", "Emp Management/Employee", "Employee Edit", "N", "Employee", "Employee", "Edit", 4);
            if (payfun.sessionRecreate() == "expires")
            {
                // comfun.SendEmail("Error-  Expire", "jagvir.singh@horizontelecom.in", "redirecting from expire home", "");
                return RedirectToAction("login", "account");
            }
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {


                string Emp_image = "";
                if (System.Web.HttpContext.Current.Request.Files.AllKeys.Any())
                {
                    var EmpImage = System.Web.HttpContext.Current.Request.Files["EmpImage"];
                    if (EmpImage.ContentLength > 0)
                    {
                        var fileName = Path.GetFileName(EmpImage.FileName);
                        var _ext = Path.GetExtension(EmpImage.FileName);
                        Emp_image = Guid.NewGuid().ToString();
                        var _comPath = Server.MapPath("/EmpImg/") + Emp_image + _ext;
                        Emp_image = Emp_image + _ext;

                        ViewBag.Msg = _comPath;
                        var path = _comPath;

                        // Saving Image in Original Mode
                        EmpImage.SaveAs(path);//---Old
                        //if (commonFunctions.isLocalHost())
                        //{
                        //    path = System.Configuration.ConfigurationManager.AppSettings["BaseDrivePathLocal"].ToString();
                        //}
                        //else
                        //{
                        //    path = System.Configuration.ConfigurationManager.AppSettings["BaseDrivePath"].ToString();

                        //}
                        EmpImage.SaveAs(path + @"\" + Emp_image + _ext);
                        // resizing image
                        MemoryStream ms = new MemoryStream();
                        WebImage img = new WebImage(_comPath);

                        if (img.Width > 800)
                            img.Resize(800, 800);
                        img.Save(_comPath);
                        // end resize
                    }
                }
                list.Add("@DomainId", Request.Cookies["DomainId"].Value);
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@employeeId", Request.Form["employeeId"].ToString());
                list.Add("@empDevice", Request.Form["empDevice"].ToString());
                list.Add("@conveyancemode", Request.Form["conveyancemode"].ToString());
                list.Add("@manual", Request.Form["manual"].ToString());
                list.Add("@FirstName", Request.Form["FirstName"].ToString());
                list.Add("@lastName", Request.Form["lastName"].ToString());
                list.Add("@phone", Request.Form["phone"].ToString());
                list.Add("@OEmail", Request.Form["OEmail"].ToString());
                list.Add("@mStatus", Request.Form["mStatus"].ToString());
                list.Add("@gender", Request.Form["gender"].ToString());
                list.Add("@Country", Request.Form["Country"].ToString());
                list.Add("@State", Request.Form["State"].ToString());
                list.Add("@Branch", Request.Form["Branch"].ToString());
                list.Add("@Deptt", Request.Form["Deptt"].ToString());
                list.Add("@desi", Request.Form["desi"].ToString());
                list.Add("@Type", Request.Form["Type"].ToString());
                list.Add("@Grade", Request.Form["Grade"].ToString());
                list.Add("@EmpCategory", Request.Form["EmpCategory"].ToString());
                list.Add("@cardNo", Request.Form["cardNo"].ToString());
                list.Add("@DOC", Request.Form["DOC"].ToString());
                list.Add("@DOJ", Request.Form["DOJ"].ToString());
                list.Add("@Role", Request.Form["Role"].ToString());
                list.Add("@Status", Request.Form["Status"].ToString());
                list.Add("@ANo", Request.Form["ANo"].ToString());
                list.Add("@bgroup", Request.Form["bgroup"].ToString());
                list.Add("@Fname", Request.Form["Fname"].ToString());
                list.Add("@MName", Request.Form["MName"].ToString());
                list.Add("@PEmail", Request.Form["PEmail"].ToString());
                list.Add("@EContact", Request.Form["EContact"].ToString());
                list.Add("@EPhone", Request.Form["EPhone"].ToString());
                list.Add("@relation", Request.Form["relation"].ToString());
                list.Add("@isadmin", Request.Form["isadmin"].ToString());
                list.Add("@IsHod", Request.Form["IsHod"].ToString());
                list.Add("@Ismoble", Request.Form["Ismoble"].ToString());
                list.Add("@DOW", Request.Form["DOW"].ToString());
                list.Add("@DOB", Request.Form["DOB"].ToString());
                list.Add("@MCN", Request.Form["MCN"].ToString());
                list.Add("@raddress", Request.Form["raddress"].ToString());
                list.Add("@paddress", Request.Form["paddress"].ToString());
                list.Add("@City", Request.Form["City"].ToString());
                list.Add("@birthplace", Request.Form["birthplace"].ToString());
                list.Add("@Nominee1", Request.Form["Nominee1"].ToString());
                list.Add("@Nominee2", Request.Form["Nominee2"].ToString());
                list.Add("@PAN", Request.Form["PAN"].ToString());
                list.Add("@UAN", Request.Form["UAN"].ToString());
                list.Add("@ESINO", Request.Form["ESINO"].ToString());
                list.Add("@PinCode", Request.Form["PinCode"].ToString());
                list.Add("@Phone2", Request.Form["Phone2"].ToString());
                list.Add("@branchid", Request.Form["branchid"].ToString());
                list.Add("@bankid", Request.Form["bankid"].ToString());
                list.Add("@AccountNo", Request.Form["AccountNo"].ToString());
                list.Add("@MarkAttendance", Request.Form["MarkAttendance"].ToString());
                list.Add("@AutoAttendance", Request.Form["AutoAttendance"].ToString());
                list.Add("@AttenanceImage", Request.Form["AttenanceImage"].ToString());
                list.Add("@RM", Request.Form["RM"].ToString());
                list.Add("@EmpImg", Emp_image);
                list.Add("@EmpPath", Emp_image);
                mes = comfun.executeNonQueryWMessage("stpEmployee_Edit", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = "Error:" + ex.Message;
            }
            return Json(mes);

        }
        public JsonResult _EmployeeStatusUpdate()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpEmployeeStatusUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }
        public ActionResult ddlBanKBranchselect()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@BanKId", Request.Form["BankId"].ToString());
            dt = comfun.fillDataTable("stpBankBranchSelectById", "", list);

            return PartialView("ddlBanKBranchselect", dt);
        }
        public ActionResult TimeSheetAttendance()

        {
            comfun.saveformname("TimeSheetAttendance", "/TimeSheet/TimeSheetAttendance", "Attendance Management/TimeSheetAttendance", "Time Sheet Attendance form", "Y", "TimeSheetAttendance", "TimeSheetAttendance", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("TimeSheetAttendance");

            if (payfun.sessionRecreate() == "expires")
            {
                return RedirectToAction("Login", "account");
            }
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            //SortedList list = new SortedList(); 
            ////***************Parameters*******************************************
            //DateTime startDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            //DateTime endDate = startDate.AddMonths(1).AddDays(-1); 
            //list.Add("@FromDate", Convert.ToDateTime(startDate).ToString("dd-MMM-yyyy"));
            //list.Add("@ToDate", Convert.ToDateTime(endDate).ToString("dd-MMM-yyyy"));
            //list.Add("@loginId", Convert.ToInt32(Session["EmpId"].ToString()));
            //list.Add("@Role", Enum.GetName(typeof(payrollFunctions.roleIds), Convert.ToInt32(Session["RoleId"].ToString())));
            //DataTable dt = comfun.fillDataTable("Attendance_SalaryTags", "", list);
            return View();
        }
        public JsonResult encryptString()
        {
            return Json(comfun.encryptString(Request.Form["str"].ToString()));
        }
        public ActionResult _TimeSheetAttendance(FormCollection form)
        {
            comfun.saveformname("_TimeSheetAttendance", "/TimeSheet/_TimeSheetAttendance", " Attendance Management /Monthly Time Sheet", "Attendance Correction form", "Y", "TimeSheetAttendance", "TimeSheetAttendance", "List", 4);
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}

            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@DepartmentId", Request.Form["ScrumId"].ToString());
            list.Add("@BranchId", Request.Form["FunnelId"].ToString());
            list.Add("@StartDate", Request.Form["FromDate"].ToString());
            list.Add("@EndDate", Request.Form["ToDate"].ToString());
            list.Add("@LogInId", Session["EmpId"].ToString());
            list.Add("@Tag", Session["Tag"].ToString());
            DataTable dt = comfun.fillDataTable("stpEmployeeAttendance_Monthly", "", list);

            //DataTable dt = comfun.fillDataTable("Attendance_SalaryTags", "", list);
            /*
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "Attendance", "_salaryAttMonthly", "_salaryAttMonthly");

            HtmlString htm = new HtmlString(paging);
            */
            ViewData["paging"] = "";

            int startMonth = DateTime.Now.Month;
            int startYear = DateTime.Now.Year;
            ViewData["FirstDay"] = Convert.ToDateTime(Request.Form["FromDate"]).Day;
            ViewData["LastDay"] = Convert.ToDateTime(Request.Form["ToDate"]).Day;

            ViewData["StartYear"] = Convert.ToDateTime(Request.Form["FromDate"]).Year;
            ViewData["StartMonth"] = Convert.ToDateTime(Request.Form["FromDate"]).Month;
            ViewData["EmpId"] = "1";// Request.Form["EmpId"].ToString();

            //DateTime startDate = new DateTime(startYear, startMonth, 1);
            //DateTime endDate = startDate.AddMonths(1).AddDays(-1);

            return PartialView(dt);
        }
        public ActionResult salaryAttMonthly()

        {
            comfun.saveformname("salaryAttMonthly", "/TimeSheet/salaryAttMonthly", "Attendance Management/salaryAttMonthly", "Attendance Correction form", "Y", "salaryAttMonthly", "salaryAttMonthly", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("salaryAttMonthly");

            if (payfun.sessionRecreate() == "expires")
            {
                return RedirectToAction("Login", "account");
            }
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            //SortedList list = new SortedList(); 
            ////***************Parameters*******************************************
            //DateTime startDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            //DateTime endDate = startDate.AddMonths(1).AddDays(-1); 
            //list.Add("@FromDate", Convert.ToDateTime(startDate).ToString("dd-MMM-yyyy"));
            //list.Add("@ToDate", Convert.ToDateTime(endDate).ToString("dd-MMM-yyyy"));
            //list.Add("@loginId", Convert.ToInt32(Session["EmpId"].ToString()));
            //list.Add("@Role", Enum.GetName(typeof(payrollFunctions.roleIds), Convert.ToInt32(Session["RoleId"].ToString())));
            //DataTable dt = comfun.fillDataTable("Attendance_SalaryTags", "", list);
            return View();
        }
        // [customAuthorize(Roles = payrollFunctions.rolePayrollTeam)]
        public ActionResult _salaryAttMonthly(FormCollection form)
        {
            comfun.saveformname("_salaryAttMonthly", "/TimeSheet/_salaryAttMonthly", " Attendance Management / salaryAttMonthly", "Attendance Correction form", "Y", "salaryAttMonthly", "salaryAttMonthly", "List", 4);
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}


            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@DepartmentId", Request.Form["ScrumId"].ToString());
            list.Add("@BranchId", Request.Form["FunnelId"].ToString());
            list.Add("@LogInId", Session["EmpId"].ToString());
            list.Add("@Tag", Session["Tag"].ToString());
            list.Add("@StartDate", Request.Form["FromDate"].ToString());
            list.Add("@EndDate", Request.Form["ToDate"].ToString());
            DataTable dt = comfun.fillDataTable("stpEmployeeAttendance_Monthly", "", list);

            //DataTable dt = comfun.fillDataTable("Attendance_SalaryTags", "", list);
            /*
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);
                 string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "Attendance", "_salaryAttMonthly", "_salaryAttMonthly");
             HtmlString htm = new HtmlString(paging);
            */
            ViewData["paging"] = "";

            int startMonth = DateTime.Now.Month;
            int startYear = DateTime.Now.Year;
            DateTime startDate = new DateTime(startYear, startMonth, 1);
            DateTime endDate = startDate.AddMonths(1).AddDays(-1);
            DateTime LastDate = Convert.ToDateTime(Request.Form["ToDate"].ToString());
            if (LastDate > endDate)
            {
                LastDate = endDate;
            }

            ViewData["FirstDay"] = Convert.ToDateTime(Request.Form["FromDate"]).Day;
            ViewData["LastDay"] = Convert.ToDateTime(LastDate).Day;
            ViewData["StartYear"] = Convert.ToDateTime(Request.Form["FromDate"]).Year;
            ViewData["StartMonth"] = Convert.ToDateTime(Request.Form["FromDate"]).Month;
            ViewData["EmpId"] = Request.Form["EmpId"].ToString();// Request.Form["EmpId"].ToString();
            ViewData["EndDate"] = endDate;
            //DateTime startDate = new DateTime(startYear, startMonth, 1);
            //DateTime endDate = startDate.AddMonths(1).AddDays(-1);

            return PartialView(dt);
        }

        //[customAuthorize(Roles = payrollFunctions.rolePayrollTeam)]
        public JsonResult _changeAttendance()
        {
            string mes = "";
            try
            {
                if (payfun.sessionRecreate() == "expires")
                {
                    return Json("Session expires");
                }
                SortedList list = new SortedList();
                list.Add("@EmpId", Request.Form["EmpId"].ToString());
                list.Add("@Date", Request.Form["Date"].ToString());
                list.Add("@DayTag", Request.Form["DayTag"].ToString());
                list.Add("@AttendanceId", "0");
                list.Add("@fiSessionID", Session["SessionId"].ToString());
                list.Add("@CreatedDate", DateTime.Now.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                mes = comfun.executeNonQueryWMessage("stpEmployee_TimeSheetDirectHit", "", list).ToString();

            }

            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        public ActionResult AttendanceCalendar()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                //  return RedirectToAction("Login", "account");
            }
            if (Request.QueryString["key"] != null)
            {
                //  Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            //comfun.saveformname("AttendanceCalendar", "/Attendance/AttendanceCalendar", "Attendance Calendar", "Attendance Calendar", "Y", "Attendance Calendar", "AttendanceCalendar", "List");
            return View();
        }

        public ActionResult _calendar()
        {
            try
            {
                if (payfun.sessionRecreate() == "expires")
                {
                    return PartialView("_sessionExpired");
                }
                int startMonth = 0;
                int startYear = 0;
                if (Request.Form["StartMonth"] != null)
                {
                    startMonth = Convert.ToInt32(Request.Form["StartMonth"]);
                }
                if (Request.Form["StartYear"] != null)
                {
                    startYear = Convert.ToInt32(Request.Form["StartYear"]);
                }
                ViewData["StartYear"] = startYear;
                ViewData["StartMonth"] = startMonth;
                ViewData["EmpId"] = Request.Form["EmpId"].ToString();

                DateTime startDate = new DateTime(startYear, startMonth, 1);
                DateTime endDate = startDate.AddMonths(1).AddDays(-1);
                DateTime endMDate = startDate.AddMonths(1).AddMonths(1);
                return PartialView(GetAttendance(startDate, endDate, Request.Form["EmpId"].ToString()));
            }
            catch (Exception ex)
            {
                ViewData["Error"] = ex.Message;
                return PartialView("_error");
            }

        }

        private DataTable GetAttendance(DateTime fromDate, DateTime toDate, string empId)
        {
            SortedList list = new SortedList();
            list.Add("@EmpCode", empId);
            list.Add("@EmpId", empId);
            list.Add("@FromDate", fromDate.ToString("dd-MMM-yyyy"));
            list.Add("@ToDate", toDate.ToString("dd-MMM-yyyy"));
            list.Add("@CostCenterId", 1);
            //DataTable dt = comfun.fillDataTable("DeviceLogCMR_OtView", "", list);
            //DataTable dt = comfun.fillDataTable("DeviceLogCMR_CalendarDayStatus", "", list);
            DataTable dt = comfun.fillDataTable("TimeSheetRegularizeCalendar", "", list);
            return dt;
        }

        [customAuthorize(Roles = payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam + "," + payrollFunctions.roleAdmin)]
        public ActionResult timingReport()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return RedirectToAction("Login", "account");
            }
            comfun.saveformname("timingReport", "/attendance/timingReport", "Timing Report", "Timing Report", "Y", "", "timingReport", "List");
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }

            return View();
        }
        public ActionResult _TimingReport(FormCollection form)
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return PartialView("_sessionExpired");
            }
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            ViewData["FromDate"] = DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy");
            ViewData["ToDate"] = DateTime.UtcNow.AddMinutes(330).AddDays(1).ToString("dd-MMM-yyyy");

            //ViewData["UserId"] = "";
            //ViewData["TimeFilter"] = "0";
            //ViewData["TimeFilter"] = Request.Form["TimeFilter"].ToString();
            //ViewData["FromDate"] = Request.Form["FromDate"].ToString();
            //ViewData["ToDate"] = Request.Form["ToDate"].ToString();

            list.Add("@FromDate", Request.Form["FromDate"].ToString());
            list.Add("@ToDate", Request.Form["ToDate"].ToString());
            list.Add("@TimeFilter", Request.Form["TimeFilter"].ToString());
            if (Request.Form["CostCenterId"].ToString() != "0" && Request.Form["CostCenterId"].ToString() != "")
            {
                list.Add("@CostCenterId", Request.Form["CostCenterId"].ToString());
            }
            else
            {
                list.Add("@CostCenterId", Session["CostCenterId"].ToString());
            }
            list.Add("@Role", Enum.GetName(typeof(payrollFunctions.roleIds), Convert.ToInt32(Session["RoleId"].ToString())));
            list = payfun.globalParameterList(form, list);

            dt = comfun.fillDataTable("DeviceLogCMR_TimeFilterM", "", list);

            Session["dtExcel"] = dt;
            return PartialView("_TimingReport", dt);
        }
        public ActionResult _TimingReportSummary(FormCollection form)
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return PartialView("_sessionExpired");
            }
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            ViewData["FromDate"] = DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy");
            ViewData["ToDate"] = DateTime.UtcNow.AddMinutes(330).AddDays(1).ToString("dd-MMM-yyyy");

            if (Request.Form["CostCenterId"].ToString() != "0" && Request.Form["CostCenterId"].ToString() != "")
            {
                list.Add("@CostCenterId", Request.Form["CostCenterId"].ToString());
            }
            else
            {
                list.Add("@CostCenterId", Session["CostCenterId"].ToString());
            }
            list.Add("@FromDate", Request.Form["FromDate"].ToString());
            list.Add("@ToDate", Request.Form["ToDate"].ToString());
            list.Add("@TimeFilter", Request.Form["TimeFilter"].ToString());
            list.Add("@Role", Enum.GetName(typeof(payrollFunctions.roleIds), Convert.ToInt32(Session["RoleId"].ToString())));
            list = payfun.globalParameterList(form, list);

            dt = comfun.fillDataTable("Emp_LateSummaryAll", "", list);

            //Session["dtExcel"] = dt;
            return PartialView("_TimeSheetSummary", dt);
        }
        //Leave Management//

        //ApprovalofAttendance//

        public ActionResult ApprovalofAttendance()

        {
            comfun.saveformname("ApprovalofAttendance", "/TimeSheet/ApprovalofAttendance", "Attendance Management/ApprovalofAttendance", "Attendance Correction form", "Y", "ApprovalofAttendance", "ApprovalofAttendance", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("ApprovalofAttendance");

            if (payfun.sessionRecreate() == "expires")
            {
                return RedirectToAction("Login", "account");
            }
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            //SortedList list = new SortedList(); 
            ////***************Parameters*******************************************
            //DateTime startDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            //DateTime endDate = startDate.AddMonths(1).AddDays(-1); 
            //list.Add("@FromDate", Convert.ToDateTime(startDate).ToString("dd-MMM-yyyy"));
            //list.Add("@ToDate", Convert.ToDateTime(endDate).ToString("dd-MMM-yyyy"));
            //list.Add("@loginId", Convert.ToInt32(Session["EmpId"].ToString()));
            //list.Add("@Role", Enum.GetName(typeof(payrollFunctions.roleIds), Convert.ToInt32(Session["RoleId"].ToString())));
            //DataTable dt = comfun.fillDataTable("Attendance_SalaryTags", "", list);
            return View();
        }
        // [customAuthorize(Roles = payrollFunctions.rolePayrollTeam)]
        public ActionResult _ApprovalofAttendance(FormCollection form)
        {
            comfun.saveformname("_ApprovalofAttendance", "/TimeSheet/_ApprovalofAttendance", " Attendance Management / ApprovalofAttendance", "Attendance Correction form", "Y", "ApprovalofAttendance", "ApprovalofAttendance", "List", 4);
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}

            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@DepartmentId", Request.Form["ScrumId"].ToString());
            list.Add("@BranchId", Request.Form["FunnelId"].ToString());
            list.Add("@StartDate", Request.Form["FromDate"].ToString());
            list.Add("@EndDate", Request.Form["ToDate"].ToString());
            DataTable dt = comfun.fillDataTable("stpEmployeeAttendance_Monthly", "", list);

            //DataTable dt = comfun.fillDataTable("Attendance_SalaryTags", "", list);
            /*
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "Attendance", "_salaryAttMonthly", "_salaryAttMonthly");

            HtmlString htm = new HtmlString(paging);
            */
            ViewData["paging"] = "";

            int startMonth = DateTime.Now.Month;
            int startYear = DateTime.Now.Year;
            ViewData["FirstDay"] = Convert.ToDateTime(Request.Form["FromDate"]).Day;
            ViewData["LastDay"] = Convert.ToDateTime(Request.Form["ToDate"]).Day;

            ViewData["StartYear"] = Convert.ToDateTime(Request.Form["FromDate"]).Year;
            ViewData["StartMonth"] = Convert.ToDateTime(Request.Form["FromDate"]).Month;
            ViewData["EmpId"] = "1";// Request.Form["EmpId"].ToString();

            //DateTime startDate = new DateTime(startYear, startMonth, 1);
            //DateTime endDate = startDate.AddMonths(1).AddDays(-1);

            return PartialView(dt);
        }
        public class EmployeeStats
        {
            public string EmpCode { get; set; }
            public string EmpName { get; set; }
            public string Branch { get; set; }
            public string Designation { get; set; }
            public string Department { get; set; }
        }
        public class MonthlyStats
        {
            public string EmpId { get; set; }
            public string AttendanceDate { get; set; }
            public string HolidayDate { get; set; }
            public string AttendanceStatus { get; set; }
            public string TotalAttendance { get; set; }
            public string ApprovalStatus { get; set; }
            public string ApprovedAttendance { get; set; }
            public string IsRejected { get; set; }
            public string IsAbsent { get; set; }
            public string IsWeekOff { get; set; }
            public string IsLeave { get; set; }
            public string IsHoliday { get; set; }
            public string EmpINTime { get; set; }
            public string EmpOUTTime { get; set; }
            public string Title { get; set; }
        }

        public class AttSummary
        {
            public string EmpCode { get; set; }
            public string EmpName { get; set; }
            public string Funnel { get; set; }
            public string Scrum { get; set; }
            public string Designation { get; set; }
            public string DOJ { get; set; }
            public string PP { get; set; }
            public string AA { get; set; }
            public string WW { get; set; }
            public string HH { get; set; }
            public string PayDays { get; set; }
        }

        public class LeaveStats
        {
            public string LeavePeriod { get; set; }
            public string LeaveType { get; set; }
            public string NoOfDays { get; set; }
            public int LeaveStatusId { get; set; }
            public string LeaveStatus { get; set; }
        }
        public class CompOffStats
        {
            public string AttendanceDate { get; set; }
            public string LeaveType { get; set; }
            public string NoOfDays { get; set; }
            public string ShiftName { get; set; }
            public int CompOffStatusId { get; set; }
            public string CompOffStatus { get; set; }
        }
        public class OutdoorStatus
        {
            public string FromDate { get; set; }
            public string ToDate { get; set; }
            public string FromTime { get; set; }
            public string ToTime { get; set; }
            public string IsNightShift { get; set; }
            public int ODStatusId { get; set; }
            public string ODStatus { get; set; }
        }
        public ActionResult TeamStatistics()
        {
            return View();
        }
        public JsonResult _EmpTeam()
        {
            SortedList list = new SortedList();
            list.Add("@LoginID", Session["EmpId"]);
            list.Add("@Tag", Session["Tag"].ToString());
            DataTable dt = comfun.fillDataTable("sp_EmpTeamList", "", list);
            var jsondata = JsonConvert.SerializeObject(dt);
            return Json(jsondata);

        }
        public JsonResult _TeamStatistics(string userid, string dateval)
        {
            List<EmployeeStats> EmpStats = new List<EmployeeStats>();
            List<MonthlyStats> MonthlyStats = new List<MonthlyStats>();
            List<LeaveStats> LeaveStats = new List<LeaveStats>();
            List<CompOffStats> CompOffStats = new List<CompOffStats>();
            List<AttSummary> AttStats = new List<AttSummary>();

            SortedList list = new SortedList();
            list.Add("@LoginID", userid);
            list.Add("@FromDate", Convert.ToDateTime(dateval).ToString("dd/MMM/yyyy"));
            list.Add("@ToDate", Convert.ToDateTime(dateval).AddMonths(1).AddDays(-1).ToString("dd/MMM/yyyy"));
            list.Add("@SessionID", 1);
            DataSet ds = comfun.fillDataSet("stp_TimeSheetTeam_Statistics", "", list);

            SortedList lista = new SortedList();
            lista.Add("@LoginID", Session["EmpId"].ToString());
            lista.Add("@FromDate", Convert.ToDateTime(dateval).ToString("dd-MMM-yyyy"));
            lista.Add("@ToDate", Convert.ToDateTime(dateval).AddMonths(1).AddDays(-1).ToString("dd-MMM-yyyy"));
            lista.Add("@SessionID", 1);

            lista.Add("@Tag", Session["Tag"].ToString());
            DataSet dsa = comfun.fillDataSet("stp_TimeSheetTeam_StatisticsAtt", "", lista);

            //DataSet ds = comfun.fillDataSet("sp_TeamStatistics", "", list);
            EmpStats = (from DataRow dr in ds.Tables[0].Rows
                        select new EmployeeStats()
                        {
                            EmpCode = dr["EmpCode"].ToString(),
                            EmpName = dr["EmpName"].ToString(),
                            Branch = dr["Branch"].ToString(),
                            Designation = dr["Designation"].ToString(),
                            Department = dr["Department"].ToString(),
                        }).ToList();
            MonthlyStats = (from DataRow dr in ds.Tables[1].Rows
                            select new MonthlyStats()
                            {
                                EmpId = dr["EmpID"].ToString(),
                                AttendanceDate = dr["AttDate"].ToString(), // Convert.ToDateTime(dr["AttDate"]).ToString("dd/MMM/yyyy"),
                                HolidayDate = dr["HolidayDate"].ToString(),
                                AttendanceStatus = dr["AttStatus"].ToString(),
                                TotalAttendance = dr["TotalAttendance"].ToString(),
                                ApprovalStatus = dr["PendingApprove"].ToString(),
                                ApprovedAttendance = dr["ApprovedAttendance"].ToString(),
                                IsRejected = dr["Rejected"].ToString(),
                                IsAbsent = dr["Absent"].ToString(),
                                IsWeekOff = dr["WeekOff"].ToString(),
                                IsLeave = dr["Leave"].ToString(),
                                IsHoliday = dr["Holiday"].ToString(),
                                //EmpINTime = Convert.ToDateTime(dr["EmpINTime"]).ToString("hh:mm tt"),
                                //EmpOUTTime = Convert.ToDateTime(dr["EmpOUTTime"]).ToString("hh:mm tt"),
                                Title = dr["Title"].ToString(),
                            }).ToList();
            AttStats = (from DataRow dr in dsa.Tables[0].Rows
                        select new AttSummary()
                        {
                            EmpCode = dr["EmpCode"].ToString(),
                            EmpName = dr["EmpName"].ToString(),
                            Funnel = dr["Funnel"].ToString(),
                            Scrum = dr["Scrum"].ToString(),
                            Designation = dr["Designation"].ToString(),
                            DOJ = dr["Doj"].ToString(),
                            PP = dr["PP"].ToString(),
                            AA = dr["AA"].ToString(),
                            WW = dr["WW"].ToString(),
                            HH = dr["HH"].ToString(),
                            PayDays = dr["PayDays"].ToString()
                        }).ToList();
            //LeaveStats = (from DataRow dr in ds.Tables[1].Rows
            //            select new LeaveStats()
            //            {
            //                LeavePeriod = dr["FromTo"].ToString(),
            //                LeaveType = dr["LeaveType"].ToString(),
            //                NoOfDays = dr["NoOfDays"].ToString(),
            //                LeaveStatusId = Convert.ToInt32(dr["StatusId"]),
            //                LeaveStatus = dr["Status"].ToString(),
            //            }).ToList();
            //CompOffStats = (from DataRow dr in ds.Tables[2].Rows
            //              select new CompOffStats()
            //              {
            //                  AttendanceDate = Convert.ToDateTime(dr["AttendanceDate"]).ToString("dd-MMM-yyyy"),
            //                  LeaveType = dr["Leavetype"].ToString(),
            //                  NoOfDays = dr["NoOfDays"].ToString(),
            //                  ShiftName = dr["ShiftName"].ToString(),
            //                  CompOffStatusId = Convert.ToInt32(dr["StatusId"]),
            //                  CompOffStatus = dr["Status"].ToString(),
            //              }).ToList();
            //ODStats = (from DataRow dr in ds.Tables[3].Rows
            //                select new OutdoorStatus()
            //                {
            //                    FromDate = Convert.ToDateTime(dr["FromDate"]).ToString("dd-MMM-yyyy"),
            //                    ToDate = Convert.ToDateTime(dr["ToDate"]).ToString("dd-MMM-yyyy"),
            //                    FromTime = Convert.ToDateTime(dr["FromTime"]).ToString("hh:mm tt"),
            //                    ToTime = Convert.ToDateTime(dr["ToTime"]).ToString("hh:mm tt"),
            //                    IsNightShift = dr["IsNightShift"].ToString(),
            //                    ODStatusId = Convert.ToInt32(dr["StatusId"]),
            //                    ODStatus = dr["Status"].ToString(),
            //                }).ToList();


            return Json(new
            {
                Status = "Success",
                EmployeeStats = EmpStats,
                MonthlyStats = MonthlyStats,
                AttSummary = AttStats
                //LeaveRequestStats = LeaveStats,
                //CompOffStats = CompOffStats,
                //OutdoorStats = ODStats
            }, JsonRequestBehavior.AllowGet);
        }
        public ActionResult WeekOffList()
        {
            comfun.saveformname("WeekOff", "/TimeSheet/WeekOffList", "Emp Management/Week Off", "Week Off Main Form", "Y", "WeekOff", "WeekOff", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("WeekOff");
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            DataSet ds = new DataSet();
            ds = comfun.fillDataSet("stpEmpWeekOff_ddl", "", null);
            return View(ds);
        }
        public ActionResult _WeekOffList()
        {
            comfun.saveformname("_WeekOffList", "/TimeSheet/_WeekOffList", "Emp Management/Week Off", "Week Off List", "N", "WeekOff", "WeekOff", "List", 4);
            DataTable dt = new DataTable();

            decimal total_records = 0;
            decimal pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            decimal pagesize = 10;

            SortedList list = new SortedList();

            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            list.Add("@EmpName", Request.Form["EmpName"].ToString());

            dt = comfun.fillDataTable("stpWeekOfflist", "", list);

            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            //string paging = comfun.create_paging(total_records, pagesize, pageno, "home", "Index");
            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "attendance", "_WeekOffDays", "_WeekOffDays");
            HtmlString htm = new HtmlString(paging);
            ViewData["paging"] = htm;
            var json = JsonConvert.SerializeObject(dt);
            return Json(new
            {
                Paging = paging,
                data = json
            });
        }
        public ActionResult _WeekOffAdd()
        {
            DataSet dt = new DataSet();
            dt = comfun.fillDataSet("stpEmpWeekOff_ddl", "", null);
            return PartialView("_WeekOffAdd", dt);

        }
        public JsonResult _WeekOffonedayAddSubmit()
        {

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@CategoryID", Request.Form["Category"].ToString());
                list.Add("@branchId", Request.Form["branchId"].ToString());
                list.Add("@WeekOffDate", Request.Form["WeekOffDate"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                list.Add("@fvIsAuto", Request.Form["IsAuto"].ToString());
                list.Add("@SessionId", Request.Form["SessionId"].ToString());
                list.Add("@EmpId", Request.Form["EmpId"].ToString());
                mes = comfun.executeNonQueryWMessage("stpWeekOffOneDay_Accept", "", list).ToString();
            }
            catch (Exception ex)
            {

                mes = ex.Message;
            }



            return Json(mes);
        }
        public ActionResult _weekDays()
        {

            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("stpWeekOffDayName", "", null);
            return PartialView("_weekDays", dt);
        }
        public JsonResult _WeekOffAddSubmit(string id, string categoryId, string branchId, string noofyears, string weekOfData, string IsAuto, string SessionId)
        {
            comfun.saveformname("_WeekOffAddSubmit", "/TimeSheet/_WeekOffAddSubmit", "Emp Management/Week Off", "Week Off Add", "N", "WeekOff", "WeekOff", "Add", 2);

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject(weekOfData);
                DataTable table = new DataTable();
                table.Columns.Add("CategoryId", typeof(int));
                table.Columns.Add("WeekDayName", typeof(string));
                table.Columns.Add("OptionalIds", typeof(string));

                foreach (var item in jsonData)
                {
                    DataRow dr = table.NewRow();
                    dr["CategoryId"] = item.CategoryId;
                    dr["WeekDayName"] = item.WeekDayName;
                    dr["OptionalIds"] = item.OptionalIds;
                    table.Rows.Add(dr);
                }


                SqlCommand com = new SqlCommand();
                connection conObj = new connection();

                com.Connection = conObj.con;

                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "stpWeekOff_AcceptK";
                com.Parameters.AddWithValue("@CategoryID", categoryId);
                com.Parameters.AddWithValue("@BranchID", branchId);

                com.Parameters.AddWithValue("@fvIsAuto", IsAuto);
                com.Parameters.AddWithValue("@NoofYear", noofyears);
                com.Parameters.AddWithValue("@SessionID", SessionId);
                com.Parameters.AddWithValue("@CreatedBy", Request.Cookies["EmpId"].Value);

                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@WeekDay";
                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table;
                com.Parameters.Add(parameter);

                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);


                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();

                com.ExecuteNonQuery();

                conObj.con.Close();

                mes = sp.Value.ToString();

                return Json(new { Message = mes, Status = "success" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Message = "Error:" + ex.Message, Status = "fail" }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult SalaryCalculate()
        {
            comfun.saveformname("SalaryCalculate", "/TimeSheet/SalaryCalculate", "Time Sheet Management/SalaryCalculate", "Salary Calculate form", "Y", "SalaryCalculate", "SalaryCalculate", "View", 1);
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            return View();
        }
        public JsonResult _SalaryCalculation()
        {
            SortedList list = new SortedList();
            string mes = string.Empty;
            try
            {
                if (Request.Form["FromDate"] != null)
                {
                    list.Add("@StartDate", Request.Form["FromDate"].ToString());
                }
                if (Request.Form["ToDate"] != null)
                {
                    list.Add("@EndDate", Convert.ToDateTime(Request.Form["ToDate"]).ToString("dd-MMM-yyyy"));
                }
                //list.Add("@loginId", Convert.ToInt32(Session["EmpId"].ToString()));
                //
                //list.Add("@Role", Enum.GetName(typeof(payrollFunctions.roleIds), Convert.ToInt32(Session["RoleId"].ToString())));
                if (Request.Form["EmpId"] != "")
                {
                    list.Add("@EmpId", Request.Form["EmpId"].ToString());
                }
                if (Request.Form["BranchId"] != "")
                {
                    list.Add("@BranchId", Request.Form["BranchId"].ToString());
                }
                if (Request.Form["DesignationId"] != "")
                {
                    list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
                }
                if (Request.Form["DepartmentId"] != "")
                {
                    list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
                }
                if (payfun.getQueryStringValueOrDefault("CategoryId") != "")
                {
                    list.Add("@CategoryId", Request.QueryString["CategoryId"].ToString());
                }
                if (payfun.getQueryStringValueOrDefault("CircleId") != "")
                {
                    list.Add("@CircleId", Request.QueryString["CircleId"].ToString());
                }
                //list.Add("@Tag", Session["Tag"].ToString());
                mes = comfun.executeNonQueryWMessage("stpSalaryCalculation", "", list).ToString();

            }
            catch (Exception ex)
            {

                mes = "Error:" + ex.Message;
            }

            return Json(mes);
        }
        public ActionResult AttendanceApproval()

        {
            comfun.saveformname("AttendanceApproval", "/TimeSheet/AttendanceApproval", "Attendance Management/AttendanceApproval", "Attendance Approval form", "Y", "AttendanceApproval", "AttendanceApproval", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("AttendanceApproval");

            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return RedirectToAction("Login", "account");
            //}
            //if (Request.QueryString["key"] != null)
            //{
            //    Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            //}
            //SortedList list = new SortedList(); 
            ////***************Parameters*******************************************
            //DateTime startDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            //DateTime endDate = startDate.AddMonths(1).AddDays(-1); 
            //list.Add("@FromDate", Convert.ToDateTime(startDate).ToString("dd-MMM-yyyy"));
            //list.Add("@ToDate", Convert.ToDateTime(endDate).ToString("dd-MMM-yyyy"));
            //list.Add("@loginId", Convert.ToInt32(Session["EmpId"].ToString()));
            //list.Add("@Role", Enum.GetName(typeof(payrollFunctions.roleIds), Convert.ToInt32(Session["RoleId"].ToString())));
            //DataTable dt = comfun.fillDataTable("Attendance_SalaryTags", "", list);
            return View();
        }
        public ActionResult _AttendanceStatusddl()
        {

            DataTable dt = comfun.fillDataTable("StpTimeSheet_AttStatus", "", null);

            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        // [customAuthorize(Roles = payrollFunctions.rolePayrollTeam)] b
        public ActionResult _TimeSheetApproval(FormCollection form)
        {
            comfun.saveformname("_TimeSheetApproval", "/TimeSheet/_TimeSheetApproval", " Attendance Management /TimeSheetApproval", "Attendance Approval form", "Y", "AttendanceApproval", "AttendanceApproval", "List", 4);
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");

            string fromDate = Request.Form["FromDate"].ToString();
            int startYear = Convert.ToDateTime(fromDate).Year;
            int startMonth = Convert.ToDateTime(fromDate).Month;

            DateTime startDate = new DateTime(startYear, startMonth, 1);
            DateTime endDate = startDate.AddMonths(1).AddDays(-1);




            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmployeeID"].ToString());
            list.Add("@DepartmentId", Request.Form["ScrumId"].ToString());
            list.Add("@StartDate", Request.Form["FromDate"].ToString());
            list.Add("@EndDate", Request.Form["ToDate"].ToString());

            DataTable dt = comfun.fillDataTable("stpEmployeeAttendanceApproval_Monthly", "", list);


            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public JsonResult _AttendanceApprovalSubmit(string FreezeEmpIds, string FromDate)
        {
            string mes = string.Empty;
            try
            {
                DataTable table = (DataTable)JsonConvert.DeserializeObject(FreezeEmpIds, (typeof(DataTable)));


                SqlCommand com = new SqlCommand();
                connection conObj = new connection();

                com.Connection = conObj.con;

                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "Attendance_SalaryForTimeLockSubmit";
                SortedList list = new SortedList();
                DateTime monthYear = Convert.ToDateTime(FromDate);
                com.Parameters.AddWithValue("@MonthYearId", monthYear.ToString("MMyyyy"));

                com.Parameters.AddWithValue("@loginID", Session["EmpId"].ToString());
                com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());


                com.Parameters.AddWithValue("@Tag", "E");
                com.Parameters.AddWithValue("@RoleId", Session["RoleId"].ToString());
                com.Parameters.AddWithValue("@LockDate", FromDate);
                com.Parameters.AddWithValue("@CreatedDate", DateTime.Now.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));

                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@AttendanceTable";
                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table;
                com.Parameters.Add(parameter);

                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);


                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();

                com.ExecuteNonQuery();

                conObj.con.Close();

                mes = sp.Value.ToString();

                //return Json(
                //new
                //{
                //    Message = mes
                //},
                //JsonRequestBehavior.AllowGet
                //);
            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
                //return Json(
                //new
                //{
                //    Message = mes
                //},
                //JsonRequestBehavior.AllowGet
                //);

            }
            return Json(mes);
        }
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
        public JsonResult _EmpPunchStatus()
        {
            SortedList list = new SortedList();
            list.Add("@EmpID", Session["EmpId"]);
            DataTable dt = comfun.fillDataTable("stpEmployeePunch_LstStatus", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public string GetAddress(string Latitude, string Longitude)
        {
            var GoogleApiKey = Convert.ToString(ConfigurationManager.AppSettings["Google_API_Key"]);
            string Output = string.Empty;
            string Country = string.Empty;
            string CountryCode = string.Empty;
            string State = string.Empty;
            string City = string.Empty;
            string ZipCode = string.Empty;
            try
            {
                var GetAddress = "https://maps.googleapis.com/maps/api/geocode/xml?latlng=" + Latitude + "," + Longitude + "&key=" + GoogleApiKey;
                var LocData = new System.Net.WebClient().DownloadString(GetAddress);
                var xmlEle = XElement.Parse(LocData);
                string FullAddress = (from ele in xmlEle.Descendants() where ele.Name == "formatted_address" select ele).FirstOrDefault().Value;
                var Result = (from ele in xmlEle.Descendants() where ele.Name == "address_component" select ele);
                foreach (XElement elm in Result)
                {
                    string type = elm.Elements().Where(e => e.Name.LocalName == "type").FirstOrDefault().Value;
                    if (type.ToLower().Trim() == "country")
                    {
                        Country = elm.Elements().Where(e => e.Name.LocalName == "long_name").Single().Value;
                        CountryCode = elm.Elements().Where(e => e.Name.LocalName == "short_name").Single().Value;
                    }
                    if (type.ToLower().Trim() == "administrative_area_level_1")
                    {
                        State = elm.Elements().Where(e => e.Name.LocalName == "long_name").Single().Value;
                    }
                    if (type.ToLower().Trim() == "locality")
                    {
                        City = elm.Elements().Where(e => e.Name.LocalName == "long_name").Single().Value;
                    }
                    else
                    {
                        if (type.ToLower().Trim() == "administrative_area_level_2")
                        {
                            City = elm.Elements().Where(e => e.Name.LocalName == "long_name").Single().Value;
                        }
                    }
                    if (type.ToLower().Trim() == "postal_code")
                    {
                        ZipCode = elm.Elements().Where(e => e.Name.LocalName == "long_name").Single().Value;
                    }
                }
                xmlEle.RemoveAll();
                Output = Country + "~" + State + "~" + City + "~" + ZipCode + "~" + FullAddress;
                return Output;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public JsonResult WebAttendanceSubmit()
        {
            string mes = string.Empty;
            string OutputMsg = string.Empty;
            string StatusMsg = string.Empty;
            try
            {
                //string Latitude = Request.Form["Latitude"].ToString();
                //string Longitude = Request.Form["Longitude"].ToString();
                //string data = GetAddress(Latitude, Longitude);
                //string[] address = data.Split('~');
                //string Country = address[0];
                //string State = address[1];
                //string County = address[2];
                //string PostalCode = address[3];
                //string FullAddress = address[4];
                SortedList list = new SortedList();
                list.Add("@EmpID", Session["EmpId"]);
                list.Add("@StartDate", DateTime.Now.ToString("dd-MMM-yyyy"));
                //list.Add("@StartDate", DateTime.Now.ToString("dd-MMM-yyyy HH:mm:ss"));
                list.Add("@StartTime", DateTime.Now.ToString("dd-MMM-yyyy HH:mm:ss"));
                list.Add("@Lat", "");
                list.Add("@Long", "");
                list.Add("@LocationAddress", "");
                list.Add("@PunchCountry", "");
                list.Add("@PunchState", "");
                list.Add("@PunchCity", "");
                list.Add("@DomainID", Session["DomainId"]);
                list.Add("@PunchType", Request.Form["PunchType"].ToString());
                list.Add("@PunchMode", "W");
                mes = comfun.executeNonQueryWMessage("stpEmployeeDayStart_Accept", "", list).ToString();
                string[] NewMessage = mes.Split('#');
                OutputMsg = NewMessage[0];
                StatusMsg = NewMessage[1];
            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
            }
            return Json(OutputMsg);
        }


        public ActionResult _Roledll()
        {
            DataTable dt = comfun.fillDataTable("StpTimeSheet_Roleddl", "", null);

            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _Applicationdll()
        {
            DataTable dt = comfun.fillDataTable("StpTimeSheet_Applicationddl", "", null);

            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _devsupportdll()
        {
            DataTable dt = comfun.fillDataTable("StpTimeSheet_Supportddl", "", null);

            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }


        //Role
        public ActionResult Role()
        {
            comfun.saveformname("Role", "/TimeSheet/Role", "Master/General/Role", "Role Main form", "Y", "Role", "Role", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("Role");
            return View();
        }
        public ActionResult _RoleList()
        {
            comfun.saveformname("_RoleList", "/TimeSheet/_RoleList", "Master/General/Role", "Role List", "N", "Role", "Role", "List", 4);
            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 50;
            if (Request.Form["Page"] != null)
            {
                pagesize = Convert.ToDecimal(Request.Form["Page"]);
            }
            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }
            SortedList list = new SortedList();
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            DataTable dt = comfun.fillDataTable("Role_Select", "", list);
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "admin", "_RoleList", "_RoleList");

            HtmlString htm = new HtmlString(paging);
            ViewData["paging"] = htm;
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        //[customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult _RoleAdd()
        {

            // comfun.saveformname("DesignationAdd", "/admin/DesignationAdd", "Designation Add", "Designation add view", "N", "DesignationAdd", "Designation", "Add");
            DataTable dt = comfun.fillDataTable("stpViksatPayGradeSelect", "", null);
            return PartialView("_RoleAdd", dt);
        }
        public ActionResult _RoleStatusUpdate()
        {
            comfun.saveformname("_RoleStatusUpdate", "/TimeSheet/_RoleStatusUpdate", "Master/General/Role", "Role Status", "N", "Role", "Role", "Status", 5);

            // comfun.saveformname("DesignationAdd", "/admin/DesignationAdd", "Designation Add", "Designation add view", "N", "DesignationAdd", "Designation", "Add");

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpUpdateStatusRole", "", list).ToString();
            }
            catch (Exception ex)
            {

                mes = ex.Message;
            }

            return Json(mes);
        }
        public JsonResult _RoleAddSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            comfun.saveformname("_RoleEditSubmit", "/TimeSheet/_RoleEditSubmit", "Master/General/Role", "Role Edit", "N", "Role", "Role", "Edit", 3);
            comfun.saveformname("_RoleAddSubmit", "/TimeSheet/_RoleAddSubmit", "Master/General/Role", "Role Add", "N", "Role", "Role", "Add", 2);
            try
            {
                SortedList list = new SortedList();

                //  list.Add("@fiGradeId", Request.Form["GradeID"].ToString());
                list.Add("@RoleCode", Request.Form["RoleCode"].ToString());
                list.Add("@RoleName", Request.Form["RoleName"].ToString());
                //list.Add("@CreatedBy", Session["EmpId"].ToString());
                //list.Add("@CreatedDate", comfun.dateISTstr());
                mes = comfun.executeNonQueryWMessage("Role_Save", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_RoleAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        // [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult _RoleEdit2()
        {

            //comfun.saveformname("designationEdit", "/admin/designationEdit", "Designation Edit", "Designation Edit view", "N", "DesignationEdit", "Designation", "Edit");
            string designationid = "0";
            //if (Request.QueryString["id"] == null)
            //{
            //    return RedirectToAction("Designation");
            //}
            designationid = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@RoleId", designationid);
            DataSet ds = comfun.fillDataSet("Role_SelectWithId", "", list);
            return PartialView(ds);
        }
        public JsonResult _RoleEditSubmit3()
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}
            //comfun.saveformname("_designationEditSubmit", "/admin/_designationEditSubmit", "General/Designation", "Designation Edit", "N", "Designation", "Designation", "Edit");
            comfun.saveformname("_RoleEditSubmit", "/TimeSheet/_RoleEditSubmit", "Master/General/Role", "Role Edit", "N", "Role", "Role", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                //list.Add("@GradeID", Request.Form["GradeID"].ToString());
                list.Add("@RoleName", Request.Form["RoleName"].ToString());
                list.Add("@fvRoleCode", Request.Form["RoleCode"].ToString());
                list.Add("@RoleId", Request.Form["RoleId"].ToString());
                //list.Add("@Isactive", Request.Form["Isactive"].ToString());
                mes = comfun.executeNonQueryWMessage("Role_Edit", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_RoleEditSubmit3", ex.Message);
            }
            return Json(mes);
        }

        //Application
        public ActionResult Application()
        {
            comfun.saveformname("Application", "/TimeSheet/Application", "Master/General/Application", "Application Main form", "Y", "Application", "Application", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("Application");
            return View();
        }
        public ActionResult _ApplicationList()
        {
            comfun.saveformname("_ApplicationList", "/TimeSheet/_ApplicationList", "Master/General/Application", "Application List", "N", "Application", "Application", "List", 4);
            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 50;
            if (Request.Form["Page"] != null)
            {
                pagesize = Convert.ToDecimal(Request.Form["Page"]);
            }
            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }
            SortedList list = new SortedList();
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            DataTable dt = comfun.fillDataTable("Application_Select", "", list);
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "TimeSheet", "_ApplicationList", "_ApplicationList");

            HtmlString htm = new HtmlString(paging);
            ViewData["paging"] = htm;
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        //[customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult _ApplicationAdd()
        {

            // comfun.saveformname("DesignationAdd", "/admin/DesignationAdd", "Designation Add", "Designation add view", "N", "DesignationAdd", "Designation", "Add");
            DataTable dt = comfun.fillDataTable("stpViksatPayGradeSelect", "", null);
            return PartialView("_ApplicationAdd", dt);
        }
        public ActionResult _ApplicationStatusUpdate()
        {
            comfun.saveformname("_ApplicationStatusUpdate", "/TimeSheet/_ApplicationStatusUpdate", "Master/General/Application", "Application Status", "N", "Application", "Application", "Status", 5);

            // comfun.saveformname("DesignationAdd", "/admin/DesignationAdd", "Designation Add", "Designation add view", "N", "DesignationAdd", "Designation", "Add");

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpUpdateStatusApplication", "", list).ToString();
            }
            catch (Exception ex)
            {

                mes = ex.Message;
            }

            return Json(mes);
        }
        public JsonResult _ApplicationAddSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            comfun.saveformname("_ApplicationEditSubmit", "/TimeSheet/_ApplicationEditSubmit", "Master/General/Application", "Application Edit", "N", "Application", "Application", "Edit", 3);
            comfun.saveformname("_ApplicationAddSubmit", "/TimeSheet/_ApplicationAddSubmit", "Master/General/Application", "Application Add", "N", "Application", "Application", "Add", 2);
            try
            {
                SortedList list = new SortedList();

                //  list.Add("@fiGradeId", Request.Form["GradeID"].ToString());
                list.Add("@ApplicationCode", Request.Form["ApplicationCode"].ToString());
                list.Add("@ApplicationName", Request.Form["ApplicationName"].ToString());
                //list.Add("@CreatedBy", Session["EmpId"].ToString());
                //list.Add("@CreatedDate", comfun.dateISTstr());
                mes = comfun.executeNonQueryWMessage("Application_Save", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_ApplicationAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        // [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult _ApplicationEdit2()
        {

            //comfun.saveformname("designationEdit", "/admin/designationEdit", "Designation Edit", "Designation Edit view", "N", "DesignationEdit", "Designation", "Edit");
            string designationid = "0";
            //if (Request.QueryString["id"] == null)
            //{
            //    return RedirectToAction("Designation");
            //}
            designationid = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@ApplicationId", designationid);
            DataSet ds = comfun.fillDataSet("Application_SelectWithId", "", list);
            return PartialView(ds);
        }
        public JsonResult _ApplicationEditSubmit3()
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}
            //comfun.saveformname("_designationEditSubmit", "/admin/_designationEditSubmit", "General/Designation", "Designation Edit", "N", "Designation", "Designation", "Edit");
            comfun.saveformname("_ApplicationEditSubmit", "/TimeSheet/_ApplicationEditSubmit", "Master/General/Application", "Application Edit", "N", "Application", "Application", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                //list.Add("@GradeID", Request.Form["GradeID"].ToString());
                list.Add("@ApplicationName", Request.Form["ApplicationName"].ToString());
                list.Add("@fvApplicationCode", Request.Form["ApplicationCode"].ToString());
                list.Add("@ApplicationId", Request.Form["ApplicationId"].ToString());
                //list.Add("@Isactive", Request.Form["Isactive"].ToString());
                mes = comfun.executeNonQueryWMessage("Application_Edit", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_ApplicationEditSubmit3", ex.Message);
            }
            return Json(mes);
        }

        //DevSupport
        public ActionResult DevSupport()
        {
            comfun.saveformname("DevSupport", "/TimeSheet/DevSupport", "Master/General/DevSupport", "DevSupport Main form", "Y", "DevSupport", "DevSupport", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("DevSupport");
            return View();
        }
        public ActionResult _DevSupportList()
        {
            comfun.saveformname("_DevSupportList", "/TimeSheet/_DevSupportList", "Master/General/DevSupport", "DevSupport List", "N", "DevSupport", "DevSupport", "List", 4);
            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 50;
            if (Request.Form["Page"] != null)
            {
                pagesize = Convert.ToDecimal(Request.Form["Page"]);
            }
            if (Request.Form["PageNo"] != null)
            {

                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }
            SortedList list = new SortedList();
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            DataTable dt = comfun.fillDataTable("DevSupport_Select", "", list);
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "TimeSheet", "_DevSupportList", "_DevSupportList");

            HtmlString htm = new HtmlString(paging);
            ViewData["paging"] = htm;
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        //[customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult _DevSupportAdd()
        {
            // comfun.saveformname("DesignationAdd", "/admin/DesignationAdd", "Designation Add", "Designation add view", "N", "DesignationAdd", "Designation", "Add");
            DataTable dt = comfun.fillDataTable("stpViksatPayGradeSelect", "", null);
            return PartialView("_DevSupportAdd", dt);
        }
        public ActionResult _DevSupportStatusUpdate()
        {
            comfun.saveformname("_DevSupportStatusUpdate", "/TimeSheet/_DevSupportStatusUpdate", "Master/General/DevSupport", "DevSupport Status", "N", "DevSupport", "DevSupport", "Status", 5);

            // comfun.saveformname("DesignationAdd", "/admin/DesignationAdd", "Designation Add", "Designation add view", "N", "DesignationAdd", "Designation", "Add");

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpUpdateStatusDevSupport", "", list).ToString();
            }
            catch (Exception ex)
            {

                mes = ex.Message;
            }

            return Json(mes);
        }
        public JsonResult _DevSupportAddSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            comfun.saveformname("_DevSupportEditSubmit", "/TimeSheet/_DevSupportEditSubmit", "Master/General/DevSupport", "DevSupport Edit", "N", "DevSupport", "DevSupport", "Edit", 3);
            comfun.saveformname("_DevSupportAddSubmit", "/TimeSheet/_DevSupportAddSubmit", "Master/General/DevSupport", "DevSupport Add", "N", "DevSupport", "DevSupport", "Add", 2);
            try
            {
                SortedList list = new SortedList();

                //  list.Add("@fiGradeId", Request.Form["GradeID"].ToString());
                list.Add("@DevSupportCode", Request.Form["DevSupportCode"].ToString());
                list.Add("@DevSupportName", Request.Form["DevSupportName"].ToString());
                //list.Add("@CreatedBy", Session["EmpId"].ToString());
                //list.Add("@CreatedDate", comfun.dateISTstr());
                mes = comfun.executeNonQueryWMessage("DevSupport_Save", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_DevSupportAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        // [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult _DevSupportEdit2()
        {

            //comfun.saveformname("designationEdit", "/admin/designationEdit", "Designation Edit", "Designation Edit view", "N", "DesignationEdit", "Designation", "Edit");
            string designationid = "0";
            //if (Request.QueryString["id"] == null)
            //{
            //    return RedirectToAction("Designation");
            //}
            designationid = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@DevSupportId", designationid);
            DataSet ds = comfun.fillDataSet("DevSupport_SelectWithId", "", list);
            return PartialView(ds);
        }
        public JsonResult _DevSupportEditSubmit3()
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}
            //comfun.saveformname("_designationEditSubmit", "/admin/_designationEditSubmit", "General/Designation", "Designation Edit", "N", "Designation", "Designation", "Edit");
            comfun.saveformname("_DevSupportEditSubmit", "/TimeSheet/_DevSupportEditSubmit", "Master/General/DevSupport", "DevSupport Edit", "N", "DevSupport", "DevSupport", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                //list.Add("@GradeID", Request.Form["GradeID"].ToString());
                list.Add("@DevSupportName", Request.Form["DevSupportName"].ToString());
                list.Add("@fvDevSupportCode", Request.Form["DevSupportCode"].ToString());
                list.Add("@DevSupportId", Request.Form["DevSupportId"].ToString());
                //list.Add("@Isactive", Request.Form["Isactive"].ToString());
                mes = comfun.executeNonQueryWMessage("DevSupport_Edit", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_DevSupportEditSubmit3", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult Calculate()
        {
            comfun.saveformname("Calculate", "/reports/Calculate", "Reports/Calculate", "Report Calculate form", "Y", "ReportCalculate", "ReportCalculate", "View", 1);
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            return View();
        }
        public JsonResult _AttendanceCalculation()
        {
            SortedList list = new SortedList();
            string mes = string.Empty;
            try
            {
                list.Add("@EmpID", Session["EmpId"]);
                list.Add("@StartDate", DateTime.Now.ToString("dd-MMM-yyyy"));
                list.Add("@EndDate", DateTime.Now.ToString("dd-MMM-yyyy"));
                //list.Add("@StartTime", DateTime.Now.ToString("dd-MMM-yyyy HH:mm:ss"));
                //if (Request.Form["DesignationId"] != "")
                //{
                //    list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
                //}
                //if (Request.Form["DepartmentId"] != "")
                //{
                //    list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
                //}
                mes = comfun.executeNonQueryWMessage("stpTimesheet_Calculation", "", list).ToString();

            }
            catch (Exception ex)
            {

                mes = "Error:" + ex.Message;
            }

            return Json(mes);
        }

        //My Team        
        public ActionResult MyTeam()
        {
            comfun.saveformname("MyTeam", "/TimeSheet/MyTeam", "Master/General/MyTeam", "MyTeam Main form", "Y", "MyTeam", "MyTeam", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("DevSupport");
            return View();
        }
        public ActionResult _MyTeamList()
        {
            comfun.saveformname("_MyTeamList", "/TimeSheet/_MyTeamList", "Master/General/MyTeam", "MyTeam List", "N", "MyTeam", "MyTeam", "List", 4);
            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 50;
            if (Request.Form["Page"] != null)
            {
                pagesize = Convert.ToDecimal(Request.Form["Page"]);
            }
            if (Request.Form["PageNo"] != null)
            {

                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }
            SortedList list = new SortedList();
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            DataTable dt = comfun.fillDataTable("MyTeam_Select", "", list);
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "TimeSheet", "_MyTeamList", "_MyTeamList");

            HtmlString htm = new HtmlString(paging);
            ViewData["paging"] = htm;
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        //[customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult _MyTeamAdd()
        {
            // comfun.saveformname("DesignationAdd", "/admin/DesignationAdd", "Designation Add", "Designation add view", "N", "DesignationAdd", "Designation", "Add");
            DataTable dt = comfun.fillDataTable("stpViksatPayGradeSelect", "", null);
            return PartialView("_MyTeamAdd", dt);
        }
        public ActionResult _MyTeamStatusUpdate()
        {
            comfun.saveformname("_MyTeamStatusUpdate", "/TimeSheet/_MyTeamStatusUpdate", "Master/General/MyTeam", "MyTeam Status", "N", "MyTeam", "MyTeam", "Status", 5);

            // comfun.saveformname("DesignationAdd", "/admin/DesignationAdd", "Designation Add", "Designation add view", "N", "DesignationAdd", "Designation", "Add");

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpUpdateStatusMyTeam", "", list).ToString();
            }
            catch (Exception ex)
            {

                mes = ex.Message;
            }

            return Json(mes);
        }
        public JsonResult _MyTeamAddSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            comfun.saveformname("_MyTeamEditSubmit", "/TimeSheet/_MyTeamEditSubmit", "Master/General/MyTeam", "MyTeam Edit", "N", "MyTeam", "MyTeam", "Edit", 3);
            comfun.saveformname("_MyTeamAddSubmit", "/TimeSheet/_MyTeamAddSubmit", "Master/General/MyTeam", "MyTeam Add", "N", "MyTeam", "MyTeam", "Add", 2);
            try
            {
                SortedList list = new SortedList();

                //  list.Add("@fiGradeId", Request.Form["GradeID"].ToString());
                list.Add("@TeamCode", Request.Form["TeamCode"].ToString());
                list.Add("@TeamName", Request.Form["TeamName"].ToString());
                //list.Add("@CreatedBy", Session["EmpId"].ToString());
                //list.Add("@CreatedDate", comfun.dateISTstr());
                mes = comfun.executeNonQueryWMessage("MyTeam_Save", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_MyTeamAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        // [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult _MyTeamEdit2()
        {

            //comfun.saveformname("designationEdit", "/admin/designationEdit", "Designation Edit", "Designation Edit view", "N", "DesignationEdit", "Designation", "Edit");
            string designationid = "0";
            //if (Request.QueryString["id"] == null)
            //{
            //    return RedirectToAction("Designation");
            //}
            designationid = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@TeamId", designationid);
            DataSet ds = comfun.fillDataSet("MyTeam_SelectWithId", "", list);
            return PartialView(ds);
        }
        public JsonResult _MyTeamEditSubmit3()
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}
            //comfun.saveformname("_designationEditSubmit", "/admin/_designationEditSubmit", "General/Designation", "Designation Edit", "N", "Designation", "Designation", "Edit");
            comfun.saveformname("_MyTeamEditSubmit", "/TimeSheet/_MyTeamEditSubmit", "Master/General/MyTeam", "MyTeam Edit", "N", "MyTeam", "MyTeam", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                //list.Add("@GradeID", Request.Form["GradeID"].ToString());
                list.Add("@TeamName", Request.Form["TeamName"].ToString());
                list.Add("@fvTeamCode", Request.Form["TeamCode"].ToString());
                list.Add("@TeamId", Request.Form["TeamId"].ToString());
                //list.Add("@Isactive", Request.Form["Isactive"].ToString());
                mes = comfun.executeNonQueryWMessage("MyTeam_Edit", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_MyTeamEditSubmit3", ex.Message);
            }
            return Json(mes);
        }
        //[customAuthorize(Roles = payrollFunctions.rolePayrollTeam)]
        public ActionResult AttendanceApprovalByDepartment()

        {
            comfun.saveformname("AttendanceApprovalByDepartment", "/TimeSheet/AttendanceApprovalByDepartment", "Attendance Approval By Department", "Attendance Approval By Department", "Y", "Attendance Approval By Department", "AttendanceApprovalByDepartment", "List");


            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return RedirectToAction("Login", "account");
            //}
            //if (Request.QueryString["key"] != null)
            //{
            //    Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            //}
            //SortedList list = new SortedList(); 
            ////***************Parameters*******************************************
            //DateTime startDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            //DateTime endDate = startDate.AddMonths(1).AddDays(-1); 
            //list.Add("@FromDate", Convert.ToDateTime(startDate).ToString("dd-MMM-yyyy"));
            //list.Add("@ToDate", Convert.ToDateTime(endDate).ToString("dd-MMM-yyyy"));
            //list.Add("@loginId", Convert.ToInt32(Session["EmpId"].ToString()));
            //list.Add("@Role", Enum.GetName(typeof(payrollFunctions.roleIds), Convert.ToInt32(Session["RoleId"].ToString())));
            //DataTable dt = comfun.fillDataTable("Attendance_SalaryTags", "", list);
            return View();
        }
        // [customAuthorize(Roles = payrollFunctions.rolePayrollTeam)]






        public ActionResult _AttendanceApprovalByDepartment(FormCollection form)
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}

            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
            list.Add("@StartDate", Request.Form["FromDate"].ToString());
            list.Add("@EndDate", Request.Form["ToDate"].ToString());
            DataTable dt = comfun.fillDataTable("stpEmployeeAttendance_Monthly", "", list);

            //DataTable dt = comfun.fillDataTable("Attendance_SalaryTags", "", list);
            /*
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "Attendance", "_salaryAttMonthly", "_salaryAttMonthly");

            HtmlString htm = new HtmlString(paging);
            */
            ViewData["paging"] = "";

            int startMonth = DateTime.Now.Month;
            int startYear = DateTime.Now.Year;
            ViewData["FirstDay"] = Convert.ToDateTime(Request.Form["FromDate"]).Day;
            ViewData["LastDay"] = Convert.ToDateTime(Request.Form["ToDate"]).Day;

            ViewData["StartYear"] = Convert.ToDateTime(Request.Form["FromDate"]).Year;
            ViewData["StartMonth"] = Convert.ToDateTime(Request.Form["FromDate"]).Month;
            ViewData["EmpId"] = "1";// Request.Form["EmpId"].ToString();

            //DateTime startDate = new DateTime(startYear, startMonth, 1);
            //DateTime endDate = startDate.AddMonths(1).AddDays(-1);

            return PartialView(dt);
        }
        public JsonResult _AttendanceApprovalDepartmentSubmit(string FreezeEmpIds, string FromDate, string DepartmentId, string ContractorId)
        {
            string mes = string.Empty;
            try
            {
                DataTable table = (DataTable)JsonConvert.DeserializeObject(FreezeEmpIds, (typeof(DataTable)));


                SqlCommand com = new SqlCommand();
                connection conObj = new connection();

                com.Connection = conObj.con;

                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "Attendance_TimeSheetLockSubmit";
                SortedList list = new SortedList();
                DateTime monthYear = Convert.ToDateTime(FromDate);
                com.Parameters.AddWithValue("@MonthYearId", monthYear.ToString("MMyyyy"));

                com.Parameters.AddWithValue("@loginID", Session["EmpId"].ToString());
                com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());
                com.Parameters.AddWithValue("@DepartmentId", DepartmentId);
                com.Parameters.AddWithValue("@ContractorId", ContractorId);



                com.Parameters.AddWithValue("@Tag", "E");
                com.Parameters.AddWithValue("@RoleId", Session["RoleId"].ToString());
                com.Parameters.AddWithValue("@LockDate", FromDate);
                com.Parameters.AddWithValue("@CreatedDate", DateTime.Now.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));

                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@AttendanceTable";
                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table;
                com.Parameters.Add(parameter);

                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);


                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();

                com.ExecuteNonQuery();

                conObj.con.Close();

                mes = sp.Value.ToString();

                //return Json(
                //new
                //{
                //    Message = mes
                //},
                //JsonRequestBehavior.AllowGet
                //);
            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
                //return Json(
                //new
                //{
                //    Message = mes
                //},
                //JsonRequestBehavior.AllowGet
                //);

            }
            return Json(mes);
        }


        //[customAuthorize(Roles = payrollFunctions.rolePayrollTeam)]
        public ActionResult _Contractorddl()
        {
            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("stpContractorddl", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult AttendanceApprovedByCustomer()

        {
            comfun.saveformname("AttendanceApprovedByCustomer", "/TimeSheet/AttendanceApprovedByCustomer", "Attendance Approved By Customer", "Attendance Approved By Customer", "Y", "Attendance Approved By Customer", "AttendanceApprovedByCustomer", "List");


            if (payfun.sessionRecreate() == "expires")
            {
                return RedirectToAction("Login", "account");
            }
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            //SortedList list = new SortedList(); 
            ////***************Parameters*******************************************
            //DateTime startDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            //DateTime endDate = startDate.AddMonths(1).AddDays(-1); 
            //list.Add("@FromDate", Convert.ToDateTime(startDate).ToString("dd-MMM-yyyy"));
            //list.Add("@ToDate", Convert.ToDateTime(endDate).ToString("dd-MMM-yyyy"));
            //list.Add("@loginId", Convert.ToInt32(Session["EmpId"].ToString()));
            //list.Add("@Role", Enum.GetName(typeof(payrollFunctions.roleIds), Convert.ToInt32(Session["RoleId"].ToString())));
            //DataTable dt = comfun.fillDataTable("Attendance_SalaryTags", "", list);
            return View();
        }
        // [customAuthorize(Roles = payrollFunctions.rolePayrollTeam)]
        public ActionResult _AttendanceApprovedByCustomer(FormCollection form)
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}

            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@ContractorId", Request.Form["ContractorId"].ToString());
            list.Add("@StartDate", Request.Form["FromDate"].ToString());
            list.Add("@EndDate", Request.Form["ToDate"].ToString());
            DataTable dt = comfun.fillDataTable("stpEmployeeAttendanceApproval_Monthlybycontractor1", "", list);

            //DataTable dt = comfun.fillDataTable("Attendance_SalaryTags", "", list);
            /*
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "Attendance", "_salaryAttMonthly", "_salaryAttMonthly");

            HtmlString htm = new HtmlString(paging);
            */
            ViewData["paging"] = "";

            int startMonth = DateTime.Now.Month;
            int startYear = DateTime.Now.Year;
            ViewData["FirstDay"] = Convert.ToDateTime(Request.Form["FromDate"]).Day;
            ViewData["LastDay"] = Convert.ToDateTime(Request.Form["ToDate"]).Day;

            ViewData["StartYear"] = Convert.ToDateTime(Request.Form["FromDate"]).Year;
            ViewData["StartMonth"] = Convert.ToDateTime(Request.Form["FromDate"]).Month;
            ViewData["EmpId"] = "1";// Request.Form["EmpId"].ToString();

            //DateTime startDate = new DateTime(startYear, startMonth, 1);
            //DateTime endDate = startDate.AddMonths(1).AddDays(-1);

            return PartialView("_TimeSheetAttendance", dt);
        }
        public ActionResult _AttendanceApprovedCustomerwiseList(FormCollection form)
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}

            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@ContractorId", Request.Form["ContractorId"].ToString());
            list.Add("@CustomerId", Session["Customer"]);

            list.Add("@StartDate", Request.Form["FromDate"].ToString());
            list.Add("@EndDate", Request.Form["ToDate"].ToString());
            DataTable dt = comfun.fillDataTable("stpEmployeeAttendanceApproval_Monthlybycontractor1", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        //public ActionResult _AttendanceUnApprovedCustomerwiseList(FormCollection form)
        //{
        //    //if (payfun.sessionRecreate() == "expires")
        //    //{
        //    //    return Json("Session expires");
        //    //}

        //    SortedList list = new SortedList();
        //    list.Add("@EmpId", Request.Form["EmpId"].ToString());
        //    list.Add("@ContractorId", Request.Form["ContractorId"].ToString());
        //    list.Add("@StartDate", Request.Form["FromDate"].ToString());
        //    list.Add("@EndDate", Request.Form["ToDate"].ToString());
        //    DataTable dt = comfun.fillDataTable("stpEmployeeAttendanceApproval_Monthlybycontractor1", "", list);
        //    var json = JsonConvert.SerializeObject(dt);
        //    return Json(json);
        //}
        public ActionResult _ContrcatorUnApprovedList(FormCollection form)
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}

            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@ContractorId", Request.Form["ContractorId"].ToString());
            list.Add("@StartDate", Request.Form["FromDate"].ToString());
            list.Add("@EndDate", Request.Form["ToDate"].ToString());
            DataTable dt = comfun.fillDataTable("stpEmployeeunApproval_Monthlybycontractor", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _ContrcatorApprovedList(FormCollection form)
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}

            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@ContractorId", Request.Form["ContractorId"].ToString());
            list.Add("@StartDate", Request.Form["FromDate"].ToString());
            list.Add("@EndDate", Request.Form["ToDate"].ToString());
            DataTable dt = comfun.fillDataTable("stpEmployeeApproval_Monthlybycontractorlist", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _ContrcatorApprovedList1(FormCollection form)
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}

            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@ContractorId", Request.Form["ContractorId"].ToString());
            list.Add("@StartDate", Request.Form["FromDate"].ToString());
            list.Add("@EndDate", Request.Form["ToDate"].ToString());
            DataTable dt = comfun.fillDataTable("stpEmployeeApproval_Monthlybycontractor", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }



        public JsonResult _AttendanceApprovalBillContractorSubmit(string BillDetail, string FromDate, string Id)
        {
            string mes = string.Empty;
            try
            {
                DataTable table = (DataTable)JsonConvert.DeserializeObject(BillDetail, (typeof(DataTable)));


                SqlCommand com = new SqlCommand();
                connection conObj = new connection();

                com.Connection = conObj.con;

                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "Contract_TimeSheetLockSubmit";
                SortedList list = new SortedList();
                DateTime monthYear = Convert.ToDateTime(FromDate);
                com.Parameters.AddWithValue("@MonthYearId", monthYear.ToString("MMyyyy"));
                com.Parameters.AddWithValue("@Id", Id);

                com.Parameters.AddWithValue("@loginID", Session["EmpId"].ToString());
                com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());


                com.Parameters.AddWithValue("@Tag", "E");
                com.Parameters.AddWithValue("@RoleId", Session["RoleId"].ToString());
                com.Parameters.AddWithValue("@LockDate", FromDate);
                com.Parameters.AddWithValue("@CreatedDate", DateTime.Now.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));

                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@AttendanceTable";
                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table;
                com.Parameters.Add(parameter);

                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);


                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();

                com.ExecuteNonQuery();

                conObj.con.Close();

                mes = sp.Value.ToString();

                //return Json(
                //new
                //{
                //    Message = mes
                //},
                //JsonRequestBehavior.AllowGet
                //);
            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
                //return Json(
                //new
                //{
                //    Message = mes
                //},
                //JsonRequestBehavior.AllowGet
                //);

            }
            return Json(mes);
        }

        public JsonResult _AttendanceApprovalContractorSubmit(string BillDetail, string FromDate, string BillNo, string BillDate,
            string Id, string Image)
        {

            string mes = string.Empty;
            var provider = new MultipartMemoryStreamProvider();
            Random ran = new Random();




            try
            {
                DataTable table = (DataTable)JsonConvert.DeserializeObject(BillDetail, (typeof(DataTable)));



                //string path = AppDomain.CurrentDomain.BaseDirectory + "Uploads/";  
                //string filename = Path.GetFileName(Request.Files[i].FileName);  



                // Checking for Internet Explorer  




                SqlCommand com = new SqlCommand();
                connection conObj = new connection();

                com.Connection = conObj.con;

                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "ContractorAttendance_TimeSheetLockSubmit";
                SortedList list = new SortedList();
                DateTime monthYear = Convert.ToDateTime(FromDate);
                com.Parameters.AddWithValue("@MonthYearId", monthYear.ToString("MMyyyy"));
                com.Parameters.AddWithValue("@Id", Id);
                com.Parameters.AddWithValue("@loginID", Session["EmpId"].ToString());
                com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());
                com.Parameters.AddWithValue("@BillDate", BillDate);
                com.Parameters.AddWithValue("@BillNo", BillNo);
                com.Parameters.AddWithValue("@Tag", "E");
                com.Parameters.AddWithValue("@RoleId", Session["RoleId"].ToString());
                com.Parameters.AddWithValue("@LockDate", FromDate);
                com.Parameters.AddWithValue("@CreatedDate", DateTime.Now.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));

                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@AttendanceTable";
                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table;
                com.Parameters.Add(parameter);

                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);


                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();

                com.ExecuteNonQuery();

                conObj.con.Close();

                mes = sp.Value.ToString();

                //return Json(
                //new
                //{
                //    Message = mes
                //},
                //JsonRequestBehavior.AllowGet
                //);
            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
                //return Json(
                //new
                //{
                //    Message = mes
                //},
                //JsonRequestBehavior.AllowGet
                //);

            }
            return Json(mes);
        }
        public ActionResult DepartmentWiseBill()
        {
            return View();
        }

        public ActionResult DepartmentBillGenerate()
        {
            comfun.saveformname("DepartmentBillGenerate", "/TimeSheet/DepartmentBillGenerate", "TimeSheet/DepartmentBillGenerate", "Department Bill Generate form", "Y", "DepartmentBillGenerate", "DepartmentBillGenerate", "View", 1);
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            return View();
        }
        public JsonResult _DepartmentBillGenerate()
        {
            SortedList list = new SortedList();
            string mes = string.Empty;
            try
            {
                if (Request.Form["FromDate"] != null)
                {
                    list.Add("@StartDate", Request.Form["FromDate"].ToString());
                }
                if (Request.Form["ToDate"] != null)
                {
                    list.Add("@EndDate", Convert.ToDateTime(Request.Form["ToDate"]).ToString("dd-MMM-yyyy"));
                }
                //list.Add("@loginId", Convert.ToInt32(Session["EmpId"].ToString()));
                //
                //list.Add("@Role", Enum.GetName(typeof(payrollFunctions.roleIds), Convert.ToInt32(Session["RoleId"].ToString())));
                if (Request.Form["EmpId"] != "")
                {
                    list.Add("@EmpId", Request.Form["EmpId"].ToString());
                }
                if (Request.Form["BranchId"] != "")
                {
                    list.Add("@BranchId", Request.Form["BranchId"].ToString());
                }
                if (Request.Form["DesignationId"] != "")
                {
                    list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
                }
                if (Request.Form["DepartmentId"] != "")
                {
                    list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
                }
                if (payfun.getQueryStringValueOrDefault("CategoryId") != "")
                {
                    list.Add("@CategoryId", Request.QueryString["CategoryId"].ToString());
                }
                //if (payfun.getQueryStringValueOrDefault("CircleId") != "")
                //{
                //    list.Add("@CircleId", Request.QueryString["CircleId"].ToString());
                //}
                mes = comfun.executeNonQueryWMessage("stpDepartmentSalaryCalculation", "", list).ToString();

            }
            catch (Exception ex)
            {

                mes = "Error:" + ex.Message;
            }

            return Json(mes);
        }

        public ActionResult _DepartmentCertificateGenerate()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            if (Request.Form["FromDate"] != null)
            {
                list.Add("@StartDate", Request.Form["FromDate"].ToString());
            }
            if (Request.Form["DepartmentId"].ToString() != "0")
            {
                list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
            }
            else
            {
                list.Add("@DepartmentId", Session["Customer"].ToString());
            }
            list.Add("@BranchId", Request.Form["BranchId"].ToString());
            dt = comfun.fillDataTable("stpAttendanceApproveDepartmentWise", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _DepartmentCertificateDetail(string Certificate, string Id, string FromDate)
        {
            string mes = string.Empty;
            try
            {
                DataTable table = (DataTable)JsonConvert.DeserializeObject(Certificate, (typeof(DataTable)));


                SqlCommand com = new SqlCommand();
                connection conObj = new connection();

                com.Connection = conObj.con;

                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "stpAttendanceCertficateGenerate_Accept";
                SortedList list = new SortedList();
                DateTime monthYear = Convert.ToDateTime(FromDate);
                com.Parameters.AddWithValue("@Id", Id);
                com.Parameters.AddWithValue("@MonthYear", monthYear.ToString("MMyyyy"));
                com.Parameters.AddWithValue("@lockby", Session["EmpId"].ToString());
                com.Parameters.AddWithValue("@LockDate", DateTime.Now.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                com.Parameters.AddWithValue("@LockMonth", FromDate);

                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@AttendanceCertificate";
                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table;
                com.Parameters.Add(parameter);

                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);


                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();

                com.ExecuteNonQuery();

                conObj.con.Close();

                mes = sp.Value.ToString();

                //return Json(
                //new
                //{
                //    Message = mes
                //},
                //JsonRequestBehavior.AllowGet
                //);
            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
                //return Json(
                //new
                //{
                //    Message = mes
                //},
                //JsonRequestBehavior.AllowGet
                //);

            }
            return Json(mes);
        }

        public ActionResult _DepartmentCertificateList()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            if (Request.Form["FromDate"] != null)
            {
                list.Add("@FromDate", Request.Form["FromDate"].ToString());
            }
            // list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
            if (Request.Form["DepartmentId"].ToString() != "0")
            {
                list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
            }
            else
            {
                list.Add("@DepartmentId", Session["Customer"].ToString());
            }
            list.Add("@BranchId", Request.Form["BranchId"].ToString());
            dt = comfun.fillDataTable("stpEmployeeCertificateList", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult CertificatePrint()
        {
            ViewBag.Title = "Certificate";
            string mes = string.Empty;


            HtmlString htmlString = new HtmlString(mes);

            StreamReader reader = new StreamReader(Server.MapPath("~/PrintTemplates/Certificate.html"));
            mes = reader.ReadToEnd();

            string table = string.Empty;
            SortedList list = new SortedList();
            list.Add("@CertificateId", Request.QueryString["id"].ToString());

            DataSet ds = comfun.fillDataSet("stpDepatmentCertificatePrint", "", list);
            mes = mes.Replace("#Department#", ds.Tables[0].Rows[0]["DepatmentName"].ToString());
            mes = mes.Replace("#Certificate#", ds.Tables[0].Rows[0]["CertificateId"].ToString());
            mes = mes.Replace("#MonthYear#", ds.Tables[0].Rows[0]["LockMonth"].ToString());
            mes = mes.Replace("#Date#", ds.Tables[0].Rows[0]["GeneratedDate"].ToString());
            //mes = mes.Replace("#Board#", ds.Tables[0].Rows[0]["Board"].ToString());



            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                table += "<tr>";
                //table += "<td>" + (i + 1).ToString() + "</td>";
                table += "<td align=center>" + ds.Tables[0].Rows[i]["EmpCode"].ToString() + "</td>";
                table += "<td  align=center>" + ds.Tables[0].Rows[i]["EmpName"].ToString() + "</td>";
                table += "<td  align=center>" + ds.Tables[0].Rows[i]["PayDays"].ToString() + "</td>";

            }

            mes = mes.Replace("#tbody", table);

            mes = mes.Replace("#tbody", table);
            htmlString = new HtmlString(mes);
            ViewData["Template"] = htmlString;
            reader.Close();
            reader.Dispose();
            return View();
        }

        public JsonResult UploadFiles()
        {
            string mes = string.Empty;
            if (Request.Files.Count > 0)
            {
                var fname = "";
                var ImageName = "";
                try
                {
                    //  Get all files from Request object  
                    HttpFileCollectionBase files = Request.Files;

                    if (System.Web.HttpContext.Current.Request.Files.AllKeys.Any())
                    {
                        Random ran = new Random();
                        string randomno = ran.Next(111, 999).ToString();
                        for (int i = 0; i < files.Count; i++)
                        {
                            //string path = AppDomain.CurrentDomain.BaseDirectory + "Uploads/";  
                            //string filename = Path.GetFileName(Request.Files[i].FileName);  

                            HttpPostedFileBase file = files[i];

                            // Checking for Internet Explorer  
                            if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                            {
                                string[] testfiles = file.FileName.Split(new char[] { '\\' });
                                fname = testfiles[testfiles.Length - 1];

                            }
                            else
                            {
                                fname = file.FileName;
                            }
                            var _ext = Path.GetExtension(fname);
                            fname = Path.GetFileNameWithoutExtension(fname);

                            // Get the complete folder path and store the file inside it.  

                            ImageName = "Certificate" + fname + Request.Form["Category"].ToString() + randomno + _ext;
                            string filePath = Path.Combine(Server.MapPath("~/Certificate/DepartmentCertificate/") + ImageName);

                            fname = filePath;
                            var filePathA = filePath;
                            //  file.SaveAs(filePathA);
                            file.SaveAs(filePathA);
                            SortedList list = new SortedList();
                            list.Add("@CertificateId", Request.Form["Category"].ToString());

                            list.Add("@fvFileName", ImageName);
                            list.Add("@fvFilePath", filePathA);



                            mes = comfun.executeNonQueryWMessage("stpDepartmentfileUpload", "", list).ToString();
                        }
                    }

                    // Returns message that successfully uploaded  

                }
                catch (Exception ex)
                {
                    return Json("Error occurred. Error details: " + ex.Message);
                }
            }
            else
            {
                return Json("No files selected.");
            }


            /*-------------------------------------------------------------*/

            return Json(mes);
        }
        public JsonResult UploadCustomerFiles()
        {
            string mes = string.Empty;
            if (Request.Files.Count > 0)
            {
                var fname = "";
                var ImageName = "";
                try
                {
                    //  Get all files from Request object  
                    HttpFileCollectionBase files = Request.Files;

                    if (System.Web.HttpContext.Current.Request.Files.AllKeys.Any())
                    {
                        Random ran = new Random();
                        string randomno = ran.Next(111, 999).ToString();
                        for (int i = 0; i < files.Count; i++)
                        {
                            //string path = AppDomain.CurrentDomain.BaseDirectory + "Uploads/";  
                            //string filename = Path.GetFileName(Request.Files[i].FileName);  

                            HttpPostedFileBase file = files[i];

                            // Checking for Internet Explorer  
                            if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                            {
                                string[] testfiles = file.FileName.Split(new char[] { '\\' });
                                fname = testfiles[testfiles.Length - 1];

                            }
                            else
                            {
                                fname = file.FileName;
                            }
                            var _ext = Path.GetExtension(fname);
                            fname = Path.GetFileNameWithoutExtension(fname);

                            // Get the complete folder path and store the file inside it.  

                            ImageName = "Certificate" + fname + Request.Form["Category"].ToString() + randomno + _ext;
                            string filePath = Path.Combine(Server.MapPath("~/Certificate/CustomerBill/") + ImageName);

                            fname = filePath;
                            var filePathA = filePath;
                            //  file.SaveAs(filePathA);
                            file.SaveAs(filePathA);
                            SortedList list = new SortedList();
                            list.Add("@CertificateId", Request.Form["Category"].ToString());

                            list.Add("@fvFileName", ImageName);
                            list.Add("@fvFilePath", filePathA);



                            mes = comfun.executeNonQueryWMessage("stpCustomerfileUpload", "", list).ToString();
                        }
                    }

                    // Returns message that successfully uploaded  

                }
                catch (Exception ex)
                {
                    return Json("Error occurred. Error details: " + ex.Message);
                }
            }
            else
            {
                return Json("No files selected.");
            }


            /*-------------------------------------------------------------*/

            return Json(mes);
        }

        public ActionResult ContractorBillGenerate()
        {
            comfun.saveformname("ContractorBillGenerate", "/TimeSheet/ContractorBillGenerate", "TimeSheet/ContractorBillGenerate", "Contractor Bill Generate form", "Y", "ContractorBillGenerate", "ContractorBillGenerate", "View", 1);
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            return View();
        }
        public JsonResult _ContractorBillGenerate(string BillDetail, string CustomerId, string Id, string FromDate)
        {

            string mes = string.Empty;


            try
            {
                DataTable table = (DataTable)JsonConvert.DeserializeObject(BillDetail, (typeof(DataTable)));


                SqlCommand com = new SqlCommand();
                connection conObj = new connection();

                com.Connection = conObj.con;



                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "stpContractorBillSalaryCalculation";
                SortedList list = new SortedList();
                DateTime monthYear = Convert.ToDateTime(FromDate);
                com.Parameters.AddWithValue("@MonthDate", FromDate);
                com.Parameters.AddWithValue("@ID", Id);
                com.Parameters.AddWithValue("@CustomerId", CustomerId);


                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@CustomerInvoiceP2";
                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table;
                com.Parameters.Add(parameter);

                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);


                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();

                com.ExecuteNonQuery();

                conObj.con.Close();

                mes = sp.Value.ToString();

                //return Json(
                //new
                //{
                //    Message = mes
                //},
                //JsonRequestBehavior.AllowGet
                //);
            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
                //return Json(
                //new
                //{
                //    Message = mes
                //},
                //JsonRequestBehavior.AllowGet
                //);

            }
            return Json(mes);

        }
        public ActionResult DepartmentApproval()
        {
            return View();

        }
        public ActionResult _EmployeeAttendanceList()
        {


            SortedList list = new SortedList();

            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@FromDate", Request.Form["FromDate"].ToString());
            list.Add("@Category", Request.Form["CategoryId"].ToString());
            list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
            list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
            list.Add("@BranchId", Request.Form["BranchId"].ToString());
            list.Add("@CustomerId", Session["Customer"].ToString());
            // list.Add("@CircleId", Request.Form["CircleId"].ToString());

            DataTable dt = comfun.fillDataTable("stpEmployeeAttendanceApprovedList", "", list);

            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult _EmployeeAttendanceList1()
        {
            decimal total_records = 0;
            decimal pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            decimal pagesize = Convert.ToDecimal(Request.Form["PageSize"]);

            SortedList list = new SortedList();
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@FromDate", Request.Form["FromDate"].ToString());
            list.Add("@Category", Request.Form["CategoryId"].ToString());
            list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
            list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
            list.Add("@BranchId", Request.Form["BranchId"].ToString());
            list.Add("@CircleId", Request.Form["CircleId"].ToString());
            list.Add("@LoginId", Session["EmpId"].ToString());
            list.Add("@CustomerId", Session["Customer"].ToString());
            DataTable dt = comfun.fillDataTable("stpEmployeeAttendanceMannualImportList", "", list);
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            //string paging = comfun.create_paging(total_records, pagesize, pageno, "home", "Index");
            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "Attendance", "_EmployeeAttendanceList1", "_EmployeeAttendanceList1");
            HtmlString htm = new HtmlString(paging);
            ViewData["paging"] = htm;
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(new
            {
                Paging = paging,
                data = json1
            });
        }
        public JsonResult _AttendanceImportSubmit1(string data, string AttendanceDate, string Otp)
        {
            string mes = string.Empty;
            try
            {
                DataTable table = (DataTable)JsonConvert.DeserializeObject(data, (typeof(DataTable)));


                SqlCommand com = new SqlCommand();
                connection conObj = new connection();

                com.Connection = conObj.con;

                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "stpEmployeeAttendanceDepartmentApproval_Accept";
                SortedList list = new SortedList();
                com.Parameters.AddWithValue("@AttendanceDate", AttendanceDate);
                com.Parameters.AddWithValue("@Otp", Otp);
                com.Parameters.AddWithValue("@loginEmail", Session["Email"]);
                com.Parameters.AddWithValue("@Month", Convert.ToDateTime(AttendanceDate).ToString("MMyyyy"));
                com.Parameters.AddWithValue("@CreateBy", Session["EmpId"].ToString());
                com.Parameters.AddWithValue("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));


                // com.Parameters.AddWithValue("@fdWeeklyoffDate", attendanceDate);
                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@AttendanceBulkEmp";
                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table;
                com.Parameters.Add(parameter);
                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);
                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();
                com.ExecuteNonQuery();

                conObj.con.Close();

                mes = sp.Value.ToString();


            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;

            }
            return Json(mes);

        }


        public ActionResult DisplayUploadedExcel()
        {
            SortedList list = new SortedList();

            if (Session["ExcelId"] != null)
            {
                list.Add("ExcelId", Session["ExcelId"].ToString());
            }
            //DataTable dt = (DataTable)Session["Excel"];
            DataTable dt = comfun.fillDataTable("UploadedExcel_Display", null, list);
            return View(dt);
        }
        public ActionResult CustomerBillGenerate()
        {
            return View();

        }
        public ActionResult InvoicetoCustomer()
        {
            return View();
        }
        public ActionResult _ContrcatorInvoiceList()
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}

            SortedList list = new SortedList();

            list.Add("@ContractorId", "0");
            list.Add("@CustomerId", Request.Form["CustomerId"].ToString());
            list.Add("@EntityId", Request.Form["EntityId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthYear"].ToString());

            DataTable dt = comfun.fillDataTable("stpAttendanceApproval_InvoiceList", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _ContrcatorApprovedInvoiceList()
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}

            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@ContractorId", Request.Form["ContractorId"].ToString());
            list.Add("@CustomerId", Request.Form["CustomerId"].ToString());
            list.Add("@StartDate", Request.Form["FromDate"].ToString());
            list.Add("@EndDate", Request.Form["ToDate"].ToString());
            DataTable dt = comfun.fillDataTable("stpContrcatorApproved_InvoiceList", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult BillAmount()
        {
            SortedList list = new SortedList();

            list.Add("@CustomerInvoiceP2", Request.Form["BillDetail"].ToString());
            list.Add("@CustomerId", Request.Form["CustomerId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthYear"].ToString());
            DataSet dt = comfun.fillDataSet("stpBillGenerate", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _AttendanceCustomerInvoiceSubmit(string BillDetail, string BillDesignation, string FromDate, string BillNo,
            string BillDate, string Id, string Remarks, string LetterNo, string Letterdate, string ReferanceNo, string CustomerID, string Amount, string GSTAmount, string SGSTAmount, string CGSTAmount
            , string HPSEDC, string BillAmount, string RoundOff
           )
        {

            string mes = string.Empty;
            try
            {
                DataTable table = (DataTable)JsonConvert.DeserializeObject(BillDetail, (typeof(DataTable)));

                DataTable table1 = (DataTable)JsonConvert.DeserializeObject(BillDesignation, (typeof(DataTable)));
                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "stpCustomerInvoice_Accept";
                SortedList list = new SortedList();
                DateTime monthYear = Convert.ToDateTime(FromDate);
                com.Parameters.AddWithValue("@MonthDate", FromDate);
                com.Parameters.AddWithValue("@ID", Id);
                com.Parameters.AddWithValue("@Createby", Session["EmpId"].ToString());
                com.Parameters.AddWithValue("@BillDate", BillDate);
                com.Parameters.AddWithValue("@fvBillNo", BillNo);
                com.Parameters.AddWithValue("@RefNo", ReferanceNo);
                com.Parameters.AddWithValue("@LetterNo", LetterNo);
                com.Parameters.AddWithValue("@Letterdate", Letterdate);
                com.Parameters.AddWithValue("@Remarks", Remarks);
                com.Parameters.AddWithValue("@CustomerID", CustomerID);
                //com.Parameters.AddWithValue("@VendorID", VendorID);
                com.Parameters.AddWithValue("@SessionID", Session["SessionId"]);

                com.Parameters.AddWithValue("@Amount", Amount);
                com.Parameters.AddWithValue("@GSTAmount", GSTAmount);
                com.Parameters.AddWithValue("@SGSTAmount", SGSTAmount);
                com.Parameters.AddWithValue("@CGSTAmount", CGSTAmount);

                com.Parameters.AddWithValue("@OtherCharges", HPSEDC);
                com.Parameters.AddWithValue("@BillAmount", BillAmount);
                com.Parameters.AddWithValue("@RoundOff", RoundOff);
                com.Parameters.AddWithValue("@createdDate", DateTime.Now.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));

                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@CustomerInvoiceP2";
                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table;
                com.Parameters.Add(parameter);

                SqlParameter parameter1 = new SqlParameter();
                parameter1.ParameterName = "@CustomerInvoiceP3";
                parameter1.SqlDbType = System.Data.SqlDbType.Structured;
                parameter1.Value = table1;
                com.Parameters.Add(parameter1);
                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);


                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();

                com.ExecuteNonQuery();

                conObj.con.Close();

                mes = sp.Value.ToString();

                //return Json(
                //new
                //{
                //    Message = mes
                //},
                //JsonRequestBehavior.AllowGet
                //);
            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
                //return Json(
                //new
                //{
                //    Message = mes
                //},
                //JsonRequestBehavior.AllowGet
                //);

            }
            return Json(mes);
        }
        public ActionResult _CustomerEntityddl()
        {
            SortedList list = new SortedList();
            list.Add("@CustomerId", Session["Customer"].ToString());
            DataTable dt = comfun.fillDataTable("stpCustomerEntityddl", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult CustomerBill()
        {
            comfun.saveformname("CustomerBill", "/TimeSheet/CustomerBill", "Attendance Mgt/AttendanceApproval/CustomerBill", "Customer Invoice  Main form", "Y", "CustomerBill", "CustomerBill", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("CustomerBill");
            return View();
        }
        public ActionResult _CustomerBillDetailList()
        {
            comfun.saveformname("_CustomerBillDetailList", "/TimeSheet/_CustomerBillDetailList", "Master/Address/CustomerBill", "Customer Invoice  list", "N", "CustomerBill", "CustomerBill", "List", 4);

            SortedList list = new SortedList();
            list.Add("@CustomerId", Request.Form["Customer"].ToString());
            list.Add("@SupplierId", Request.Form["SupplierId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthYear"].ToString());
            DataTable dt = comfun.fillDataTable("stpBillDetailList", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public ActionResult CustomerPayment()
        {
            comfun.saveformname("CustomerPayment", "/TimeSheet/CustomerPayment", "Attendance Mgt/AttendanceApproval/CustomerPayment", "Customer Payment  Main form", "Y", "CustomerPayment", "CustomerPayment", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("CustomerPayment");
            return View();
        }

        public ActionResult _CustomerPaymentddl()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@CustomerId", Session["Customer"]);
            dt = comfun.fillDataTable("stpCustomerPaymentddl", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }

        public ActionResult _Billddl()
        {
            // comfun.saveformname("CustomerPayment", "/TimeSheet/CustomerPayment", "Attendance Mgt/AttendanceApproval/CustomerPayment", "Customer Payment  Main form", "Y", "CustomerPayment", "CustomerPayment", "View", 1);
            //  ViewData["AccessRights"] = payfun.getAccessRights("CustomerPayment");
            DateTime monthYear = Convert.ToDateTime(Request.Form["BillDate"]);
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@CustomerID", Request.Form["CustomerId"].ToString());
            list.Add("@MonthYear", monthYear.ToString("MMyyyy"));
            dt = comfun.fillDataTable("stpCustomerBill_Pending", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public ActionResult _BillAmountDetail()
        {
            // comfun.saveformname("CustomerPayment", "/TimeSheet/CustomerPayment", "Attendance Mgt/AttendanceApproval/CustomerPayment", "Customer Payment  Main form", "Y", "CustomerPayment", "CustomerPayment", "View", 1);
            //  ViewData["AccessRights"] = payfun.getAccessRights("CustomerPayment");
            DateTime monthYear = Convert.ToDateTime(Request.Form["BillDate"]);
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@CustomerID", Request.Form["CustomerId"].ToString());
            list.Add("@BillId", Request.Form["BillId"].ToString());
            list.Add("@MonthYear", monthYear.ToString("MMyyyy"));
            dt = comfun.fillDataTable("stpCustomerBillAmount", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }

        public ActionResult _CustomerPaymentDetailList()
        {
            comfun.saveformname("_CustomerPaymentDetailList", "/TimeSheet/_CustomerPaymentDetailList", "Attendance Mgt/AttendanceApproval/CustomerPayment", "Customer Payment list", "N", "CustomerPayment", "CustomerPayment", "List", 4);

            SortedList list = new SortedList();
            list.Add("@CustomerId", Request.Form["Customer"].ToString());
            list.Add("@SupplierId", Request.Form["SupplierId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthYear"].ToString());
            DataTable dt = comfun.fillDataTable("stpCustomerPaymentList", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public JsonResult _CustomerPaymentSubmit()
        {
            comfun.saveformname("_CustomerPaymentSubmit", "/TimeSheet/_CustomerPaymentSubmit", "Attendance Mgt/AttendanceApproval/CustomerPayment", "Customer Payment  add", "N", "CustomerPayment", "CustomerPayment", "Add", 2);
            comfun.saveformname("_CustomerPaymentSubmit", "/TimeSheet/_CustomerPaymentSubmit", "Attendance Mgt/AttendanceApproval/CustomerPayment", "Customer Payment  edit", "N", "CustomerPayment", "CustomerPayment", "Edit", 3);
            DateTime monthYear = Convert.ToDateTime(Request.Form["BillDate"]);
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@ID", Request.Form["Id"].ToString());
                list.Add("@CustomerId", Session["Customer"].ToString());
                list.Add("@BillDate", Request.Form["BillDate"].ToString());
                list.Add("@MonthYear", monthYear.ToString("MMyyyy"));
                list.Add("@SessionID", Session["SessionId"].ToString());
                list.Add("@MonthDate", monthYear.ToString("dd-MMM-yyyy"));
                list.Add("@BillId", Request.Form["BillId"].ToString());
                list.Add("@Createby", Session["EmpId"].ToString());
                list.Add("@createdDate", DateTime.Now.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                list.Add("@PaidAmount", Request.Form["PaidAmount"].ToString());
                list.Add("@BillAmount", Request.Form["BillAmount"].ToString());
                list.Add("@DeductionAmount", Request.Form["DeductionAmount"].ToString());
                list.Add("@RefNo", Request.Form["RefNo"].ToString());
                list.Add("@AccountNo", Request.Form["AccountNo"].ToString());
                mes = comfun.executeNonQueryWMessage("stpCustomerPayment_Accept", "", list).ToString();

            }
            catch (Exception ex)
            {
                list.Clear();
                list.Add("Error", mes);
            }
            return Json(mes);
        }
        public JsonResult _CustomerPaymentStatusUpdate()
        {
            comfun.saveformname("_CustomerPaymentStatusUpdate", "/TimeSheet/_CustomerPaymentStatusUpdate", "Attendance Mgt/AttendanceApproval/CustomerPayment", "Customer Payment  Status", "N", "CustomerPayment", "CustomerPayment", "Status", 5);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpCustomerPaymentActivationStatusUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }
        public ActionResult _CustomerPaymentEdit()
        {
            //comfun.saveformname("AllowanceEdit", "/admin/AllowanceEdit", "Allowance Edit", "Allowance Edit", "N", "AllowanceEdit", "Allowance", "Edit");

            string Id = "0";
            if (Request.Form["Id"].ToString() == null)
            {
                return RedirectToAction("Allowances");
            }
            Id = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@Id", Id);
            DataTable dt = comfun.fillDataTable("stpCustomerPaymentEdit", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);

        }

        public JsonResult UploadPaymentFiles()
        {
            string mes = string.Empty;
            if (Request.Files.Count > 0)
            {
                var fname = "";
                var ImageName = "";
                try
                {
                    //  Get all files from Request object  
                    HttpFileCollectionBase files = Request.Files;

                    if (System.Web.HttpContext.Current.Request.Files.AllKeys.Any())
                    {
                        Random ran = new Random();
                        string randomno = ran.Next(111, 999).ToString();
                        for (int i = 0; i < files.Count; i++)
                        {
                            //string path = AppDomain.CurrentDomain.BaseDirectory + "Uploads/";  
                            //string filename = Path.GetFileName(Request.Files[i].FileName);  

                            HttpPostedFileBase file = files[i];

                            // Checking for Internet Explorer  
                            if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                            {
                                string[] testfiles = file.FileName.Split(new char[] { '\\' });
                                fname = testfiles[testfiles.Length - 1];

                            }
                            else
                            {
                                fname = file.FileName;
                            }
                            var _ext = Path.GetExtension(fname);
                            fname = Path.GetFileNameWithoutExtension(fname);

                            // Get the complete folder path and store the file inside it.  

                            ImageName = "Payment" + fname + Request.Form["Category"].ToString() + randomno + _ext;
                            string filePath = Path.Combine(Server.MapPath("~/Payment/Customer/") + ImageName);

                            fname = filePath;
                            var filePathA = filePath;
                            //  file.SaveAs(filePathA);
                            file.SaveAs(filePathA);
                            SortedList list = new SortedList();
                            list.Add("@Id", Request.Form["Category"].ToString());

                            list.Add("@fvFileName", ImageName);
                            list.Add("@fvFilePath", filePathA);



                            mes = comfun.executeNonQueryWMessage("stpPaymentfileUpload", "", list).ToString();
                        }
                    }

                    // Returns message that successfully uploaded  

                }
                catch (Exception ex)
                {
                    return Json("Error occurred. Error details: " + ex.Message);
                }
            }
            else
            {
                return Json("No files selected.");
            }


            /*-------------------------------------------------------------*/

            return Json(mes);
        }
        public ActionResult InvoicePrint()
        {

            return Redirect("~/CrystalReports/PIprint.aspx?Id=" + Request.QueryString["Id"].ToString() + "&&" + "Method=" + "Invoice");
        }

        public ActionResult CustomerBillPrint()
        {

            return Redirect("~/CrystalReports/PIprint.aspx?Id=" + Request.QueryString["Id"].ToString() + "&&" + "Method=" + "CustomerPrint");
        }
        public ActionResult CoveringLetterPrint()
        {

            return Redirect("~/CrystalReports/PIprint.aspx?Id=" + Request.QueryString["Id"].ToString() + "&&" + "Method=" + "CoveringLetter");
        }
        public ActionResult PrintQrcodeStudent(string Id)
        {

            //DataTable table = (DataTable)JsonConvert.DeserializeObject(Id, (typeof(DataTable)));
            return Redirect("~/CrystalReports/PIprint.aspx?Id=" + Id + "&&" + "Method=" + "StudentInfo");


            //for (int i = 0; i <= table.Rows.Count-1;i++)
            //{
            //    string ids = table.Rows[i]["Id"].ToString();

            //     Response.Redirect("~/CrystalReports/PIprint.aspx?Id=" + ids + "&&" + "Method=" + "StudentInfo");

            //}


            // return RedirectToAction("~/TimeSheet/LocationwiseMail");
        }
        public ActionResult CustomerBillRemove()
        {
            SortedList list = new SortedList();
            string mes = string.Empty;
            list.Add("@Id", Request.Form["Id"].ToString());
            mes = comfun.executeNonQueryWMessage("StpCustomerInvoice_Remove", "", list).ToString();
            return Json(mes);
        }
        public ActionResult CustomerbillDetail()
        {
            return View();
        }
        public ActionResult _BillDetailexcel()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@billId", Request.Form["BillId"].ToString());
            dt = comfun.fillDataTable("StpBillDetailExcel", "", list);

            return PartialView("_BillDetailexcel", dt);
        }
        public ActionResult _BillNo()
        {
            SortedList list = new SortedList();
            DataTable dt = new DataTable();
            list.Add("@MonthYearId", Request.Form["MonthYear"].ToString());
            dt = comfun.fillDataTable("stpBillddl", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        private string sendWelcomeEmail(string email, string name, string password)
        {


            //string mes = "";
            //string subject = "Otp";
            //mes = payfun.SendEmailOTP("WFMS", email, subject, name+" "+password);
            //if (!mes.ToLower().Contains("error"))
            //{
            //    mes = "Your account has been created successfully, Please check your inbox to verify your Email Address";
            //}
            string mes = "";
            using (MailMessage mm = new MailMessage(email, email))
            {
                mm.Subject = "Otp";
                mm.Body = name + " " + password;
                mm.CC.Add(new MailAddress("munish.kumar@horizontelecom.in"));
                //    mm.Attachments.Add(new Attachment(CustomerINV.ExportToStream(ExportFormatType.PortableDocFormat), "Crystal.pdf"));
                mm.IsBodyHtml = true;
                using (SmtpClient smtp = new SmtpClient())
                {
                    smtp.Host = "smtp.rediffmailpro.com";
                    smtp.UseDefaultCredentials = true;
                    smtp.Credentials = new NetworkCredential
                    {
                        UserName = "Wfms@horizontelecom.in",
                        Password = "wfms@9002"
                    };
                    smtp.Port = 587;
                    smtp.EnableSsl = true;
                    smtp.Send(mm);
                    mes = "Otp send Sucessfully";
                }
            }

            return mes;
        }
        public ActionResult DepartmentAttendanceApproval()
        {
            return View();
        }
        public ActionResult _sendEmail()
        {
            //SqlConnection conLogin = new SqlConnection(ConfigurationManager.ConnectionStrings["cnLogin"].ConnectionString);
            Random rnd = new Random();
            string mes = string.Empty;
            string password = rnd.Next(111111, 999999).ToString();
            SortedList list = new SortedList();
            list.Add("@CustomerId", Session["Customer"].ToString());
            list.Add("@Email", Session["Email"].ToString());
            list.Add("@Otp", password);
            mes = comfun.executeNonQueryWMessage("stpCustomergetOTP", "", list).ToString();
            if (!mes.ToLower().Contains("error"))
            {

                sendWelcomeEmail(Session["Email"].ToString(), Request.Form["Customer"].ToString() + "Department Attendance Approval OTP :", password);

            }
            return Json(mes);
        }

        public ActionResult _DepartmentAttendanceapproval()
        {

            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@FromDate", Request.Form["FromDate"].ToString());
            list.Add("@Category", Request.Form["CategoryId"].ToString());
            list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
            list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
            list.Add("@BranchId", Request.Form["BranchId"].ToString());
            list.Add("@CustomerId", Session["Customer"].ToString());
            DataTable dt = comfun.fillDataTable("stpEmployeeAttendanceApprovedList", "", list);
            return PartialView("_DepartmentAttendanceapproval", dt);
        }
        public ActionResult _sendVerifyEmail()
        {

            string AttendanceDate = Convert.ToDateTime(Request.Form["AttendanceDate"]).ToString("MMM-yyyy");

            string mes = "";
            //string displayName = "Verification Mail";
            //string toEmail = Session["Email"].ToString();
            //string subject = "Verification Mail";
            string message = "";
            SortedList list2 = new SortedList();
            list2.Add("@CustomerId", Session["Customer"].ToString());
            list2.Add("@Email", Session["Email"].ToString());
            DataTable dt = comfun.fillDataTable("stpCustomerDetail", null, list2);


            //if (dt.Rows.Count > 0)
            //{
            //     message = "Attendance of Respective Department"+"  "+  dt.Rows[0]["AccountName"].ToString() + " for Month " + AttendanceDate + " is Approved. ";
            //}
            try
            {
                //    if (siteUrl.Contains("localhost"))
                //    {

                //        toEmail = "munish.kumar@horizontelecom.in";

                //    }

                //    string email = "Wfms@horizontelecom.in";
                //    string password = "wfms@9002"; ;
                //    displayName = "Attendance Verification "+" of  " + dt.Rows[0]["AccountName"].ToString()+" Department ";

                //    var loginInfo = new NetworkCredential(email, password);
                //    var msg = new MailMessage();
                //    var smtpClient = new SmtpClient("smtp.rediffmailpro.com");
                //      smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                //    smtpClient.Port = 587;

                //    msg.From = new MailAddress("hrms@wfms.in", displayName);
                //    msg.To.Add(new MailAddress(toEmail));
                //    msg.CC.Add(new MailAddress("munish.kumar@horizontelecom.in"));
                //    msg.CC.Add(new MailAddress("kuldeep.singh@horizontelecom.in"));
                //    HtmlString htmlString = new HtmlString(message);
                //             msg.Subject = subject;
                //    msg.Body = message;
                //    msg.ReplyTo = new MailAddress("hrms@wfms.in");
                //    msg.Sender = new MailAddress("hrms@wfms.in", displayName);
                //    msg.IsBodyHtml = true;

                //    smtpClient.EnableSsl = true;
                //    smtpClient.UseDefaultCredentials = true;
                //    smtpClient.Credentials = loginInfo;
                //    smtpClient.Timeout = 600000;
                //    smtpClient.Send(msg);
                //    m = "email sent";

                using (MailMessage mm = new MailMessage("hrms@wfms.in", Session["Email"].ToString()))
                {
                    mm.Subject = "Attendance Verification " + " of  " + dt.Rows[0]["AccountName"].ToString() + " Department ";

                    HtmlString htmlString = new HtmlString(mes);

                    StreamReader reader = new StreamReader(Server.MapPath("~/EmailTemplates/DepartmentAttendaceApproval.htm"));
                    mes = reader.ReadToEnd();

                    mes = mes.Replace("#MonthYear#", AttendanceDate);
                    mes = mes.Replace("#Department#", Request.Form["Customer"].ToString());

                    mes = mes.Replace("#Portal#", "wfms.htistelecom.in");



                    htmlString = new HtmlString(mes);
                    reader.Close();
                    reader.Dispose();
                    mm.Body = mes;

                    mm.CC.Add(new MailAddress("munish.kumar@horizontelecom.in"));
                    //  mm.CC.Add(new MailAddress("kuldeep.singh@horizontelecom.in"));  
                    mm.IsBodyHtml = true;
                    using (SmtpClient smtp = new SmtpClient())
                    {
                        smtp.Host = "smtp.rediffmailpro.com";
                        smtp.UseDefaultCredentials = true;
                        smtp.Credentials = new NetworkCredential
                        {
                            UserName = "Wfms@horizontelecom.in",
                            Password = "wfms@9002"
                        };
                        smtp.Port = 587;
                        smtp.EnableSsl = true;
                        smtp.Send(mm);
                        mes = "email sent";
                    }
                }



            }
            catch (Exception ex)
            {
                mes = "Error : " + ex.Message;
            }


            return Json(mes);
        }

        public ActionResult UpdateMEMINO()
        {


            return View();
        }
        public JsonResult _MEMISubmit()
        {


            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {

                list.Add("@EmpId", Request.Form["EmpId"].ToString());

                mes = comfun.executeNonQueryWMessage("stpUpdateMEMINo", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = "Error : " + ex.Message;
            }

            return Json(mes);
        }
        public ActionResult DepartmentAttendancApprovalPending()
        {
            return View();
        }
        public ActionResult _DepartmentAttendanceapprovalPending()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@MonthYear", Request.Form["MonthYear"].ToString());
            list.Add("@ContractorId", Session["Customer"]);
            dt = comfun.fillDataTable("stpDepartmentAttedance_List", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult AttendanceImport()
        {
            return View();

        }
        //public ActionResult _EmployeeDepartmentAttendanceList()
        //{


        //    SortedList list = new SortedList();

        //    list.Add("@EmpId", Request.Form["EmpId"].ToString());
        //    list.Add("@FromDate", Request.Form["FromDate"].ToString());
        //    list.Add("@Category", Request.Form["CategoryId"].ToString());
        //    list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
        //    list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
        //    list.Add("@BranchId", Request.Form["BranchId"].ToString());

        //    DataTable dt = comfun.fillDataTable("stpEmployeeAttendanceImportList", "", list);

        //    var json1 = JsonConvert.SerializeObject(dt);
        //    return Json(json1);
        //}
        //public ActionResult _EmployeeAttendanceList1()
        //{
        //    decimal total_records = 0;
        //    decimal pageno = Convert.ToDecimal(Request.Form["PageNo"]);
        //    decimal pagesize = Convert.ToDecimal(Request.Form["PageSize"]);

        //    SortedList list = new SortedList();
        //    list.Add("@PageNo", pageno);
        //    list.Add("@PageSize", pagesize);
        //    list.Add("@EmpId", Request.Form["EmpId"].ToString());
        //    list.Add("@FromDate", Request.Form["FromDate"].ToString());
        //    list.Add("@Category", Request.Form["CategoryId"].ToString());
        //    list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
        //    list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
        //    list.Add("@BranchId", Request.Form["BranchId"].ToString());
        //    list.Add("@CircleId", Request.Form["CircleId"].ToString());
        //    DataTable dt = comfun.fillDataTable("stpEmployeeAttendanceMannualImportList", "", list);
        //    if (dt.Rows.Count > 0)
        //        total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

        //    //string paging = comfun.create_paging(total_records, pagesize, pageno, "home", "Index");
        //    string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "Attendance", "_EmployeeAttendanceList1", "_EmployeeAttendanceList1");
        //    HtmlString htm = new HtmlString(paging);
        //    ViewData["paging"] = htm;
        //    var json1 = JsonConvert.SerializeObject(dt);
        //    return Json(new
        //    {
        //        Paging = paging,
        //        data = json1
        //    });
        //}
        public JsonResult _AttendanceImportSubmit(string data, string AttendanceDate)
        {
            string mes = string.Empty;
            try
            {
                DataTable table = (DataTable)JsonConvert.DeserializeObject(data, (typeof(DataTable)));


                SqlCommand com = new SqlCommand();
                connection conObj = new connection();

                com.Connection = conObj.con;

                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "stpEmployeeAttendanceImport_Accept";
                SortedList list = new SortedList();
                com.Parameters.AddWithValue("@AttendanceDate", AttendanceDate);
                com.Parameters.AddWithValue("@Month", Convert.ToDateTime(AttendanceDate).ToString("MMyyyy"));
                com.Parameters.AddWithValue("@CreateBy", Session["EmpId"].ToString());
                com.Parameters.AddWithValue("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));


                // com.Parameters.AddWithValue("@fdWeeklyoffDate", attendanceDate);
                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@AttendanceBulkEmp";
                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table;
                com.Parameters.Add(parameter);
                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);
                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();
                com.ExecuteNonQuery();

                conObj.con.Close();

                mes = sp.Value.ToString();
            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
            }
            return Json(mes);

        }
        public ActionResult ImportFile()
        {
            string mes = "";
            DateTime today = DateTime.UtcNow;
            var todays = today.ToString("dd/MMM/yyyy");

            var postedFile = System.Web.HttpContext.Current.Request.Files["ExcelFile"];
            DataTable dtExcel = new DataTable();

            dtExcel.Columns.Add("EmpId", typeof(string));
            dtExcel.Columns.Add("EmpName", typeof(string));
            dtExcel.Columns.Add("Designation", typeof(string));
            dtExcel.Columns.Add("PresentDays", typeof(string));
            dtExcel.Columns.Add("AbsentDays", typeof(int));
            dtExcel.Columns.Add("LeaveDays", typeof(string));
            dtExcel.Columns.Add("OTDays", typeof(string));
            dtExcel.Columns.Add("OTHours", typeof(string));
            dtExcel.Columns.Add("Days", typeof(string));

            //string mes = "0";

            List<string> data = new List<string>();
            if (postedFile != null)
            {
                // tdata.ExecuteCommand("truncate table OtherCompanyAssets");  
                if (postedFile.ContentType == "application/vnd.ms-excel" || postedFile.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                {
                    string filename = postedFile.FileName;
                    string targetpath = Server.MapPath("~/Content/");
                    postedFile.SaveAs(targetpath + filename);
                    string pathToExcelFile = targetpath + filename;
                    var connectionString = "";
                    if (filename.EndsWith(".xls"))
                    {
                        connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", pathToExcelFile);
                    }
                    else if (filename.EndsWith(".xlsx"))
                    {
                        connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\";", pathToExcelFile);
                    }

                    var adapter = new OleDbDataAdapter("SELECT * FROM [Sheet1$]", connectionString);
                    var ds = new DataSet();

                    adapter.Fill(ds, "ExcelTable");
                    DataRow dr;
                    foreach (DataRow item in ds.Tables["ExcelTable"].Rows)
                    {
                        dr = dtExcel.NewRow();

                        dr["EmpId"] = item["EmpId"];
                        dr["EmpName"] = item["EmpName"];
                        dr["Designation"] = item["Designation"];
                        dr["PresentDays"] = item["PresentDays"];
                        dr["AbsentDays"] = item["AbsentDays"];
                        dr["LeaveDays"] = item["LeaveDays"];
                        dr["OTDays"] = item["OTDays"];
                        dr["OTHours"] = item["OTHours"];
                        dr["Days"] = item["Days"];

                        dtExcel.Rows.Add(dr);
                    }
                    string sheetName = "Sheet1";

                    SqlCommand com = new SqlCommand();
                    connection conObj = new connection();

                    com.Connection = conObj.con;


                    com.CommandType = CommandType.StoredProcedure;
                    com.CommandText = "stpEmployeeAttendanceImportexcel_Accept1";
                    SortedList list = new SortedList();
                    com.Parameters.AddWithValue("@AttendanceDate", Request.Form["AttendanceDate"].ToString());
                    com.Parameters.AddWithValue("@Month", Convert.ToDateTime(Request.Form["AttendanceDate"].ToString()).ToString("MMyyyy"));
                    com.Parameters.AddWithValue("@CreateBy", Session["EmpId"].ToString());
                    com.Parameters.AddWithValue("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));


                    // com.Parameters.AddWithValue("@fdWeeklyoffDate", attendanceDate);
                    SqlParameter parameter = new SqlParameter();
                    parameter.ParameterName = "@AttendanceBulkEmpImport";
                    parameter.SqlDbType = System.Data.SqlDbType.Structured;
                    parameter.Value = dtExcel;
                    com.Parameters.Add(parameter);

                    //Session["Excel"] = dtExcel;
                    SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                    sp.Direction = ParameterDirection.Output;
                    com.Parameters.Add(sp);
                    if (conObj.con.State == ConnectionState.Closed)
                        conObj.con.Open();
                    mes = com.ExecuteNonQuery().ToString();
                    //Session["ExcelId"] = mes;
                    conObj.con.Close();
                    mes = sp.Value.ToString();
                    ///mes = "File uploaded successfully";
                }
            }
            else
            {
                mes = "Error: Please select file";
            }
            //}
            // return Json(mes);
            return Json(mes, JsonRequestBehavior.AllowGet);
        }
        public ActionResult City()
        {
            comfun.saveformname("City", "/TimeSheet/City", "Master/Address/City", "City  Main Form", "Y", "City", "City", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("City");

            DataTable dt = comfun.fillDataTable("CountryState_Select", "", null);
            return View(dt);
        }
        public ActionResult _District_selectddl()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@StateId", Request.Form["StateId"].ToString());

            dt = comfun.fillDataTable("stpDistrictddl", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult _city_select()
        {
            comfun.saveformname("_city_select", "/TimeSheet/_city_select", "Master/Address/City", "City  List", "N", "City", "City", "List", 4);
            SortedList list = new SortedList();
            list.Add("@StateId", Request.Form["StateId"].ToString());
            list.Add("@DistrictId", Request.Form["DistrictId"].ToString());
            DataTable dt = comfun.fillDataTable("City_SelectByStateDistrictId", "", list);

            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        //For City Save/Update Submit//
        public JsonResult _CitiesSaveUpdateSubmit()
        {
            comfun.saveformname("_CitiesSaveUpdateSubmit", "/timeSheet/_CitiesSaveUpdateSubmit", "Master/Address/City", "City  Add", "N", "City", "City", "Add", 2);
            comfun.saveformname("_CitiesSaveUpdateSubmit", "/timeSheet/_CitiesSaveUpdateSubmit", "Master/Address/City", "City  Edit", "N", "City", "City", "Edit", 3);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@CityId", Request.Form["CityId"].ToString());
                list.Add("@CityCode", Request.Form["CityCode"].ToString());
                list.Add("@CityName", Request.Form["CityName"].ToString());
                list.Add("@StateId", Request.Form["StateId"].ToString());
                list.Add("@DistrictId", Request.Form["DistrictId"].ToString());
                list.Add("@IsActive", "Y");
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                list.Add("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                mes = comfun.executeNonQueryWMessage("sp_CitySaveUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
                //  mes = comfun.errorMessage("_CitiesAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult _CitiesEdit()
        {
            SortedList list = new SortedList();
            list.Add("@CityId", Request.Form["CityId"].ToString());
            DataTable dt = comfun.fillDataTable("City_SelectWithId", "", list);
            return PartialView("_CitiesEdit", dt);
        }
        public ActionResult CustomerDetail()
        {
            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("stpAccountHeadddl", "", null);
            return View(dt);
        }

        public ActionResult _AccountDetailList()
        {
            decimal total_records = 0;
            decimal pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            decimal pagesize = Convert.ToDecimal(Request.Form["PageSize"]);
            SortedList list = new SortedList();
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            DataTable dt = comfun.fillDataTable("stpAccountCustomerDetailList", "", list);
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "TimeSheet", "_AccountDetailList", "_AccountDetailList");
            var json = JsonConvert.SerializeObject(dt);
            return Json(new
            {
                Paging = paging,
                data = json
            });
        }


        public JsonResult _AccountSubmit()
        {
            string mes = string.Empty;

            SortedList list = new SortedList();
            try
            {
                string Customer_Logo = "";
                string CustomerLogo = "";
                string _comPathCustomerLogo = "";
                string Customer_Signature = "";
                string CustomerSignature = "";
                string _comPathCustomerSignature = "";

                if (Request.Files.Count > 0)
                {

                    var files = Request.Files;

                    for (int i = 0; i < files.Count; i++)
                    {
                        HttpPostedFileBase file = files[i];

                        // Checking for Internet Explorer  
                        if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                        {
                            string[] testfiles = file.FileName.Split(new char[] { '\\' });
                            CustomerSignature = testfiles[testfiles.Length - 1];

                        }
                        else
                        {
                            if (i == 0)
                            {
                                CustomerLogo = file.FileName;
                                var _ext = Path.GetExtension(CustomerLogo);
                                Customer_Logo = Path.GetFileNameWithoutExtension(CustomerLogo);
                                // Get the complete folder path and store the file inside it.  
                                Customer_Logo = Customer_Logo + _ext;
                                string filePath = Path.Combine(Server.MapPath("/CustomerLogo/") + Customer_Logo);
                                CustomerLogo = filePath;
                                _comPathCustomerLogo = filePath;
                                file.SaveAs(_comPathCustomerLogo);
                            }
                            if (i == 1)
                            {
                                CustomerSignature = file.FileName;
                                var _ext = Path.GetExtension(CustomerSignature);
                                Customer_Signature = Path.GetFileNameWithoutExtension(CustomerSignature);
                                // Get the complete folder path and store the file inside it.  
                                Customer_Signature = Customer_Signature + _ext;
                                string filePath = Path.Combine(Server.MapPath("/CustomerSign/") + Customer_Signature);
                                CustomerSignature = filePath;
                                _comPathCustomerSignature = filePath;
                                file.SaveAs(_comPathCustomerSignature);
                            }

                        }
                    }
                }


                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@fvAccountName", Request.Form["fvAccountName"].ToString());
                list.Add("@fiAccountHeadId", Request.Form["fiAccountHeadId"].ToString());
                list.Add("@fnOpeningBalance", Request.Form["fnOpeningBalance"].ToString());
                list.Add("@fvGSTNo", Request.Form["fvGSTNo"].ToString());
                list.Add("@fvAccountAddress", Request.Form["fvAccountAddress"].ToString());
                list.Add("@fiCountryId", Request.Form["fiCountryId"].ToString());
                list.Add("@fiStateID", Request.Form["fiStateID"].ToString());
                list.Add("@fiCityID", Request.Form["fiCityID"].ToString());
                list.Add("@fvZipPostalCode", Request.Form["fvZipPostalCode"].ToString());
                list.Add("@fvMobileNumber", Request.Form["fvMobileNumber"].ToString());
                list.Add("@fvEMail", Request.Form["fvEMail"].ToString());
                list.Add("@fvPhoneOffice", Request.Form["fvPhoneOffice"].ToString());
                list.Add("@fvContactPerson", Request.Form["fvContactPerson"].ToString());
                list.Add("@DistrictId", Request.Form["districtId"].ToString());
                list.Add("@fvPanNumber", Request.Form["fvPanNumber"].ToString());
                list.Add("@fvOpeningBalanceType", Request.Form["fvOpeningBalanceType"].ToString());
                list.Add("@EntityId", Request.Form["EntityId"].ToString());
                list.Add("@ContractRefNo", Request.Form["ContractRefNo"].ToString());
                list.Add("@FromDate", Request.Form["FromDate"].ToString());
                //list.Add("@ToDate", Request.Form["ToDate"].ToString());
                //list.Add("@Attachment", Request.Form["Attachment"].ToString());
                list.Add("@Logo", CustomerLogo);
                list.Add("@Signature", CustomerSignature);
                mes = comfun.executeNonQueryWMessage("stpAccountDetailCustomerSave_update", "", list).ToString();
            }
            catch (Exception ex)
            {

                mes = "Error:" + ex.Message;
            }
            return Json(mes);
        }
        public ActionResult _AccountDetailEdit()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("stpAccountDetailEdit", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);

        }
        public JsonResult _AccountDetailStatusUpdate()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpAccountStatusUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }
        public ActionResult Department()
        {
            comfun.saveformname("Department", "/TimeSheet/Department", "Department", "Department  form", "Y", "", "Department", "List");
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            return View();
        }
        public ActionResult _DepartmentJsonList()
        {
            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 50;
            if (Request.Form["Page"] != null)
            {
                pagesize = Convert.ToDecimal(Request.Form["Page"]);
            }
            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }
            SortedList list = new SortedList();
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            DataTable dt = comfun.fillDataTable("stpCustomerEntity_List", "", null);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }

        //[customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]

        public ActionResult _DepartmentSubmitView()
        {
            DataTable dt = comfun.fillDataTable("stpCustomerEntity_Accept", "", null);
            // comfun.saveformname("CountriesAdd", "/admin/CountriesAdd", "Country Add", "Country view", "N", "Country Add", "Country", "Add");
            ViewData["Id"] = "0";
            ViewData["EntityCode"] = "";
            ViewData["EntityName"] = "";
            if (Request.Form["Id"].ToString() != "0")
            {
                ViewData["Id"] = Request.Form["Id"].ToString();
                ViewData["EntityCode"] = Request.Form["EntityCode"].ToString();
                ViewData["EntityName"] = Request.Form["EntityName"].ToString();
                //ViewData["Shiftids"] = Request.Form["Shiftids"].ToString();
            }
            return PartialView("_CustomerTypeSubmit", dt);
        }

        public JsonResult _DepartmentSubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@EntityCode", Request.Form["EntityCode"].ToString());
                list.Add("@EntityName", Request.Form["EntityName"].ToString());
                //list.Add("@Shiftids", Request.Form["Shiftids"].ToString());
                //list.Add("@CreatedBy", "1");
                //list.Add("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                mes = comfun.executeNonQueryWMessage("stpCustomerEntity_Accept", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
            }
            return Json(mes);
        }
        //public ActionResult ddlCustomerTypeselect()
        //{
        //    DataTable dt = comfun.fillDataTable("stpSHIFT_Select", "", null);
        //    return PartialView("_ddlCustomerTypeselect", dt);
        //}
        public JsonResult _DepartmentStatusUpdate()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpCustomerEntity_Status", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }
        public ActionResult _DepartmentEdit()
        {
            string Id = "0";
            Id = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@Id", Id);
            DataTable dt = comfun.fillDataTable("stpCustomerEntity_Edit", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult DownloadAttendanceCertificate()
        {
            comfun.saveformname("DownloadAttendanceCertificate", "/TimeSheet/DownloadAttendanceCertificate", "TimeSheet/DownloadAttendanceCertificate", "Download Attendance Certificate form", "Y", "DownloadAttendanceCertificate", "DownloadAttendanceCertificate", "View", 1);
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            return View();
        }
        public ActionResult _DepartmentgenratedCertificateList()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            if (Request.Form["FromDate"] != null)
            {
                list.Add("@FromDate", Request.Form["FromDate"].ToString());
            }
            // list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
            if (Request.Form["DepartmentId"].ToString() != "0")
            {
                list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
            }
            else
            {
                list.Add("@DepartmentId", Session["Customer"].ToString());
            }
            list.Add("@BranchId", Request.Form["BranchId"].ToString());
            list.Add("@LoginId", Session["EmpId"].ToString());
            dt = comfun.fillDataTable("stpEmployeegenratedCertificateList", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult EmployeeInfo()
        {
            return View();
        }
        public ActionResult _EmployeeInfo()
        {
            SortedList list = new SortedList();


            if (Request.Form["DepartmentId"].ToString() != "0")
            {
                list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
            }
            else
            {
                list.Add("@DepartmentId", Session["Customer"].ToString());
            }
            list.Add("@BranchId", Request.Form["BranchId"].ToString());
            list.Add("@DistrictId", Request.Form["DistrictId"].ToString());
            DataTable dt = comfun.fillDataTable("stpEmployeeInfo_HPSEDC", "", list);
            return PartialView("_SalarySheet", dt);
        }
        public ActionResult Vendor()
        {
            DataTable dt = comfun.fillDataTable("stpCustomerTypeddl", "", null);
            return View(dt);
        }
        public ActionResult SupplierDetail()
        {
            comfun.saveformname("SupplierDetail", "/Pm/SupplierDetail", "SupplierDetail", "SupplierDetail  form", "Y", "", "SupplierDetail", "List");
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            DataTable dt = comfun.fillDataTable("stpCustomerTypeddl", "", null);

            return View(dt);
        }
        public ActionResult _SupplierDetailList()
        {
            decimal total_records = 0;
            decimal pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            decimal pagesize = Convert.ToDecimal(Request.Form["PageSize"]);
            SortedList list = new SortedList();
            list.Add("@pageNo", pageno);
            list.Add("@pageSize", pagesize);
            DataTable dt = comfun.fillDataTable("stpSupplierList", "", list);
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "TimeSheet", "_SupplierDetailList", "_SupplierDetailList");
            var json = JsonConvert.SerializeObject(dt);
            return Json(new
            {
                Paging = paging,
                data = json
            });

        }

        public JsonResult _SpplierDetailSubmit(string persons, string SupplierId, string SupplierName, string SuppliertypeID, string SupplierCode, string MobileNo, string EMail, string PanNumber, string TinNumber, string EntityId,
        string GSTIN, string CIN, string RegdOffice, string RegedDistrictId, string RegedCityId, string IsActive, string Bills, string Shipping, string social, string RegdPostalCode)
        {
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                string Vendor_Logo = "";
                string Vendorlogo = "";
                string _comPathVendorlogo = "";
                string Vendor_Signature = "";
                string VendorSignature = "";
                string _comPathVendorSignature = "";

                if (Request.Files.Count > 0)
                {
                    var files = Request.Files;

                    for (int i = 0; i < files.Count; i++)
                    {
                        HttpPostedFileBase file = files[i];
                        // Checking for Internet Explorer
                        if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                        {
                            string[] testfiles = file.FileName.Split(new char[] { '\\' });
                            VendorSignature = testfiles[testfiles.Length - 1];

                        }
                        else
                        {
                            if (i == 0)
                            {
                                Vendorlogo = file.FileName;
                                var _ext = Path.GetExtension(Vendorlogo);
                                Vendor_Logo = Path.GetFileNameWithoutExtension(Vendorlogo);
                                // Get the complete folder path and store the file inside it.  
                                Vendor_Logo = Vendor_Logo + _ext;
                                string filePath = Path.Combine(Server.MapPath("/CustomerLogo/") + Vendor_Logo);
                                Vendorlogo = filePath;
                                _comPathVendorlogo = filePath;
                                file.SaveAs(_comPathVendorlogo);
                            }
                            if (i == 1)
                            {
                                VendorSignature = file.FileName;
                                var _ext = Path.GetExtension(VendorSignature);
                                Vendor_Signature = Path.GetFileNameWithoutExtension(VendorSignature);
                                // Get the complete folder path and store the file inside it.  
                                Vendor_Signature = Vendor_Signature + _ext;
                                string filePath = Path.Combine(Server.MapPath("/CustomerSign/") + Vendor_Signature);
                                VendorSignature = filePath;
                                _comPathVendorSignature = filePath;
                                file.SaveAs(_comPathVendorSignature);
                            }
                        }

                    }

                }
                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject(persons);
                DataTable table = new DataTable();
                table.Columns.Add("PersonName", typeof(string));
                table.Columns.Add("MobileNo", typeof(string));
                table.Columns.Add("Email", typeof(string));
                table.Columns.Add("DesignationId", typeof(int));
                foreach (var item in jsonData)
                {
                    DataRow dr = table.NewRow();
                    dr["PersonName"] = item.PersonName;
                    dr["MobileNo"] = item.PersonContactNo;
                    dr["Email"] = item.PersonEmailId;
                    table.Rows.Add(dr);
                }
                dynamic jsonData1 = Newtonsoft.Json.JsonConvert.DeserializeObject(Bills);
                DataTable table1 = new DataTable();
                table1.Columns.Add("BillAddress", typeof(string));
                table1.Columns.Add("BillPostalCode", typeof(string));
                table1.Columns.Add("BillCountry", typeof(int));
                table1.Columns.Add("BillState", typeof(int));
                table1.Columns.Add("BillDistrict", typeof(int));

                table1.Columns.Add("BillCity", typeof(int));
                foreach (var item1 in jsonData1)
                {
                    DataRow dr1 = table1.NewRow();
                    dr1["BillAddress"] = item1.BillAddress;
                    dr1["BillPostalCode"] = item1.BillPostalCode;
                    dr1["BillCountry"] = item1.BillCountry;
                    dr1["BillState"] = item1.BillState;
                    dr1["BillDistrict"] = item1.BillDistrict;
                    dr1["BillCity"] = item1.BillCity;

                    table1.Rows.Add(dr1);
                }
                dynamic jsonData2 = Newtonsoft.Json.JsonConvert.DeserializeObject(Shipping);
                DataTable table2 = new DataTable();
                table2.Columns.Add("ShipAddress", typeof(string));
                table2.Columns.Add("ShipPostalCode", typeof(string));
                table2.Columns.Add("Country", typeof(int));
                table2.Columns.Add("State", typeof(int));
                table2.Columns.Add("District", typeof(int));
                table2.Columns.Add("City", typeof(int));
                foreach (var item2 in jsonData2)
                {
                    DataRow dr2 = table2.NewRow();
                    dr2["ShipAddress"] = item2.ShipAddress;
                    dr2["ShipPostalCode"] = item2.ShipPostalCode;
                    dr2["Country"] = item2.Country;
                    dr2["State"] = item2.State;
                    dr2["District"] = item2.District;
                    dr2["City"] = item2.City;

                    table2.Rows.Add(dr2);
                }
                dynamic jsonData3 = Newtonsoft.Json.JsonConvert.DeserializeObject(social);
                DataTable table3 = new DataTable();
                table3.Columns.Add("SocialAddress", typeof(string));

                table3.Columns.Add("TypeID", typeof(int));

                foreach (var item3 in jsonData3)
                {
                    DataRow dr3 = table3.NewRow();
                    dr3["SocialAddress"] = item3.MediaAddress;
                    dr3["TypeId"] = item3.MediaType;


                    table3.Rows.Add(dr3);
                }


                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "stpVendorDetail_Accept_Update";
                com.Parameters.AddWithValue("@SupplierId", SupplierId);
                com.Parameters.AddWithValue("@SupplierName", SupplierName);
                com.Parameters.AddWithValue("@SuppliertypeID", SuppliertypeID);
                com.Parameters.AddWithValue("@SupplierCode", SupplierCode);
                com.Parameters.AddWithValue("@MobileNo", MobileNo);
                com.Parameters.AddWithValue("@EMail", EMail);
                com.Parameters.AddWithValue("@PanNumber", PanNumber);
                com.Parameters.AddWithValue("@TinNumber", TinNumber);
                com.Parameters.AddWithValue("@EntityId", EntityId);
                com.Parameters.AddWithValue("@GSTIN", GSTIN);
                com.Parameters.AddWithValue("@RegdOffice", RegdOffice);
                com.Parameters.AddWithValue("@CIN", CIN);
                com.Parameters.AddWithValue("@RegedCityId", RegedCityId);
                com.Parameters.AddWithValue("@RegedDistrictId", RegedDistrictId);
                com.Parameters.AddWithValue("@IsActive", IsActive);
                com.Parameters.AddWithValue("@RegdPostalCode", RegdPostalCode);
                com.Parameters.AddWithValue("@Logo", Vendor_Logo);
                com.Parameters.AddWithValue("@Signature", Vendor_Signature);
                // com.Parameters.AddWithValue("@fdWeeklyoffDate", attendanceDate);
                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@PersonTable";
                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table;
                com.Parameters.Add(parameter);
                SqlParameter parameter1 = new SqlParameter();
                parameter1.ParameterName = "@BillAddress";
                parameter1.SqlDbType = System.Data.SqlDbType.Structured;
                parameter1.Value = table1;
                com.Parameters.Add(parameter1);
                SqlParameter parameter2 = new SqlParameter();
                parameter2.ParameterName = "@ShippingAddress";
                parameter2.SqlDbType = System.Data.SqlDbType.Structured;
                parameter2.Value = table2;
                com.Parameters.Add(parameter2);
                SqlParameter parameter3 = new SqlParameter();
                parameter3.ParameterName = "@SocialAddress";
                parameter3.SqlDbType = System.Data.SqlDbType.Structured;
                parameter3.Value = table3;
                com.Parameters.Add(parameter3);
                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);


                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();

                com.ExecuteNonQuery();

                conObj.con.Close();

                mes = sp.Value.ToString();

            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
            }
            return Json(mes);
        }
        public ActionResult CustomerInvoiceList()
        {
            return View();
        }
        public ActionResult _CustomerInvoiceList()
        {
            comfun.saveformname("_CustomerInvoiceList", "/TimeSheet/_CustomerInvoiceList", "Master/Address/CustomerInvoiceList", "Customer Invoice  list", "N", "CustomerInvoiceList", "CustomerInvoiceList", "List", 4);
            SortedList list = new SortedList();
            list.Add("@LoginId", Session["loginID"].ToString());
            list.Add("@MonthYear", Request.Form["MonthYear"].ToString());
            DataTable dt = comfun.fillDataTable("stpDepartmentInvoiceList", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult CustomerInvoicePayment()
        {
            return View();
        }
        public ActionResult _CustomerInvoicePayment()
        {
            comfun.saveformname("_CustomerInvoicePayment", "/TimeSheet/_CustomerInvoicePayment", "Attendancemanagement/_CustomerInvoicePayment", "Customer Invoice  list", "N", "_CustomerInvoicePayment", "_CustomerInvoicePayment", "List", 4);
            SortedList list = new SortedList();
            list.Add("@LoginId", Session["loginID"].ToString());
            list.Add("@MonthYear", Request.Form["MonthYear"].ToString());
            list.Add("@StatusId", Request.Form["StatusId"].ToString());
            DataTable dt = comfun.fillDataTable("stpDepartmentInvoicePaymentStatus", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult StudentInfo()
        {
            return View();
        }
        public ActionResult _StudentLocddl()
        {
            DataTable dt = comfun.fillDataTable("stpStudentLocddl", "", null);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);

        }
        public ActionResult _StudentJsonList()
        {
            SortedList list = new SortedList();
            list.Add("@locationID", Request.Form["locationID"].ToString());
            DataTable dt = comfun.fillDataTable("stpStudent_List", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public JsonResult _StudentEditSubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@ID", Request.Form["Id"].ToString());
                list.Add("@StudentName", Request.Form["StudentName"].ToString());
                list.Add("@UsrID", Session["Loginid"]);
                list.Add("@fatherName", Request.Form["FatherName"].ToString());
                list.Add("@MotherName", Request.Form["MotherName"].ToString());
                list.Add("@Email", Request.Form["Email"].ToString());
                list.Add("@MobileNo", Request.Form["MobileNo"].ToString());
                list.Add("@LocationID", Request.Form["location"].ToString());
                //list.Add("@EntityName", Request.Form["IsActive"].ToString());
                //list.Add("@Shiftids", Request.Form["Shiftids"].ToString());
                //list.Add("@CreatedBy", "1");
                //list.Add("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                mes = comfun.executeNonQueryWMessage("stpStudent_Accept", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
            }
            return Json(mes);
        }
        //public ActionResult ddlCustomerTypeselect()
        //{
        //    DataTable dt = comfun.fillDataTable("stpSHIFT_Select", "", null);
        //    return PartialView("_ddlCustomerTypeselect", dt);
        //}
        public JsonResult _StudentStatusUpdate()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }
        public ActionResult _StudentEdit()
        {
            string Id = "0";
            Id = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@Id", Id);
            DataTable dt = comfun.fillDataTable("", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        #region vipul
        public ActionResult CustomerContract()
        {
            return View();
        }
        public ActionResult CustomerVendorRequisition()
        {
            return View();
        }

        public JsonResult CustomerVendorRequisitionList()
        {

            SortedList list = new SortedList();
            list.Add("@MappingId", Request.Form["MappingId"].ToString());
            DataTable ds = comfun.fillDataTable("stpHpsedcCVMaping_List", "", list);
            var json = JsonConvert.SerializeObject(ds);
            return Json(json);
        }

        public JsonResult CustomerVendorRequisitionSubmit()
        {
            string mes = string.Empty;

            SortedList list = new SortedList();
            try
            {
                string Applicant_image = "";
                string ApplicantImage = "";
                string _comPathApplicant = "";
                string RationCard_image = "";
                string RationCardImage = "";
                string _comPathRationCard = "";

                if (Request.Files.Count > 0)
                {

                    var files = Request.Files;

                    for (int i = 0; i < files.Count; i++)
                    {
                        HttpPostedFileBase file = files[i];

                        // Checking for Internet Explorer  
                        if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                        {
                            string[] testfiles = file.FileName.Split(new char[] { '\\' });
                            RationCardImage = testfiles[testfiles.Length - 1];
                            //EmpImage2 = testfiles[testfiles.Length - 1];

                        }
                        else
                        {
                            if (i == 0)
                            {
                                ApplicantImage = file.FileName;
                                var _ext = Path.GetExtension(ApplicantImage);
                                Applicant_image = Path.GetFileNameWithoutExtension(ApplicantImage);
                                // Get the complete folder path and store the file inside it.  
                                Applicant_image = Applicant_image + _ext;
                                string filePath = Path.Combine(Server.MapPath("/AggrementDocument/") + Applicant_image);
                                ApplicantImage = filePath;
                                _comPathApplicant = filePath;
                                file.SaveAs(_comPathApplicant);
                            }
                            if (i == 1)
                            {
                                RationCardImage = file.FileName;
                                var _ext = Path.GetExtension(RationCardImage);
                                RationCard_image = Path.GetFileNameWithoutExtension(RationCardImage);
                                // Get the complete folder path and store the file inside it.  
                                RationCard_image = RationCard_image + _ext;
                                string filePath = Path.Combine(Server.MapPath("/RequisitionLetter/") + RationCard_image);
                                RationCardImage = filePath;
                                _comPathRationCard = filePath;
                                file.SaveAs(_comPathRationCard);
                            }

                        }
                    }
                }
                list.Add("@MappingId", Request.Form["MappingId"].ToString());
                list.Add("@CustomerId", Request.Form["CustomerId"].ToString());
                list.Add("@VendorId", Request.Form["VendorId"].ToString());
                list.Add("@AggrementRefNo", Request.Form["AggrementRefNo"].ToString());
                list.Add("@AggrementDate", Request.Form["AggrementDate"].ToString());
                list.Add("@FromDate", Request.Form["FromDate"].ToString());
                list.Add("@ToDate", Request.Form["ToDate"].ToString());
                list.Add("@RequisitionDate", Request.Form["RequisitionDate"].ToString());
                list.Add("@ReqRefNo", Request.Form["RequisitionReferenceNo"].ToString());
                list.Add("@AggrementDoc", Applicant_image);
                list.Add("@RequisitionLetter", RationCard_image);
                list.Add("@CreatedBy", Session["EmpId"]);
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpHpsedcCVMaping_AcceptUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = "Error:" + ex.Message;
            }
            return Json(mes);
        }


        public JsonResult CustomerVendorRequisitionEdit()
        {

            SortedList list = new SortedList();
            list.Add("@MappingId", Request.Form["MappingId"].ToString());
            // list.Add("@CustomerId", Request.Form["CustomerId"].ToString());
            DataTable ds = comfun.fillDataTable("stpHpsedcCVMaping_List", "", list);
            var json = JsonConvert.SerializeObject(ds);
            return Json(json);

        }
        public ActionResult CutomerContract()
        {
            return View();
        }
        #endregion
        public ActionResult LocationwiseMail()
        {
            return View();
        }
        public ActionResult _LocationMailStatus()
        {

            SortedList list = new SortedList();
            list.Add("@locationID", Request.Form["LocationId"].ToString());
            list.Add("@Status", Request.Form["StatusId"].ToString());
            DataTable dt = comfun.fillDataTable("stpStudent_MailStatus", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        //public string SignedQRCode(int StudentIds)
        //{
        //    try
        //    {
        //        string Code = string.Empty;
        //        //Get SignedQRCode from database aganist invoice id
        //        SortedList list = new SortedList();
        //        list.Add("@StudentId", StudentIds);
        //        Code = comfun.executeScaler("getStudentIdQrCode", "", list).ToString();
        //        //using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
        //        //{
        //        //    using (SqlCommand cmd = new SqlCommand("getStudentIdQrCode", con))
        //        //    {
        //        //       cmd.Parameters.Add(new SqlParameter("@StudentId", StudentIds.ToString()));
        //        //        con.Open();
        //        //        object result = cmd.ExecuteScalar();
        //        //        if (result != null)
        //        //        {
        //        //            Code = result.ToString();
        //        //        }
        //        //    }
        //        //}
        //        //string Code = "eyJhbGciOiJSUzI1NiIsImtpZCI6IkVEQzU3REUxMzU4QjMwMEJBOUY3OTM0MEE2Njk2ODMxRjNDODUwNDciLCJ0eXAiOiJKV1QiLCJ4NXQiOiI3Y1Y5NFRXTE1BdXA5NU5BcG1sb01mUElVRWMifQ.eyJkYXRhIjoie1wiU2VsbGVyR3N0aW5cIjpcIjI3QUJGUEQ0MDIxTDAwMlwiLFwiQnV5ZXJHc3RpblwiOlwiMzNBQUJDUjcxMDZHMVpRXCIsXCJEb2NOb1wiOlwiTE9HUy02Njc2N0hUSVNcIixcIkRvY1R5cFwiOlwiSU5WXCIsXCJEb2NEdFwiOlwiMDQvMDQvMjAyMlwiLFwiVG90SW52VmFsXCI6MjM2LFwiSXRlbUNudFwiOjIsXCJNYWluSHNuQ29kZVwiOlwiODUxNzYyMTBcIixcIklyblwiOlwiOWVjNzRlZWQzNWY1NGQ0Njg0YjQxZDMyYTVmMTY4NmE2MWNkZTE5MWU0NzljOGI0N2QzMWI3ZTkxZWRlYjIyZlwiLFwiSXJuRHRcIjpcIjIwMjItMDQtMDUgMTY6Mzc6NTRcIn0iLCJpc3MiOiJOSUMifQ.K3WLA2WTtmnQhF8YouA7d8wHVcLTar12ge3Z8UuiCaFvS7vII75CRsGxwofR6H3OVO16iFRfw6cQeAmdj6l6oNf8-JzmrWE2ggNt2cDM8sEbgIbLZ3BwCpE6_lwL4MmueZz4ts_oapg88s8INvOQjS_SkfHyED4nNveXTPR_eGbMCb6c8Kbqk9rEbU1HIJfijd7Wlzdf9gkdNxsE13X6jvg91OcuzC1rvhWGJRgp7MGDRQ4wwrgt70NjKDDCCfGn_viqnqWo8zIQzuXRvtlzdVWtrNK0qFpVa_19Zs_aJK1051F6ElCsz6Jg-mLkj7t-Ro-v06rNnSMs8W42JidHeQ";
        //        // string Code = Session["Invoice_SignedQRCode"].ToString();            
        //        QRCodeGenerator QrGenerator = new QRCodeGenerator();
        //        QRCodeData QrCodeInfo = QrGenerator.CreateQrCode(Code, QRCodeGenerator.ECCLevel.Q);
        //        QRCode QrCode = new QRCode(QrCodeInfo);
        //        Bitmap QrBitmap = QrCode.GetGraphic(60);
        //        byte[] BitmapArray = QrBitmap.BitmapToByteArray();
        //        byte[] img = null;
        //        System.Web.UI.WebControls.Image imgBarCode = new System.Web.UI.WebControls.Image();
        //        imgBarCode.Height = 25;
        //        imgBarCode.Width = 25;
        //        using (Bitmap bitMap = QrCode.GetGraphic(10))
        //        {
        //            using (MemoryStream ms = new MemoryStream())
        //            {
        //                bitMap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);

        //                img = new byte[ms.ToArray().Length];
        //                img = ms.ToArray();
        //            }

        //        }
        //        SqlConnection con1 = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString);
        //        if (con1.State == ConnectionState.Open)
        //            con1.Close();
        //        con1.Open();
        //        string SQL = "update tblStudent_Info set QrCode = @img  where ID = @invId ";
        //        SqlCommand command = new SqlCommand(SQL, con1);
        //        command.Parameters.AddWithValue("@invId", StudentIds);

        //        command.Parameters.AddWithValue("@img", img);

        //        int i = command.ExecuteNonQuery();
        //        //comf.executeScaler("","update InvoiceResponse set QRCode=" + Convert.ToBase64String(BitmapArray) + "Where InvoiceId=" + invoiceId,null);

        //        return string.Format("data:image/png;base64,{0}", Convert.ToBase64String(BitmapArray));
        //    }
        //    catch (Exception ex)
        //    {

        //        throw ex;
        //    }

        //}
        //public ActionResult PrintQrcode()
        //{
        //    int StudentId;
        //    int.TryParse(Request.QueryString["id"], out StudentId);
        //    string QRCode = SignedQRCode(StudentId);
        //    //comf.executeNonQuery("", "update InvoiceResponse set QRCode=" + QRCode + "Where InvoiceId=" + invoiceId, null);


        //    ViewBag.Title = "Laptop";
        //    string mes = string.Empty;


        //    HtmlString htmlString = new HtmlString(mes);

        //    StreamReader reader = new StreamReader(Server.MapPath("~/EmailTemplates/StudentMail.htm"));
        //    mes = reader.ReadToEnd();
        //    SortedList list = new SortedList();
        //    list.Add("@id", StudentId);
        //    DataTable dt = comfun.fillDataTable("stStudentpGetData", "", list);
        //    mes = mes.Replace("#QRCODE#", QRCode);

        //    mes = mes.Replace("#StudentName#", dt.Rows[0]["StudentName"].ToString());
        //    mes = mes.Replace("#FatherName#", dt.Rows[0]["FatherName"].ToString());
        //    mes = mes.Replace("#MobileNo#", dt.Rows[0]["MobileNo"].ToString());


        //    //for (int i = 0; i < ds.Tables[1].Rows.Count; i++)
        //    //{
        //    //    table += "<tr>";
        //    //    table += "<td>" + (i + 1).ToString() + "</td>";
        //    //    table += "<td style = border:1px solid border-collapse:collapse align=left>" + ds.Tables[1].Rows[i]["PrdDesc"].ToString() + "</td>";
        //    //    table += "<td style = border:1px solid border-collapse:collapse align=left>" + ds.Tables[1].Rows[i]["HsnCd"].ToString() + "</td>";
        //    //    table += "<td style = border:1px solid border-collapse:collapse align=left>" + ds.Tables[1].Rows[i]["Qty"].ToString() + "</td>";
        //    //    table += "<td style = border:1px solid border-collapse:collapse align=left>" + ds.Tables[1].Rows[i]["UnitPrice"].ToString() + "</td>";
        //    //    table += "<td style = border:1px solid border-collapse:collapse align=left>" + ds.Tables[1].Rows[i]["AssAmt"].ToString() + "</td>";
        //    //    table += "<td style = border:1px solid border-collapse:collapse>" + ds.Tables[1].Rows[i]["CGSTPer"].ToString() + "</td>" + "<td style=width:50px>" + ds.Tables[1].Rows[i]["CgstAmt"].ToString() + "</td>";
        //    //    table += "<td style = border:1px solid border-collapse:collapse>" + ds.Tables[1].Rows[i]["SGSTPer"].ToString() + "</td>" + "<td>" + ds.Tables[1].Rows[i]["SgstAmt"].ToString() + "</td>";
        //    //    table += "<td style = border:1px solid border-collapse:collapse>" + ds.Tables[1].Rows[i]["GstRt"].ToString() + "</td>" + "<td>" + ds.Tables[1].Rows[i]["IgstAmt"].ToString() + "</td>";
        //    //    table += "<td style = border:1px solid border-collapse:collapse>" + ds.Tables[1].Rows[i]["TotItemVal"].ToString() + "</td>";

        //    //}
        //    //        < td ></ td >
        //    //<td style = "border: 1px solid; border-collapse: collapse;" >#ItemName#</td>
        //    //<td style = "border: 1px solid; border-collapse: collapse;" >#HSNNo#</td>
        //    //<td style = "border: 1px solid; border-collapse: collapse;" >#Qty#</td>
        //    //<td style = "border: 1px solid; border-collapse: collapse;" >#Rate#</td>
        //    //<td style = "border: 1px solid; border-collapse: collapse;" >#TaxableAmount#</td>
        //    //<td style = "border: 1px solid; border-collapse: collapse;" colspan = "2" >< table style = "border: 1px solid; width: 100%; border-collapse: collapse;" >< tr >< td style = "border: 1px solid; border-collapse: collapse;" >#CGSTper#</td> <td style="border: 1px solid; border-collapse: collapse;">#CGSTAmt#</td></tr></table></td>
        //    //<td style = "border: 1px solid; border-collapse: collapse;" colspan = "2" >< table style = "border: 1px solid; width: 100%; border-collapse: collapse;" >< tr >< td style = "border: 1px solid; border-collapse: collapse;" >#SGSTper#</td> <td style="border: 1px solid; border-collapse: collapse;">#SGSTAmt#</td></tr></table></td>
        //    //<td style = "border: 1px solid; border-collapse: collapse;" colspan = "2" >< table style = "border: 1px solid; width: 100%; border-collapse: collapse;" >< tr >< td style = "border: 1px solid; border-collapse: collapse;" >#IGSTper#</td> <td style="border: 1px solid; border-collapse: collapse;">#IGSTAmt#</td></tr></table></td>
        //    //<td style = "border: 1px solid; border-collapse: collapse;" colspan = "2" >#Amt#</td>
        //    //mes = mes.Replace("#tbody#", table);

        //    //mes = mes.Replace("#tbody#", table);
        //    htmlString = new HtmlString(mes);
        //    ViewData["Template"] = htmlString;
        //    reader.Close();
        //    reader.Dispose();
        //    return View();
        //}

        public ActionResult EmployeeDetailImport()
        {
            return View();

        }
        public ActionResult _EmployeeCtcExcel()
        {

            _exportExcel("Employee Detail Upload");
            return View();
        }
        public void _exportExcel(string excelName)
        {
            string date = DateTime.UtcNow.AddMinutes(330).ToString("yyyyMMdd_HHmm");
            excelName = excelName + "_" + date;
            DataTable dt = new DataTable();
            dt.Columns.Add("Code", typeof(string));
            dt.Columns.Add("Empname", typeof(string));
            dt.Columns.Add("Father's Name", typeof(string));
            dt.Columns.Add("Mother's Name", typeof(string));
            dt.Columns.Add("Date Of Birth", typeof(string));
            dt.Columns.Add("Date of Joining", typeof(string));
            dt.Columns.Add("Desigation", typeof(string));
            dt.Columns.Add("Department", typeof(string));
            dt.Columns.Add("Depatment Location", typeof(string));
            dt.Columns.Add("Vendor", typeof(string));
            dt.Columns.Add("Address", typeof(string));
            dt.Columns.Add("District", typeof(string));
            dt.Columns.Add("Block", typeof(string));
            dt.Columns.Add("Panchayat Name", typeof(string));
            dt.Columns.Add("Postal Code", typeof(string));
            dt.Columns.Add("PAN", typeof(string));
            dt.Columns.Add("UAN", typeof(string));
            dt.Columns.Add("AADHAR", typeof(string));
            dt.Columns.Add("Bank", typeof(string));
            dt.Columns.Add("IFSC", typeof(string));
            dt.Columns.Add("Account No", typeof(string));
            dt.Columns.Add("Basic", typeof(string));
            dt.Columns.Add("HRA", typeof(string));
            dt.Columns.Add("Others", typeof(string));
            dt.Columns.Add("IsPf", typeof(string));
            dt.Columns.Add("IsEsi", typeof(string));
            dt.Columns.Add("Charge(%)", typeof(string));
            dt.Columns.Add("TA", typeof(string));
            dt.Columns.Add("SP", typeof(string));
            dt.Columns.Add("GrossWages", typeof(string));
            dt.Columns.Add("ESINO", typeof(string));


            XLWorkbook wb = new XLWorkbook();

            var ws = wb.Worksheets.Add(dt, "Sheet1");
            ws.SetAutoFilter(false);
            ws.Tables.FirstOrDefault().ShowAutoFilter = false;

            if (dt.Rows.Count > 0)
            {
                ws.Rows(1, 1).Style.Font.Bold = true;
                ws.SetShowGridLines(true);

                ws.Style.Border.OutsideBorder = XLBorderStyleValues.None;
                ws.Style.Border.BottomBorder = XLBorderStyleValues.None;
                Response.Clear();
                Response.Buffer = true;
                Response.Charset = "";

                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=" + excelName + ".xlsx");

                MemoryStream MyMemoryStream = new MemoryStream();

                wb.SaveAs(MyMemoryStream);
                MyMemoryStream.WriteTo(Response.OutputStream);

                Response.Flush();
                Response.End();
                MyMemoryStream.Dispose();
                MyMemoryStream.Close();
            }
            else
            {
                ws.Rows(1, 1).Style.Font.Bold = true;

                Response.Clear();
                Response.Buffer = false;
                Response.Charset = "";
                Response.AddHeader("Content-Type", "application/vnd.ms-excel");
                Response.AddHeader("content-disposition", "attachment;filename=" + excelName + ".xlsx");


                MemoryStream MyMemoryStream = new MemoryStream();

                wb.SaveAs(MyMemoryStream);
                MyMemoryStream.WriteTo(Response.OutputStream);

                Response.Flush();
                Response.End();
                MyMemoryStream.Dispose();
                MyMemoryStream.Close();

            }

        }
        public ActionResult _ImportFile01()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return PartialView("_sessionExpired");
            }
            if (User.Identity.Name == "")
            {
                return PartialView("_sessionExpired");
            }
            string mes = "";
            DateTime today = DateTime.UtcNow;
            var todays = today.ToString("dd/MMM/yyyy");
            var postedFile = System.Web.HttpContext.Current.Request.Files["ExcelFile"];

            //string mes = "0";

            List<string> data = new List<string>();
            if (postedFile != null)
            {
                // tdata.ExecuteCommand("truncate table OtherCompanyAssets");  
                if (postedFile.ContentType == "application/vnd.ms-excel" || postedFile.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                {
                    HttpPostedFileBase fileBase;
                    DataTable dataTable = new DataTable();
                    if (Request.Files.Count > 0)
                    {
                        foreach (string file in Request.Files)
                        {
                            fileBase = Request.Files[file] as HttpPostedFileBase;
                            if (fileBase != null && fileBase.ContentLength > 0)
                            {
                                Stream stream = fileBase.InputStream;
                                IExcelDataReader reader = null;
                                if (fileBase.FileName.EndsWith(".xls"))
                                {
                                    reader = ExcelReaderFactory.CreateBinaryReader(stream);
                                }
                                else if (fileBase.FileName.EndsWith(".xlsx"))
                                {
                                    reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                                }
                                var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                                {
                                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                    {
                                        UseHeaderRow = true
                                    }
                                });
                                dataTable = result.Tables[0];

                                reader.Close();
                            }
                        }
                    }
                    else
                    {
                        ViewBag.FCRewardMessage = "Please select a file for upload.";
                    }


                    DataTable resultCount = InsertFCReward(dataTable);

                    //  string sheetName = "Sheet1";

                    var json = JsonConvert.SerializeObject(resultCount);

                    return Json(json);

                }

                //}
                // return Json(mes);

            }
            else
            {
                mes = "Error: Please select file";
            }
            return Json(mes, JsonRequestBehavior.AllowGet);
        }

        public DataTable InsertFCReward(DataTable importdt)
        {

            commonFunctions comf = new commonFunctions();
            var postedFile = System.Web.HttpContext.Current.Request.Files["ExcelFile"];
            DataTable dtExcel = new DataTable();

            dtExcel.Columns.Add("Code", typeof(string));
            dtExcel.Columns.Add("Empname", typeof(string));
            dtExcel.Columns.Add("Father_Name", typeof(string));
            dtExcel.Columns.Add("Mother_Name", typeof(string));
            dtExcel.Columns.Add("DateOfBirth", typeof(string));
            dtExcel.Columns.Add("DateOfJoining", typeof(string));
            dtExcel.Columns.Add("Desigation", typeof(string));
            dtExcel.Columns.Add("Department", typeof(string));
            dtExcel.Columns.Add("Depatment_Location", typeof(string));
            dtExcel.Columns.Add("Vendor", typeof(string));
            dtExcel.Columns.Add("Address", typeof(string));
            dtExcel.Columns.Add("District", typeof(string));
            dtExcel.Columns.Add("Block", typeof(string));
            dtExcel.Columns.Add("Panchayat_Name", typeof(string));
            dtExcel.Columns.Add("Postal_Code", typeof(string));
            dtExcel.Columns.Add("PAN", typeof(string));
            dtExcel.Columns.Add("UAN", typeof(string));
            dtExcel.Columns.Add("AADHAR", typeof(string));
            dtExcel.Columns.Add("Bank", typeof(string));
            dtExcel.Columns.Add("IFSC", typeof(string));
            dtExcel.Columns.Add("Account_No", typeof(string));
            dtExcel.Columns.Add("Basic", typeof(string));
            dtExcel.Columns.Add("HRA", typeof(string));
            dtExcel.Columns.Add("Others", typeof(string));
            dtExcel.Columns.Add("IsPf", typeof(string));
            dtExcel.Columns.Add("IsEsi", typeof(string));
            dtExcel.Columns.Add("Charge", typeof(string));
            dtExcel.Columns.Add("TA", typeof(string));
            dtExcel.Columns.Add("SP", typeof(string));
            dtExcel.Columns.Add("GrossWages", typeof(string));
            dtExcel.Columns.Add("ESINO", typeof(string));
            int insertCount = 0;
            DataColumnCollection columns = importdt.Columns;
            DataTable dt = importdt;
            DataRow dr;
            for (int i = 0; i < dt.Rows.Count; i++)
            {

                if (dt.Rows[i][0].ToString() != " ")
                {
                    dr = dtExcel.NewRow();
                    //dr["EmpCode"] = dt.Rows[i]["EmpCode"].ToString();
                    //dr["eName"] = dt.Rows[i]["eName"].ToString();

                    if (dt.Rows[i]["Code"] != System.DBNull.Value)
                    {
                        dr["Code"] = dt.Rows[i]["Code"].ToString();

                    }
                    else
                    {
                        dr["Code"] = "0";
                    }

                    if (dt.Rows[i]["Empname"] != System.DBNull.Value)
                    {
                        dr["Empname"] = dt.Rows[i]["Empname"].ToString();
                    }
                    else
                    {
                        dr["Empname"] = "0";
                    }

                    if (dt.Rows[i]["Father's Name"] != System.DBNull.Value)
                    {
                        dr["Father_Name"] = dt.Rows[i]["Father's Name"].ToString();
                    }
                    else
                    {
                        dr["Father_Name"] = "0";
                    }

                    if (dt.Rows[i]["Mother's Name"] != System.DBNull.Value)
                    {
                        dr["Mother_Name"] = dt.Rows[i]["Mother's Name"].ToString();


                    }
                    else
                    {

                        dr["Mother_Name"] = "0";
                    }


                    if (dt.Rows[i]["Date Of Birth"] != System.DBNull.Value)
                    {
                        dr["DateOfBirth"] = dt.Rows[i]["Date Of Birth"].ToString();
                    }
                    else
                    {
                        dr["DateOfBirth"] = "01/01/1900";
                    }

                    if (dt.Rows[i]["Date of Joining"] != System.DBNull.Value)
                    {
                        dr["DateOfJoining"] = dt.Rows[i]["Date of Joining"].ToString();
                    }
                    else
                    {
                        dr["DateOfJoining"] = "01/01/1900";
                    }

                    if (dt.Rows[i]["Desigation"] != System.DBNull.Value)
                    {
                        dr["Desigation"] = dt.Rows[i]["Desigation"].ToString();
                    }
                    else
                    {
                        dr["Desigation"] = "0";
                    }

                    if (dt.Rows[i]["Department"] != System.DBNull.Value)
                    {
                        dr["Department"] = dt.Rows[i]["Department"].ToString();
                    }
                    else
                    {
                        dr["Department"] = "0";
                    }


                    if (dt.Rows[i]["Depatment Location"] != System.DBNull.Value)
                    {
                        dr["Depatment_Location"] = dt.Rows[i]["Depatment Location"].ToString();
                    }
                    else
                    {
                        dr["Depatment_Location"] = "0";
                    }

                    if (dt.Rows[i]["Vendor"] != System.DBNull.Value)
                    {
                        dr["Vendor"] = dt.Rows[i]["Vendor"].ToString();
                    }
                    else
                    {
                        dr["Vendor"] = "0";
                    }

                    if (dt.Rows[i]["Address"] != System.DBNull.Value)
                    {
                        dr["Address"] = dt.Rows[i]["Address"].ToString();
                    }
                    else
                    {
                        dr["Address"] = "0";
                    }

                    if (dt.Rows[i]["District"] != System.DBNull.Value)
                    {
                        dr["District"] = dt.Rows[i]["District"].ToString();
                    }
                    else
                    {
                        dr["District"] = "0";
                    }

                    if (dt.Rows[i]["Block"] != System.DBNull.Value)
                    {
                        dr["Block"] = dt.Rows[i]["Block"].ToString();
                    }
                    else
                    {
                        dr["Block"] = "0";
                    }

                    if (dt.Rows[i]["Panchayat Name"] != System.DBNull.Value)
                    {
                        dr["Panchayat_Name"] = dt.Rows[i]["Panchayat Name"].ToString();
                    }
                    else
                    {
                        dr["Panchayat_Name"] = "0";

                    }

                    if (dt.Rows[i]["Postal Code"] != System.DBNull.Value)
                    {
                        dr["Postal_Code"] = dt.Rows[i]["Postal Code"].ToString();
                    }
                    else
                    {
                        dr["Postal_Code"] = "0";
                    }

                    if (dt.Rows[i]["PAN"] != System.DBNull.Value)
                    {
                        dr["PAN"] = dt.Rows[i]["PAN"].ToString();
                    }
                    else
                    {
                        dr["PAN"] = "0";
                    }

                    if (dt.Rows[i]["UAN"] != System.DBNull.Value)
                    {
                        dr["UAN"] = dt.Rows[i]["UAN"].ToString();
                    }
                    else
                    {
                        dr["UAN"] = "0";
                    }

                    if (dt.Rows[i]["AADHAR"] != System.DBNull.Value)
                    {
                        dr["AADHAR"] = dt.Rows[i]["AADHAR"].ToString();
                    }
                    else
                    {
                        dr["AADHAR"] = "0";
                    }


                    if (dt.Rows[i]["Bank"] != System.DBNull.Value)
                    {
                        dr["Bank"] = dt.Rows[i]["Bank"].ToString();
                    }
                    else
                    {
                        dr["Bank"] = "0";
                    }

                    if (dt.Rows[i]["IFSC"] != System.DBNull.Value)
                    {
                        dr["IFSC"] = dt.Rows[i]["IFSC"].ToString();
                    }
                    else
                    {
                        dr["IFSC"] = "0";
                    }
                    if (dt.Rows[i]["Account No"] != System.DBNull.Value)
                    {
                        dr["Account_No"] = dt.Rows[i]["Account No"].ToString();
                    }
                    else
                    {
                        dr["Account_No"] = "0";
                    }

                    if (dt.Rows[i]["Basic"] != System.DBNull.Value)
                    {
                        dr["Basic"] = dt.Rows[i]["Basic"].ToString();
                    }
                    else
                    {
                        dr["Basic"] = "0";
                    }

                    if (dt.Rows[i]["HRA"] != System.DBNull.Value)
                    {
                        dr["HRA"] = dt.Rows[i]["HRA"].ToString();
                    }
                    else
                    {
                        dr["HRA"] = "0";
                    }

                    if (dt.Rows[i]["Others"] != System.DBNull.Value)
                    {
                        dr["Others"] = dt.Rows[i]["Others"].ToString();
                    }
                    else
                    {
                        dr["Others"] = "0";
                    }

                    if (dt.Rows[i]["IsPf"] != System.DBNull.Value)
                    {
                        dr["IsPf"] = dt.Rows[i]["IsPf"].ToString();
                    }
                    else
                    {
                        dr["IsPf"] = "0";
                    }

                    if (dt.Rows[i]["IsEsi"] != System.DBNull.Value)
                    {
                        dr["IsEsi"] = dt.Rows[i]["IsEsi"].ToString();
                    }
                    else
                    {
                        dr["IsEsi"] = "0";
                    }

                    if (dt.Rows[i]["Charge(%)"] != System.DBNull.Value)
                    {
                        dr["Charge"] = dt.Rows[i]["Charge(%)"].ToString();
                    }
                    else
                    {
                        dr["Charge"] = "0";
                    }
                    if (dt.Rows[i]["TA"] != System.DBNull.Value)
                    {
                        dr["TA"] = dt.Rows[i]["TA"].ToString();
                    }
                    else
                    {
                        dr["TA"] = "0";
                    }
                    if (dt.Rows[i]["SP"] != System.DBNull.Value)
                    {
                        dr["SP"] = dt.Rows[i]["SP"].ToString();
                    }
                    else
                    {
                        dr["SP"] = "0";
                    }
                    if (dt.Rows[i]["GrossWages"] != System.DBNull.Value)
                    {
                        dr["GrossWages"] = dt.Rows[i]["GrossWages"].ToString();
                    }
                    else
                    {
                        dr["GrossWages"] = "0";
                    }
                    if (dt.Rows[i]["ESINO"] != System.DBNull.Value)
                    {
                        dr["ESINO"] = dt.Rows[i]["ESINO"].ToString();
                    }
                    else
                    {
                        dr["ESINO"] = "0";
                    }

                    dtExcel.Rows.Add(dr);

                }
                //i++ ;
            }
            return dtExcel;
        }


        public JsonResult _InsertDetail()
        {

            string mes = string.Empty;
            DataTable dataTable = new DataTable();
            try
            {
                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(Request.Form["Table"].ToString());
                DataTable table = new DataTable();
                table.Columns.Add("EmpCode", typeof(string));
                table.Columns.Add("EmpName", typeof(string));
                table.Columns.Add("Father_Name", typeof(string));
                table.Columns.Add("Mother_Name", typeof(string));
                table.Columns.Add("DateOfBirth", typeof(DateTime));
                table.Columns.Add("DateOfJoining", typeof(DateTime));
                table.Columns.Add("Desigation", typeof(string));
                table.Columns.Add("Department", typeof(string));
                table.Columns.Add("Department_Location", typeof(string));
                table.Columns.Add("Vendor", typeof(string));
                table.Columns.Add("fVAddress", typeof(string));
                table.Columns.Add("District", typeof(string));
                table.Columns.Add("Town", typeof(string));
                table.Columns.Add("Panchayat_Name", typeof(string));
                table.Columns.Add("Postal_Code", typeof(string));
                table.Columns.Add("PAN", typeof(string));
                table.Columns.Add("UAN", typeof(string));
                table.Columns.Add("AADHAR", typeof(string));
                table.Columns.Add("Bank", typeof(string));
                table.Columns.Add("IFSC", typeof(string));
                table.Columns.Add("Account_No", typeof(string));
                table.Columns.Add("fnBasic", typeof(string));
                table.Columns.Add("HRA", typeof(string));
                table.Columns.Add("Others", typeof(string));
                table.Columns.Add("IsPf", typeof(string));
                table.Columns.Add("IsEsi", typeof(string));
                table.Columns.Add("Charge", typeof(string));
                table.Columns.Add("TA", typeof(string));
                table.Columns.Add("SP", typeof(string));
                table.Columns.Add("GrossWages", typeof(string));
                table.Columns.Add("ESINO", typeof(string));

                foreach (var Item1 in jsonData)
                {
                    DataRow dr1 = table.NewRow();

                    dr1["EmpCode"] = Item1.EmpCode;
                    dr1["EmpName"] = Item1.EmpName;
                    dr1["Father_Name"] = Item1.Father_Name;
                    dr1["Mother_Name"] = Item1.Mother_Name;
                    dr1["DateOfBirth"] = Convert.ToDateTime((Item1.DateOfBirth).ToString()).ToString("dd-MMM-yyyy");
                    dr1["DateOfJoining"] = Convert.ToDateTime((Item1.DateOfJoining).ToString()).ToString("dd-MMM-yyyy");
                    dr1["Desigation"] = Item1.Desigation;
                    dr1["Department"] = Item1.Department;
                    dr1["Department_Location"] = Item1.Depatment_Location;
                    dr1["Vendor"] = Item1.Vendor;
                    dr1["fVAddress"] = Item1.Address;
                    dr1["District"] = Item1.District;
                    dr1["Town"] = Item1.Block;
                    dr1["Panchayat_Name"] = Item1.Panchayat_Name;
                    dr1["Postal_Code"] = Item1.Postal_Code;
                    dr1["PAN"] = Item1.PAN;
                    dr1["UAN"] = Item1.UAN;
                    dr1["AADHAR"] = Item1.AADHAR;
                    dr1["Bank"] = Item1.Bank;
                    dr1["IFSC"] = Item1.IFSC;
                    dr1["Account_No"] = Item1.Account_No;
                    dr1["fnBasic"] = Item1.Basic;
                    dr1["HRA"] = Item1.HRA;
                    dr1["Others"] = Item1.Others;
                    dr1["IsPf"] = Item1.IsPf;
                    dr1["IsEsi"] = Item1.IsEsi;
                    dr1["Charge"] = Item1.Charge;
                    dr1["TA"] = Item1.TA;
                    dr1["SP"] = Item1.SP;
                    dr1["GrossWages"] = Item1.GrossWages;
                    dr1["ESINO"] = Item1.ESINO;

                    table.Rows.Add(dr1);
                }


                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "stpHPSEDCImportMaster";
                SqlParameter parameter = new SqlParameter();
                com.Parameters.AddWithValue("@loginId", Session["EmpId"].ToString());
                com.Parameters.AddWithValue("@SessionID", Session["SessionId"].ToString());
                parameter.ParameterName = "@EmpDetails";
                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table;
                com.Parameters.Add(parameter);
                com.CommandTimeout = 0;
                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();
                com.ExecuteNonQuery();
                SqlDataAdapter adapter = new SqlDataAdapter();
                DataTable dt = new DataTable();
                adapter.SelectCommand = com;
                adapter.Fill(dataTable);
                conObj.con.Close();

            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("List", "Error: " + ex.Message);
            }
            var json = JsonConvert.SerializeObject(dataTable);
            return Json(json);
        }

        public ActionResult FreezeAttendances()
        {
            SortedList list = new SortedList();
            list.Add("@PageNo", 1);
            list.Add("@PageSize", 1000);
            DataTable dt = comfun.fillDataTable("Department_Select", "", list);
            return View(dt);
        }

        public JsonResult _FreezeAttendanceLists()
        {
            // comfun.saveformname("_HelpdeskGroupMappingList", "/admin/_HelpdeskGroupMappingList", "Master/Help Desk/Mapping&nbsp;&nbsp;", "HelpDesk Mapping List", "N", "HelpDeskMapping", "HelpDeskMapping", "List", 4);
            SortedList list = new SortedList();
            string fromDate = Request.Form["StartDate"].ToString();
            int startYear = Convert.ToDateTime(fromDate).Year;
            int startMonth = Convert.ToDateTime(fromDate).Month;

            DateTime startDate = new DateTime(startYear, startMonth, 1);
            DateTime endDate = startDate.AddMonths(1).AddDays(-1);
            list.Add("@StartDate", startDate.ToString("dd/MMM/yyyy"));
            list.Add("@EndDate", endDate.ToString("dd/MMM/yyyy"));
            list.Add("@BranchID", Request.Form["BranchId"].ToString());
            list.Add("@DepartmentID", Request.Form["DepartmentId"].ToString());
            list.Add("@CircleId", Request.Form["CircleId"].ToString());
            list.Add("@CategoryId", Request.Form["CategoryId"].ToString());
            list.Add("@MainCategoryId", Request.Form["MainCategoryId"].ToString());
            DataTable dt = comfun.fillDataTable("stpEmployeeAttendance_MonthlyLockList", "", list);
            //return PartialView("_HelpdeskGroupMappingList", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }


        public ActionResult _FreezeEmployees()
        {
            SortedList list = new SortedList();
            //string fromDate = Request.Form["FromDate"].ToString();
            string AsonDate = Convert.ToDateTime(Request.Form["FromDate"].ToString()).ToString("dd/MMM/yyyy");
            int startYear = Convert.ToDateTime(AsonDate).Year;
            int startMonth = Convert.ToDateTime(AsonDate).Month;
            string Type = Request.Form["Type"].ToString();
            DateTime startDate = new DateTime(startYear, startMonth, 1);
            list.Add("@StartDate", startDate.ToString("dd/MMM/yyyy"));
            //DateTime endDate = startDate.AddMonths(1).AddDays(-1);

            list.Add("@EndDate", AsonDate /*endDate.ToString("dd/MMM/yyyy")*/);
            list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
            list.Add("@BranchId", Request.Form["BranchId"].ToString());
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@CategoryId", Request.Form["CategoryId"].ToString());
            list.Add("@MainCategoryId", Request.Form["MainCategoryId"].ToString());
            list.Add("@CircleId", Request.Form["CircleId"].ToString());
            list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
            list.Add("@LoginID", Session["EmpId"].ToString());
            DataTable dt = new DataTable();
            if (Type == "Freeze")
            {
                dt = comfun.fillDataTable("stpEmployeeAttendance_MonthlyLock", "", list);

            }
            if (Type == "UnFreeze")
            {
                dt = comfun.fillDataTable("stpEmployeeAttendance_MonthlyUnlock", "", list);

            }
            return PartialView(dt);
            //var json = JsonConvert.SerializeObject(dt);
            //return Json(json);
        }


        public JsonResult _FreezeSaveUpdateSubmits(string FromDate, string Type, string Emps)
        {

            string mes = string.Empty;
            try
            {

                string fromDate = FromDate;
                int startYear = Convert.ToDateTime(fromDate).Year;
                int startMonth = Convert.ToDateTime(fromDate).Month;

                DateTime startDate = new DateTime(startYear, startMonth, 1);
                //DateTime endDate = startDate.AddMonths(1).AddDays(-1);



                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(Emps);
                //dynamic jsonData1 = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(SaleOrderVendor);
                DataTable table1 = new DataTable();
                table1.Columns.Add("EmpId", typeof(string));
                table1.Columns.Add("PayDays", typeof(string));



                foreach (var Item1 in jsonData)
                {
                    DataRow dr1 = table1.NewRow();

                    dr1["EmpId"] = Item1.EmpId;
                    dr1["PayDays"] = Item1.PayDays;

                    table1.Rows.Add(dr1);
                }

                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "EmployeeAttendance_FreezeNewSubmit";
                SqlParameter parameter = new SqlParameter();
                com.Parameters.AddWithValue("@MonthYearId", startMonth.ToString() + startYear.ToString());
                com.Parameters.AddWithValue("@LockDate", FromDate /*endDate.ToString("dd-MMM-yyyy")*/);
                com.Parameters.AddWithValue("@CreatedDate", comfun.dateISTstr());
                com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());
                com.Parameters.AddWithValue("@Type", Type);
                //  com.Parameters.AddWithValue("@EmpIdexclude", Request.Form["Emps1"].ToString());
                //com.Parameters.AddWithValue("@VenderId", VenderId);
                parameter.ParameterName = "@EmpId";
                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table1;
                com.Parameters.Add(parameter);
                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);
                com.CommandTimeout = 0;

                if (conObj.con.State == ConnectionState.Closed)
                {
                    conObj.con.Open();
                    com.ExecuteNonQuery();
                    conObj.con.Close();
                }
                mes = sp.Value.ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_FreezeSaveUpdateSubmits", "Error: " + ex.Message);
            }



            return Json(mes);


        }






        #region Mukesh
        public ActionResult DepartmentAttendancApprovalPendingCustomer()
        {
            return View();
        }
        public ActionResult _DepartmentAttendanceapprovalPendingCustomer()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@MonthYear", Request.Form["MonthYear"].ToString());
            list.Add("@ContractorId", Session["Customer"]);
            dt = comfun.fillDataTable("stpDepartmentAttedanceForVendor_List", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult _ContrcatorInvoiceListCustomer()
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}

            SortedList list = new SortedList();

            list.Add("@ContractorId", "0");
            list.Add("@CustomerId", Request.Form["CustomerId"].ToString());
            list.Add("@EntityId", Request.Form["EntityId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthYear"].ToString());


            DataTable dt = comfun.fillDataTable("stpAttendanceApproval_InvoiceList", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult PurchaseBillAmount()
        {
            SortedList list = new SortedList();
            list.Add("@CustomerInvoiceP2", Request.Form["BillDetail"].ToString());
            list.Add("@CustomerId", Request.Form["CustomerId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthYear"].ToString());
            DataSet dt = comfun.fillDataSet("stpVendorBillGenerate", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _AttendanceCustomerInvoiceCustomerSubmit(string BillDetail, string BillDesignation, string FromDate, string BillNo,
           string BillDate, string Id, string Remarks, string ReferanceNo, string CustomerID, string Amount, string GSTAmount, string SGSTAmount, string CGSTAmount
           , string HPSEDC, string BillAmount, string RoundOff
          )
        {

            string mes = string.Empty;
            try
            {
                DataTable table = (DataTable)JsonConvert.DeserializeObject(BillDetail, (typeof(DataTable)));

                DataTable table1 = (DataTable)JsonConvert.DeserializeObject(BillDesignation, (typeof(DataTable)));
                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "stptblVendorSaleInvoice_Accept";
                SortedList list = new SortedList();
                DateTime monthYear = Convert.ToDateTime(FromDate);
                com.Parameters.AddWithValue("@MonthDate", FromDate);
                com.Parameters.AddWithValue("@ID", Id);
                com.Parameters.AddWithValue("@Createby", Session["EmpId"].ToString());
                com.Parameters.AddWithValue("@BillDate", BillDate);
                com.Parameters.AddWithValue("@fvBillNo", BillNo);
                com.Parameters.AddWithValue("@RefNo", ReferanceNo);
                com.Parameters.AddWithValue("@Remarks", Remarks);
                com.Parameters.AddWithValue("@CustomerID", CustomerID);
                //com.Parameters.AddWithValue("@VendorID", VendorID);
                com.Parameters.AddWithValue("@SessionID", Session["SessionId"]);

                com.Parameters.AddWithValue("@Amount", Amount);
                com.Parameters.AddWithValue("@GSTAmount", GSTAmount);
                com.Parameters.AddWithValue("@SGSTAmount", SGSTAmount);
                com.Parameters.AddWithValue("@CGSTAmount", CGSTAmount);
                com.Parameters.AddWithValue("@OtherCharges", HPSEDC);
                com.Parameters.AddWithValue("@BillAmount", BillAmount);
                com.Parameters.AddWithValue("@RoundOff", RoundOff);
                com.Parameters.AddWithValue("@createdDate", DateTime.Now.AddMinutes(330).ToString("dd-MMM-yyy HH:mm:ss"));

                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@CustomerInvoiceP2";
                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table;
                com.Parameters.Add(parameter);

                SqlParameter parameter1 = new SqlParameter();
                parameter1.ParameterName = "@CustomerInvoiceP3";
                parameter1.SqlDbType = System.Data.SqlDbType.Structured;
                parameter1.Value = table1;
                com.Parameters.Add(parameter1);
                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);


                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();

                com.ExecuteNonQuery();

                conObj.con.Close();

                mes = sp.Value.ToString();


            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;

            }
            return Json(mes);

        }
        public JsonResult _ContractorBillGenerateCustomer(string BillDetail, string CustomerId, string Id, string FromDate)
        {
            string mes = string.Empty;
            try
            {
                DataTable table = (DataTable)JsonConvert.DeserializeObject(BillDetail, (typeof(DataTable)));
                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "stpContractorBillSalaryCalculation";
                SortedList list = new SortedList();
                DateTime monthYear = Convert.ToDateTime(FromDate);
                com.Parameters.AddWithValue("@MonthDate", FromDate);
                com.Parameters.AddWithValue("@ID", Id);
                com.Parameters.AddWithValue("@CustomerId", CustomerId);
                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@CustomerInvoiceP2";
                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table;
                com.Parameters.Add(parameter);
                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);
                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();

                com.ExecuteNonQuery();

                conObj.con.Close();

                mes = sp.Value.ToString();


            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;


            }
            return Json(mes);

        }

        public ActionResult _ContrcatorApprovedInvoiceListCustomer()
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}

            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@ContractorId", Request.Form["ContractorId"].ToString());
            list.Add("@CustomerId", Request.Form["CustomerId"].ToString());
            list.Add("@StartDate", Request.Form["FromDate"].ToString());
            list.Add("@EndDate", Request.Form["ToDate"].ToString());
            DataTable dt = comfun.fillDataTable("stpContrcatorApproved_InvoiceList", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult PurchaseInvoiceList()
        {
            return View();
        }
        public ActionResult _PurchaseInvoiceList()
        {
            // comfun.saveformname("_CustomerInvoiceList", "/TimeSheet/_CustomerInvoiceList", "Master/Address/CustomerInvoiceList", "Customer Invoice  list", "N", "CustomerInvoiceList", "CustomerInvoiceList", "List", 4);
            SortedList list = new SortedList();
            list.Add("@LoginId", Session["loginID"].ToString());
            list.Add("@MonthYear", Request.Form["MonthYear"].ToString());
            DataTable dt = comfun.fillDataTable("stpVendorHpsedcInvoiceList", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult PurchaseInvoicePrint()
        {

            return Redirect("~/CrystalReports/PIprint.aspx?Id=" + Request.QueryString["Id"].ToString() + "&&" + "Method=" + "VendorInvoice");
        }

        public ActionResult VendorBillPrint()
        {

            return Redirect("~/CrystalReports/PIprint.aspx?Id=" + Request.QueryString["Id"].ToString() + "&&" + "Method=" + "VendorBillDetailPrint");

        }


        public ActionResult DeptRequestAdd()
        {

            //DataSet ds = new DataSet();
            // = comfun.fillDataSet("stpProductdropown_list", null, null);
            return View();
        }
        public ActionResult _DeptRequestList()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();//@SessionID INT=0,
            list.Add("@RequestId", 0);
            list.Add("@RequestBy", Session["EmpId"].ToString());
            list.Add("@UserRole", 6);
            dt = comfun.fillDataTable("HpsedcDeptResourceReq_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _InsertDeptRequest(string RequestDetails, string RequestId, string DepartmentCode, string DeploymentDate, string NoOfYear, string BillingAddress, string TotalResource)
        {

            string mes = string.Empty;
            try
            {

                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(RequestDetails);

                DataTable table = new DataTable();

                table.Columns.Add("ResourceId", typeof(int));
                table.Columns.Add("ResourceType", typeof(int));
                table.Columns.Add("DeploymentLoc", typeof(int));
                table.Columns.Add("BasicAmount", typeof(int));
                table.Columns.Add("NoOfResource", typeof(int));
                table.Columns.Add("MinQualification", typeof(string));
                table.Columns.Add("MinAge", typeof(int));

                foreach (var Item in jsonData)
                {
                    DataRow dr = table.NewRow();
                    dr["ResourceId"] = Item.ResourceId;
                    dr["ResourceType"] = Item.ResourceTypeId;
                    dr["DeploymentLoc"] = Item.DeploymentLocId;
                    dr["BasicAmount"] = Item.BasicAmount;
                    dr["NoOfResource"] = Item.NoOfResource;
                    dr["MinQualification"] = Item.MinQualification;
                    dr["MinAge"] = Item.MinAge;
                    table.Rows.Add(dr);
                }
                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "HpsedcDeptResourceReq_AcceptUpdate";
                SqlParameter parameter = new SqlParameter();
                com.Parameters.AddWithValue("@RequestId", RequestId);
                com.Parameters.AddWithValue("@DepartmentCode", DepartmentCode);
                com.Parameters.AddWithValue("@DeploymentDate", DeploymentDate);
                com.Parameters.AddWithValue("@NoOfYear", NoOfYear);
                com.Parameters.AddWithValue("@BillingAddress", BillingAddress);
                com.Parameters.AddWithValue("@RequestBy", Session["EmpId"]);
                com.Parameters.AddWithValue("@TotalResource", TotalResource);
                com.Parameters.AddWithValue("@IsActive", "Y");
                SqlParameter parameter1 = new SqlParameter();
                parameter1.ParameterName = "@tblDeptResource2";
                parameter1.SqlDbType = System.Data.SqlDbType.Structured;
                parameter1.Value = table;
                com.Parameters.Add(parameter1);
                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);
                com.CommandTimeout = 0;

                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();
                com.ExecuteNonQuery();
                conObj.con.Close();

                mes = "Record saved successfully";
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("POList", "Error: " + ex.Message);
            }
            return Json(mes);
        }
        public ActionResult _purchaseBillEdit()
        {
            //comfun.saveformname("_purchaseBillList", "/pm/_purchaseBillList", "Voucher/PurchaseBill", "Purchase Bill List", "N", "PurchaseBill", "PurchaseBill", "List", 4);

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@RequestId", Request.Form["Id"].ToString());
            list.Add("@RequestBy", Session["EmpId"].ToString());
            DataSet ds = comfun.fillDataSet("HpsedcDeptResourceReq_Edit", "", list);
            var json = JsonConvert.SerializeObject(ds);
            return Json(json);
        }

        #endregion

        #region Vishu (Agency New)
        public ActionResult AgencyDetail()
        {
            //DataTable dt = comfun.fillDataTable("stpCustomerTypeddl", "", null);
            return View();
        }

        public JsonResult _AgencyDetailSubmit()
        {

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {

                string TripartiteFile = "";
                string _ExtTripartiteFile = "";
                string _comPathTripartiteFile = "";
                string EmpanelmentFile = "";
                string _ExtEmpanelmentFile = "";
                string _comPathEmpanelmentFile = "";

                if (Request.Files.Count > 0)
                {
                    var files = Request.Files;

                    for (int i = 0; i < files.Count; i++)
                    {
                        HttpPostedFileBase file = files[i];
                        // Checking for Internet Explorer
                        if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                        {
                            string[] testfiles = file.FileName.Split(new char[] { '\\' });
                            EmpanelmentFile = testfiles[testfiles.Length - 1];

                        }
                        else
                        {
                            if (i == 0)
                            {
                                TripartiteFile = file.FileName;
                                var _ext = Path.GetExtension(TripartiteFile);
                                _ExtTripartiteFile = Path.GetFileNameWithoutExtension(TripartiteFile);
                                // Get the complete folder path and store the file inside it.  
                                _ExtTripartiteFile = _ExtTripartiteFile + _ext;
                                string filePath = Path.Combine(Server.MapPath("/ImageHPSEDC/") + _ExtTripartiteFile);

                                TripartiteFile = filePath;
                                _comPathTripartiteFile = filePath;
                                file.SaveAs(_comPathTripartiteFile);
                            }
                            if (i == 1)
                            {
                                EmpanelmentFile = file.FileName;
                                var _ext = Path.GetExtension(EmpanelmentFile);
                                _ExtEmpanelmentFile = Path.GetFileNameWithoutExtension(EmpanelmentFile);
                                // Get the complete folder path and store the file inside it.
                                _ExtEmpanelmentFile = _ExtEmpanelmentFile + _ext;
                                string filePath = Path.Combine(Server.MapPath("/ImageHPSEDC/") + _ExtEmpanelmentFile);

                                EmpanelmentFile = filePath;
                                _comPathEmpanelmentFile = filePath;
                                file.SaveAs(_comPathEmpanelmentFile);
                            }
                        }

                    }

                }

                list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
                list.Add("@AgencyName", Request.Form["AgencyName"].ToString());
                list.Add("@Category", Request.Form["Category"].ToString());
                list.Add("@ProprietorMDName", Request.Form["ProprietorMDName"].ToString());
                list.Add("@RegistrationNo", Request.Form["RegistrationNo"].ToString());
                list.Add("@CompanyType", Request.Form["CompanyType"].ToString());
                list.Add("@PanNumber", Request.Form["PanNumber"].ToString());
                list.Add("@EPF", Request.Form["EPF"].ToString());
                list.Add("@ESIC", Request.Form["ESIC"].ToString());
                list.Add("@Gst", Request.Form["Gst"].ToString());
                list.Add("@LabourLicense", Request.Form["LabourLicense"].ToString());
                list.Add("@LicenseNo", Request.Form["LicenseNo"].ToString());
                list.Add("@AEmpanelmentDate", Request.Form["AEmpanelmentDate"].ToString());

                list.Add("@AccountHolderName", Request.Form["AccountHolderName"].ToString());
                list.Add("@AccountNumber", Request.Form["AccountNumber"].ToString());
                list.Add("@IfscCode", Request.Form["IfscCode"].ToString());
                list.Add("@BranchDetail", Request.Form["BranchDetail"].ToString());

                list.Add("@VendorHqAddress", Request.Form["VendorHqAddress"].ToString());
                list.Add("@VendorHpAddress", Request.Form["VendorHpAddress"].ToString());
                list.Add("@OfficialEmail", Request.Form["OfficialEmail"].ToString());
                list.Add("@OfficialPhone", Request.Form["OfficialPhone"].ToString());
                list.Add("@PhoneNumber1", Request.Form["PhoneNumber1"].ToString());
                list.Add("@PhoneNumber2", Request.Form["PhoneNumber2"].ToString());
                list.Add("@EmpanelmentFeesAmount", Request.Form["EmpanelmentFeesAmount"].ToString());
                list.Add("@EmpanelmentFeesStatus", Request.Form["EmpanelmentFeesStatus"].ToString());
                list.Add("@EMDAmount", Request.Form["EMDAmount"].ToString());
                list.Add("@EMDStatus", Request.Form["EMDStatus"].ToString());
                list.Add("@PBGAmount", Request.Form["PBGAmount"].ToString());
                list.Add("@PBGStatus", Request.Form["PBGStatus"].ToString());
                list.Add("@TripartiteFileDocument", _ExtTripartiteFile);
                list.Add("@EmpanelmentFileDocument", _ExtEmpanelmentFile);

                list.Add("@BiPartyExpireDate", Request.Form["BiPartyExpireDate"].ToString());

                list.Add("@EmpanelmentRefNo", Request.Form["EmpanelmentRefNo"].ToString());
                list.Add("@EmpanelmentDate", Request.Form["EmpanelmentDate"].ToString());
                list.Add("@EmpanelmentMethod", Request.Form["EmpanelmentMethod"].ToString());
                list.Add("@EMDRefNo", Request.Form["EMDRefNo"].ToString());

                list.Add("@EMDDate", Request.Form["EMDDate"].ToString());
                list.Add("@EMDMethod", Request.Form["EMDMethod"].ToString());
                list.Add("@PBGRefNo", Request.Form["PBGRefNo"].ToString());
                list.Add("@PBGDate", Request.Form["PBGDate"].ToString());
                list.Add("@PBGMethod", Request.Form["PBGMethod"].ToString());
                mes = comfun.executeNonQueryWMessage("stpAgencyDetail_Accept", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
            }
            return Json(mes);
        }
        public ActionResult _AgencyDetailList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            dt = comfun.fillDataTable("StpAgencyDetail_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _AgencyDetailEdit()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            dt = comfun.fillDataTable("StpAgencyDetail_Edit", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        #endregion

        #region New Section
        public ActionResult EmployeePool()
        {
            return View();
        }

        public ActionResult EmployeeAdd()
        {
            return View();

        }
        public ActionResult _EmpHPSEDCAVOddl()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();

            list.Add("@DeptId", Session["DeptId"]?.ToString() ?? "0");
            dt = comfun.fillDataTable("HpsedcAVO_dll", "", list);

            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public ActionResult _Customerddl1()

        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@CustomerId", Session["Customer"]);
            list.Add("@EntityId", Request.Form["EntityId"].ToString());
            dt = comfun.fillDataTable("stpCustomerSelect", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public ActionResult _AvoEmployeeUpdate1()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            //list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@DeptHQAddressText", Request.Form["DeptHQAddressText"].ToString());
            list.Add("@BillingOfficeAddressText", Request.Form["BillingOfficeAddressText"].ToString());
            list.Add("@AvoId", Request.Form["AvoId"].ToString());

            dt = comfun.fillDataTable("HpsedcNewEmpAvo_Update", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }


        public ActionResult _AvoEmployeeUpdate()
        {

            string mes = "";
            try
            {

                SortedList list1 = new SortedList();
                DataTable dt = new DataTable();

                SortedList list = new SortedList();
                list1.Add("@EmpId", Request.Form["EmpId"].ToString());
                //list.Add("@CreatedBy", Session["EmpId"].ToString());
                list1.Add("@DeptHQAddressText", Request.Form["DeptHQAddressText"].ToString());
                list1.Add("@BillingOfficeAddressText", Request.Form["BillingOfficeAddressText"].ToString());
                list1.Add("@AvoId", Request.Form["AvoId"].ToString());


                mes = comfun.executeNonQueryWMessage("HpsedcNewEmpAvo_Update", "", list1).ToString();

            }
            catch (Exception ex) { mes = ex.Message; }

            return Json(mes);
        }



        public ActionResult _EmployeeAddList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            //list.Add("@CustomerId", Session["Customer"]);
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@DeptId", Session["DeptId"]?.ToString() ?? "0");
            dt = comfun.fillDataTable("HpsedcNewEmp_List", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _EmployeeAddEdit()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId", Session["Customer"]);
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            // list.Add("@CreatedBy", Session["EmpId"].ToString());
            dt = comfun.fillDataTable("HpsedcNewEmp_Edit", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public JsonResult _EmployeeAdd5()
        {
            string mes = string.Empty;
            try
            {
                string Applicant_image = "";
                string ApplicantImage = "";
                string _comPathApplicant = "";
                string RationCard_image = "";
                string RationCardImage = "";
                string _comPathRationCard = "";

                var files = Request.Files;
                // if (Request.Files.Count > 0)
                //{

                for (int i = 0; i < files.Count; i++)
                {
                    HttpPostedFileBase file = files[i];

                    // Checking for Internet Explorer  
                    if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                    {
                        string[] testfiles = file.FileName.Split(new char[] { '\\' });
                        RationCardImage = testfiles[testfiles.Length - 1];

                    }
                    else
                    {
                        if (i == 0)
                        {
                            ApplicantImage = file.FileName;
                            var _ext = Path.GetExtension(ApplicantImage);
                            Applicant_image = Path.GetFileNameWithoutExtension(ApplicantImage);
                            // Get the complete folder path and store the file inside it.  
                            Applicant_image = Applicant_image + _ext;
                            string filePath = Path.Combine(Server.MapPath("/EmpPhoto/") + Applicant_image);
                            ApplicantImage = filePath;
                            _comPathApplicant = filePath;
                            file.SaveAs(_comPathApplicant);
                        }
                        if (i == 1)
                        {
                            RationCardImage = file.FileName;
                            var _ext = Path.GetExtension(RationCardImage);
                            RationCard_image = Path.GetFileNameWithoutExtension(RationCardImage);
                            // Get the complete folder path and store the file inside it.  
                            RationCard_image = RationCard_image + _ext;
                            string filePath = Path.Combine(Server.MapPath("/EmpDocuments/") + RationCard_image);
                            RationCardImage = filePath;
                            _comPathRationCard = filePath;
                            file.SaveAs(_comPathRationCard);
                        }

                    }
                }
                //}
                SortedList list = new SortedList();
                list.Add("@EmpId", Request.Form["EmpId"].ToString());
                list.Add("@UploadPhoto", Applicant_image);
                list.Add("@UploadQualification", RationCard_image);
                list.Add("@CreatedBy", Session["EmpId"].ToString());

                mes = comfun.executeNonQueryWMessage("HpsedcNewEmp5_AcceptUpdate", "", list).ToString();


            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage(" ", "Error: " + ex.Message);
            }
            return Json(mes);
        }
        public JsonResult _EmployeeAdd4()
        {
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@EmpId", Request.Form["EmpId"].ToString());
                list.Add("@Designation", Request.Form["Designation"].ToString());
                list.Add("@UanNo", Request.Form["UanNo"].ToString());
                list.Add("@ESICNo", Request.Form["ESICNo"].ToString());
                list.Add("@DeptName", Request.Form["DeptName"].ToString());
                list.Add("@DeptHQAddress", Request.Form["DeptHQAddress"].ToString());
                list.Add("@BillingOfficeAddress", Request.Form["BillingOfficeAddress"].ToString());
                list.Add("@DeptHQAddressText", Request.Form["DeptHQAddressText"].ToString());
                list.Add("@BillingOfficeAddressText", Request.Form["BillingOfficeAddressText"].ToString());

                list.Add("@CreatedBy", Session["EmpId"].ToString());


                mes = comfun.executeNonQueryWMessage("HpsedcNewEmp4_AcceptUpdate", "", list).ToString();


            }

            catch (Exception ex)
            {

                mes = "Error:" + ex.Message;
            }

            return Json(mes);

        }
        public JsonResult _EmployeeAdd3()
        {
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {

                list.Add("@EmpId", Request.Form["EmpId"].ToString());
                list.Add("@BasicAmount", Request.Form["BasicAmount"].ToString());
                list.Add("@IsDailyWages", Request.Form["IsDailyWages"].ToString());
                list.Add("@IsPartTimer", Request.Form["IsPartTimer"].ToString());
                list.Add("@NoOfHours", Request.Form["NoOfHours"].ToString());
                list.Add("@IsAdminChargeFixed", Request.Form["IsAdminChargeFixed"].ToString());
                list.Add("@AdminCharge", Request.Form["AdminCharge"].ToString());
                list.Add("@SpecialAllowance", Request.Form["SpecialAllowance"].ToString());
                list.Add("@liblyAllowance", Request.Form["liblyAllowance"].ToString());
                list.Add("@Medical", Request.Form["Medical"].ToString());
                list.Add("@HRA", Request.Form["HRA"].ToString());
                list.Add("@Tribalallwoance", Request.Form["Tribalallwoance"].ToString());
                list.Add("@TribalAreaallwoance", Request.Form["TribalAreaallwoance"].ToString());
                list.Add("@ReliverCharges", Request.Form["ReliverCharges"].ToString());
                list.Add("@DressCharge", Request.Form["DressCharge"].ToString());
                list.Add("@IsEsi", Request.Form["IsEsi"].ToString());
                list.Add("@IsPf", Request.Form["IsPf"].ToString());
                list.Add("@IsSpecialEsi", Request.Form["IsSpecialEsic"].ToString());
                list.Add("@GrossAmount", Request.Form["GrossAmount"].ToString());
                list.Add("@Other", Request.Form["Other"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());


                mes = comfun.executeNonQueryWMessage("HpsedcNewEmp3_AcceptUpdate", "", list).ToString();

            }

            catch (Exception ex)
            {

                mes = "Error:" + ex.Message;
            }

            return Json(mes);

        }
        public JsonResult _EmployeeAdd2()
        {
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {

                list.Add("@EmpId", Request.Form["EmpId"].ToString());
                //list.Add("@IsPart1", Request.Form["IsPart1"].ToString());
                list.Add("@BankName", Request.Form["Bank"].ToString());
                list.Add("@BranchName", Request.Form["Branch"].ToString());
                list.Add("@IfscCode", Request.Form["IFSC"].ToString());
                list.Add("@AcNumber", Request.Form["Acno"].ToString());
                list.Add("@PanNo", Request.Form["PAN"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());


                mes = comfun.executeNonQueryWMessage("HpsedcNewEmp2_AcceptUpdate", "", list).ToString();

            }

            catch (Exception ex)
            {

                mes = "Error:" + ex.Message;
            }

            return Json(mes);

        }

        public JsonResult _EmployeeAdd()
        {
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {

                list.Add("@EmpId", Request.Form["EmpId"].ToString());
                list.Add("@EmpName", Request.Form["EmpName"].ToString());
                list.Add("@Fathername", Request.Form["fname"].ToString());
                list.Add("@Mothername", Request.Form["mname"].ToString());
                list.Add("@Gender", Request.Form["gender"].ToString());
                list.Add("@Dob", Request.Form["Dob"].ToString());
                list.Add("@Education", Request.Form["education"].ToString());
                list.Add("@Employeeaddress", Request.Form["employeeaddress"].ToString());
                list.Add("@Panchayat", Request.Form["panchayat"].ToString());
                list.Add("@state", Request.Form["state"].ToString());
                list.Add("@Distirct", Request.Form["dist"].ToString());
                list.Add("@Block", Request.Form["City"].ToString());
                list.Add("@Pin", Request.Form["pin"].ToString());
                list.Add("@Mobileno", Request.Form["mobileno"].ToString());
                list.Add("@Email", Request.Form["email"].ToString());
                list.Add("@AadharNo", Request.Form["aadhar"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());


                mes = comfun.executeNonQueryWMessage("HpsedcNewEmp_AcceptUpdate", "", list).ToString();

            }

            catch (Exception ex)
            {

                mes = "Error:" + ex.Message;
            }
            string[] EmpList = mes.Split(',');
            mes = EmpList[1].ToString();
            return Json(mes);

            //  return Json("sdfds");
        }
        public ActionResult _CustomerEntityddlNew()
        {
            DataTable dt = comfun.fillDataTable("HpsedcDepartmentddl", "", null);
            return PartialView(dt);
        }

        public ActionResult _Bankddl()
        {
            SortedList list = new SortedList();
            list.Add("@Id", 0);
            DataTable dt = comfun.fillDataTable("HpsedcBankDdl_List", "", list);
            var json1 = JsonConvert.SerializeObject(dt);

            return Json(json1);
        }


        public ActionResult _Designationdll()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();

            dt = comfun.fillDataTable("stpDesignationdll", "", null);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }


        public ActionResult _Customerddl2()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@CustomerId", Session["Customer"]);
            list.Add("@EntityId", Request.Form["EntityId"].ToString());
            dt = comfun.fillDataTable("stpCustomerSelect", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }


        #region DeptOfficeAdd
        public ActionResult DeptOfficeAdd()
        {
            return View();
        }
        public JsonResult _DeptOfficeAdd()
        {
            string mes = string.Empty;
            try
            {
                string Applicant_image = "";
                string ApplicantImage = "";
                string _comPathApplicant = "";
                string RationCard_image = "";
                string RationCardImage = "";
                string _comPathRationCard = "";

                var files = Request.Files;
                // if (Request.Files.Count > 0)
                //{

                for (int i = 0; i < files.Count; i++)
                {
                    HttpPostedFileBase file = files[i];

                    // Checking for Internet Explorer  
                    if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                    {
                        string[] testfiles = file.FileName.Split(new char[] { '\\' });
                        RationCardImage = testfiles[testfiles.Length - 1];

                    }
                    else
                    {
                        if (i == 0)
                        {
                            ApplicantImage = file.FileName;
                            var _ext = Path.GetExtension(ApplicantImage);
                            Applicant_image = Path.GetFileNameWithoutExtension(ApplicantImage);
                            // Get the complete folder path and store the file inside it.  
                            Applicant_image = Applicant_image + _ext;
                            string filePath = Path.Combine(Server.MapPath("/DeptSignature/") + Applicant_image);
                            ApplicantImage = filePath;
                            _comPathApplicant = filePath;
                            file.SaveAs(_comPathApplicant);
                        }
                        if (i == 1)
                        {
                            RationCardImage = file.FileName;
                            var _ext = Path.GetExtension(RationCardImage);
                            RationCard_image = Path.GetFileNameWithoutExtension(RationCardImage);
                            // Get the complete folder path and store the file inside it.  
                            RationCard_image = RationCard_image + _ext;
                            string filePath = Path.Combine(Server.MapPath("/DeptLogo/") + RationCard_image);
                            RationCardImage = filePath;
                            _comPathRationCard = filePath;
                            file.SaveAs(_comPathRationCard);
                        }

                    }
                }
                //}
                SortedList list = new SortedList();

                list.Add("@OfficeId", Request.Form["OfficeId"].ToString());
                list.Add("@DeptId", Request.Form["DeptId"].ToString());
                list.Add("@BillingTitle", Request.Form["BillingTitle"].ToString());
                list.Add("@Address1", Request.Form["Address1"].ToString());
                list.Add("@OfficeLevel", Request.Form["OfficeLevel"].ToString());
                list.Add("@state", Request.Form["state"].ToString());
                list.Add("@Distirct", Request.Form["Distirct"].ToString());
                list.Add("@Block", Request.Form["Block"].ToString());
                list.Add("@Panchayat", Request.Form["Panchayat"].ToString());
                list.Add("@Pin", Request.Form["Pin"].ToString());
                list.Add("@GSTNo", Request.Form["GSTNo"].ToString());
                list.Add("@ContactNo", Request.Form["ContactNo"].ToString());
                list.Add("@EmailId", Request.Form["EmailId"].ToString());
                list.Add("@OfficeLogo", RationCard_image);
                list.Add("@OfficeSignature", Applicant_image);
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                mes = comfun.executeNonQueryWMessage("HpsedcDeptOffice_AcceptUpdate", "", list).ToString();


            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage(" ", "Error: " + ex.Message);

            }
            return Json(mes);

        }
        public ActionResult _DeptOfficeList()
        {
            // comfun.saveformname("CustomerPayment", "/TimeSheet/CustomerPayment", "Attendance Mgt/AttendanceApproval/CustomerPayment", "Customer Payment  Main form", "Y", "CustomerPayment", "CustomerPayment", "View", 1);
            //  ViewData["AccessRights"] = payfun.getAccessRights("CustomerPayment");
            DateTime monthYear = Convert.ToDateTime(Request.Form["BillDate"]);
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@OfficeId", Request.Form["OfficeId"].ToString());
            dt = comfun.fillDataTable("HpsedcDeptOffice_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        #endregion



        public ActionResult MapParentChildOffice()
        {
            return View();
        }
        public ActionResult _MapParentChildSubmit()
        {
            SortedList list = new SortedList();

            list.Add("@MappId", Request.Form["MappId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@ParentOfficeId", Request.Form["ParentOfficeId"].ToString());
            list.Add("@ChieldOfficeId", Request.Form["ChieldOfficeId"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@IsActive", Request.Form["IsActive"].ToString());
            string mes = comfun.executeNonQueryWMessage("HpsedcParentOffice_AcceptUpdate", "", list).ToString();

            return Json(mes);
        }

        public ActionResult _MapParentChildOfficeList()
        {
            // comfun.saveformname("CustomerPayment", "/TimeSheet/CustomerPayment", "Attendance Mgt/AttendanceApproval/CustomerPayment", "Customer Payment  Main form", "Y", "CustomerPayment", "CustomerPayment", "View", 1);
            //  ViewData["AccessRights"] = payfun.getAccessRights("CustomerPayment");
            DateTime monthYear = Convert.ToDateTime(Request.Form["BillDate"]);
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@DeptId", Request.Form["DeptId"].ToString());

            dt = comfun.fillDataTable("HpsedcParentOfficeMap_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _DeptOfficeddl()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId, Session["Customer"]);
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            dt = comfun.fillDataTable("HpsedcParentOffice_List", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }

        #endregion
        #region AVO
        public ActionResult AddAvo()
        {

            //DataSet ds = new DataSet();
            // = comfun.fillDataSet("stpProductdropown_list", null, null);
            return View();
        }
        public ActionResult _AddAvoList()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId", Session["Customer"]);
            list.Add("@AvoId", Request.Form["AvoId"].ToString());
            list.Add("@DeptId", Session["DeptId"]?.ToString() ?? "0");
            dt = comfun.fillDataTable("HpsedcAVO_List", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }



        public ActionResult _Depttdesignation()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId, Session["Customer"]);
            list.Add("@Id", Request.Form["EmpId"].ToString());
            dt = comfun.fillDataTable("stpSSOEmployeeInfo", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }

        public ActionResult _AvoDepartmentddlNew()
        {
            SortedList list = new SortedList();
            list.Add("@DeptId", Session["DeptId"]?.ToString() ?? "0");
            list.Add("@createdBy", Session["EmpId"].ToString());
            DataTable dt = comfun.fillDataTable("stpHPSEDCSSODepartmentDDL", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult _DepartmentddlNew1()
        {
            SortedList list = new SortedList();
            list.Add("@DeptId", Session["DeptId"]?.ToString() ?? "0");
            DataTable dt = comfun.fillDataTable("stpHPSEDCSSODepartmentDDL", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult _DepttdesignationDll()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId, Session["Customer"]);
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            dt = comfun.fillDataTable("stpSSODesignationdll", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public ActionResult _DeptDesigEmpddl()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@DesignationId", Request.Form["DesignationId"].ToString());

            dt = comfun.fillDataTable("stpHPSEDCSSOEmployeeDDL", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public ActionResult _EmpDetails()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId", Session["Customer"]);
            list.Add("@Id", Request.Form["EmpId"].ToString());
            dt = comfun.fillDataTable("stpSSOEmployeeInfo", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }

        public ActionResult _Nameddl()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId", Session["Customer"]);
            list.Add("@Id", Session["DeptId"]?.ToString() ?? "0");
            dt = comfun.fillDataTable("stpHPSEDCSSOEmployeeDDL", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }


        public ActionResult _AVODesignationddl()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId", Session["Customer"]);
            list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("HpsedcAVODesignation_dll", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public ActionResult _AvoSubmit()
        {
            SortedList list = new SortedList();
            list.Add("@AvoId", Request.Form["AvoId"].ToString());
            list.Add("@Name", Request.Form["Name"].ToString());
            list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
            list.Add("@DepartmentOfficeAddress", Request.Form["DepartmentOfficeAddress"].ToString());
            list.Add("@DepartmentOfficeAddressText", Request.Form["DepartmentOfficeAddressText"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@Mobileno", Request.Form["Mobileno"].ToString());
            list.Add("@EmailId", Request.Form["EmailId"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@IsActive", Request.Form["IsActive"].ToString());
            string mes = comfun.executeNonQueryWMessage("HpsedcAVO_AcceptUpdate", "", list).ToString();

            return Json(mes);
        }

        public ActionResult _AVOWithoutSubmit()
        {
            SortedList list = new SortedList();
            list.Add("@AvoId", Request.Form["AvoId"].ToString());
            list.Add("@Name", Request.Form["Name"].ToString());
            list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
            list.Add("@DepartmentOfficeAddress", Request.Form["DepartmentOfficeAddress"].ToString());
            list.Add("@DepartmentOfficeAddressText", Request.Form["DepartmentOfficeAddressText"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@Mobileno", Request.Form["Mobileno"].ToString());
            list.Add("@EmailId", Request.Form["EmailId"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@IsActive", Request.Form["IsActive"].ToString());
            string mes = comfun.executeNonQueryWMessage("HpsedcAVOWithouSSO_AcceptUpdate", "", list).ToString();

            return Json(mes);
        }

        #endregion
        #region Requisition
        public ActionResult AddRequisition()

        {
            return View();
        }

        public ActionResult _DeptRequisitionList(string RequisitionId)
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();//@SessionID INT=0,
            list.Add("@RequisitionId", RequisitionId);
            list.Add("@RequestBy", Session["EmpId"].ToString());
            list.Add("@UserRole", 6);
            dt = comfun.fillDataTable("HpsedcDeptRequisition_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _DeptHQOfficeddl()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId, Session["Customer"]);
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            dt = comfun.fillDataTable("HpsedcHqOffice_List", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public ActionResult _Agencyddl()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId, Session["Customer"]);
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            dt = comfun.fillDataTable("HpsedcAgencyDdl_List", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public ActionResult _HPSEDCAVOddl()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();

            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            dt = comfun.fillDataTable("HpsedcAVO_dll", "", list);

            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }

        public ActionResult _DeptChildOfficeddl()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId", Session["Customer"]);
            list.Add("@ParentOfficeId", Request.Form["ParentOfficeId"].ToString());
            dt = comfun.fillDataTable("HpsedcChildOffice_List", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public JsonResult _InsertDeptRequisition(string RequestDetails, string RequisitionId, string DepartmentCode, string DeploymentDate, string ParentOffice, string ParentOfficeText, string BillingAddress, string BillingAddressText, string IsSuggestAgency, string SuggestAgencyId, string TotalResource, string RequisitionType)
        {
            string mes = string.Empty;
            try
            {
                var Requisition_image = "";
                var _comPath = "";
                var RequisitionImage1 = "";
                var filePathA = "";
                HttpFileCollectionBase files = Request.Files;
                if (System.Web.HttpContext.Current.Request.Files.AllKeys.Any())
                {
                    // string mes2 = string.Empty;
                    for (int i = 0; i < files.Count; i++)
                    {
                        HttpPostedFileBase file = files[i];

                        // Checking for Internet Explorer  
                        if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                        {
                            string[] testfiles = file.FileName.Split(new char[] { '\\' });
                            RequisitionImage1 = testfiles[testfiles.Length - 1];

                        }
                        else
                        {
                            RequisitionImage1 = file.FileName;
                        }
                        var _ext = Path.GetExtension(RequisitionImage1);
                        Requisition_image = Path.GetFileNameWithoutExtension(RequisitionImage1);
                        string filePath = Path.Combine(Server.MapPath("/RequisitionDept_Attachement/") + Requisition_image + _ext);
                        Requisition_image = Requisition_image + _ext;
                        RequisitionImage1 = filePath;
                        filePathA = filePath;
                        file.SaveAs(filePathA);

                    }

                }

                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(RequestDetails);

                DataTable table = new DataTable();

                table.Columns.Add("ResourceId", typeof(int));
                table.Columns.Add("ResourceType", typeof(int));
                table.Columns.Add("TenureofDeployment", typeof(int));
                table.Columns.Add("DeploymentLoc", typeof(int));
                table.Columns.Add("BasicAmount", typeof(float));
                table.Columns.Add("Costtodepartment", typeof(float));
                table.Columns.Add("NoOfResource", typeof(int));


                foreach (var Item in jsonData)
                {
                    DataRow dr = table.NewRow();
                    dr["ResourceId"] = Item.ResourceId;
                    dr["ResourceType"] = Item.ResourceTypeId;
                    dr["TenureofDeployment"] = Item.TenureId;
                    dr["DeploymentLoc"] = 0;
                    dr["BasicAmount"] = Item.BasicAmount;
                    dr["Costtodepartment"] = Item.DeptCost;
                    dr["NoOfResource"] = Item.NoOfResource;

                    table.Rows.Add(dr);
                }
                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "HpsedcDeptRequisition_AcceptUpdate";
                SqlParameter parameter = new SqlParameter();

                com.Parameters.AddWithValue("@RequisitionId", RequisitionId);
                com.Parameters.AddWithValue("@DepartmentCode", DepartmentCode);
                com.Parameters.AddWithValue("@DeploymentDate", DeploymentDate);
                com.Parameters.AddWithValue("@ParentOffice", "0");
                com.Parameters.AddWithValue("@ParentOfficeText", ParentOfficeText);
                com.Parameters.AddWithValue("@BillingAddress", "0");
                com.Parameters.AddWithValue("@BillingAddressText", BillingAddressText);
                com.Parameters.AddWithValue("@IsSuggestAgency", IsSuggestAgency);
                com.Parameters.AddWithValue("@SuggestAgencyId", SuggestAgencyId);
                com.Parameters.AddWithValue("@RequestBy", Session["EmpId"]);
                com.Parameters.AddWithValue("@TotalResource", TotalResource);
                com.Parameters.AddWithValue("@RequisitionType", RequisitionType);
                com.Parameters.AddWithValue("@DeparmentOrderDocument", Requisition_image);

                com.Parameters.AddWithValue("@IsActive", "Y");
                SqlParameter parameter1 = new SqlParameter();
                parameter1.ParameterName = "@tblDeptResource2";
                parameter1.SqlDbType = System.Data.SqlDbType.Structured;
                parameter1.Value = table;
                com.Parameters.Add(parameter1);
                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);
                com.CommandTimeout = 0;

                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();
                com.ExecuteNonQuery();
                conObj.con.Close();

                mes = "Record saved successfully";
                mes = sp.Value.ToString();
                string retunvalue = (string)com.Parameters["@Mes"].Value;
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("POList", "Error: " + ex.Message);
            }
            return Json(mes);
        }
        public JsonResult _InsertDeptRequisitionDetails(string RequestDetails)
        {

            string mes = string.Empty;
            try
            {
                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(RequestDetails);

                DataTable table = new DataTable();

                table.Columns.Add("PostId", typeof(int));
                table.Columns.Add("ResourcePostId", typeof(int));
                table.Columns.Add("DeploymentOffice", typeof(string));
                table.Columns.Add("AVOId", typeof(int));

                foreach (var Item in jsonData)
                {
                    DataRow dr = table.NewRow();
                    dr["PostId"] = Item.PostId;
                    dr["ResourcePostId"] = Item.ResourceId;
                    dr["DeploymentOffice"] = Item.DeploymentLocId;
                    dr["AVOId"] = Item.AVOName;
                    table.Rows.Add(dr);
                }
                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "HpsedcRequisitionDetails_AcceptUpdate";
                SqlParameter parameter1 = new SqlParameter();
                parameter1.ParameterName = "@tblRequisitionDetails";
                parameter1.SqlDbType = System.Data.SqlDbType.Structured;
                parameter1.Value = table;
                com.Parameters.Add(parameter1);
                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);
                com.CommandTimeout = 0;

                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();
                com.ExecuteNonQuery();
                conObj.con.Close();

                mes = "Record saved successfully";// sp.Value.ToString(); //"Record saved successfully";
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("POList", "Error: " + ex.Message);
            }
            return Json(mes);
        }
        public ActionResult _RequisitionSalarySubmit()
        {
            //comfun.saveformname("_purchaseBillList", "/pm/_purchaseBillList", "Voucher/PurchaseBill", "Purchase Bill List", "N", "PurchaseBill", "PurchaseBill", "List", 4);
            string mes = string.Empty;
            SortedList list = new SortedList();

            try
            {
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@PostId", Request.Form["PostId"].ToString());
                list.Add("@ResourceId", Request.Form["ResourceId"].ToString());
                list.Add("@NoOfHours", Request.Form["NoOfHours"].ToString());
                list.Add("@IsAdminChargeFixed", Request.Form["IsAdminChargeFixed"].ToString());
                list.Add("@IsEsi", Request.Form["IsEsi"].ToString());
                list.Add("@IsPf", Request.Form["IsPf"].ToString());
                list.Add("@IsSpecialEsi", Request.Form["IsSpecialEsi"].ToString());
                list.Add("@BasicAmount", Request.Form["BasicAmount"].ToString());
                list.Add("@AdminCharge", Request.Form["AdminCharge"].ToString());
                list.Add("@SpecialAllowance", Request.Form["SpecialAllowance"].ToString());
                list.Add("@liblyAllowance", Request.Form["liblyAllowance"].ToString());
                list.Add("@Medical", Request.Form["Medical"].ToString());

                list.Add("@HRA", Request.Form["HRA"].ToString());
                list.Add("@Tribalallwoance", Request.Form["Tribalallwoance"].ToString());
                list.Add("@TribalAreaallwoance", Request.Form["TribalAreaallwoance"].ToString());
                list.Add("@ReliverCharges", Request.Form["ReliverCharges"].ToString());
                list.Add("@DressCharge", Request.Form["DressCharge"].ToString());
                list.Add("@OtherAllowance", Request.Form["OtherAllowance"].ToString());
                list.Add("@ChargeAmount", Request.Form["ChargeAmount"].ToString());
                list.Add("@HpsedcCharge", Request.Form["HpsedcCharge"].ToString());
                list.Add("@TaxableAmt", Request.Form["TaxableAmt"].ToString());
                list.Add("@CGST", Request.Form["CGST"].ToString());
                list.Add("@SGST", Request.Form["SGST"].ToString());
                list.Add("@GrandTotal", Request.Form["GrandTotal"].ToString());

                mes = comfun.executeNonQueryWMessage("HpsedcDeptRequisitionSalary_AcceptUpdate", "", list).ToString();
            }

            catch (Exception ex)
            {

                mes = "Error:" + ex.Message;
            }


            return Json(mes);
        }
        public ActionResult _RequisitionSalaryEdit()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@PostId", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("HpsedcRequisitionSalary_Edit", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult HpsedcRequisitionDetails_AcceptUpdate(DataTable tblRequisitionDetails)
        {
            string message = string.Empty;
            try
            {
                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "HpsedcRequisitionDetails_AcceptUpdate";



                SqlParameter param1 = com.Parameters.AddWithValue("@tblRequisitionDetails", tblRequisitionDetails);
                param1.SqlDbType = SqlDbType.Structured;
                param1.TypeName = "RequisitionDetails";

                SqlParameter param2 = com.Parameters.Add("@Mes", SqlDbType.VarChar, -1);
                param2.Direction = ParameterDirection.Output;

                com.ExecuteNonQuery();


                message = param2.Value.ToString();
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }

            return View();
        }
        public ActionResult _InsertDeptRequisitionEdit()
        {
            //comfun.saveformname("_purchaseBillList", "/pm/_purchaseBillList", "Voucher/PurchaseBill", "Purchase Bill List", "N", "PurchaseBill", "PurchaseBill", "List", 4);

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            //list.Add("@RequestId", Request.Form["Id"].ToString());
            //list.Add("@RequestBy", Session["EmpId"].ToString());
            list.Add("@RequisitionId", Request.Form["Id"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@UserRole", Session["RoleId"].ToString());
            DataSet ds = comfun.fillDataSet("HpsedcDeptRequisition_Edit", "", list);

            var json = JsonConvert.SerializeObject(ds);
            return Json(json);
        }
        public ActionResult _RequisitionDetailsEdit()
        {
            //comfun.saveformname("_purchaseBillList", "/pm/_purchaseBillList", "Voucher/PurchaseBill", "Purchase Bill List", "N", "PurchaseBill", "PurchaseBill", "List", 4);

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@PostId", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("HpsedcRequisitionDetails_Edit", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _RequisitionDetailsDelete()
        {
            //comfun.saveformname("_purchaseBillList", "/pm/_purchaseBillList", "Voucher/PurchaseBill", "Purchase Bill List", "N", "PurchaseBill", "PurchaseBill", "List", 4);
            string mes = string.Empty;
            SortedList list = new SortedList();

            try
            {
                list.Add("@PostId", Request.Form["PostId"].ToString());

                mes = comfun.executeNonQueryWMessage("HpsedcRequisitionDetails_Delete", "", list).ToString();
            }

            catch (Exception ex)
            {

                mes = "Error:" + ex.Message;
            }


            return Json(mes);
        }
        public ActionResult _RequisitionTempDetailsDelete()
        {
            //comfun.saveformname("_purchaseBillList", "/pm/_purchaseBillList", "Voucher/PurchaseBill", "Purchase Bill List", "N", "PurchaseBill", "PurchaseBill", "List", 4);
            string mes = string.Empty;
            SortedList list = new SortedList();

            try
            {
                list.Add("@PostId", Request.Form["PostId"].ToString());

                mes = comfun.executeNonQueryWMessage("HpsedcRequisitionTempDetails_Delete", "", list).ToString();
            }

            catch (Exception ex)
            {

                mes = "Error:" + ex.Message;
            }


            return Json(mes);
        }

        #endregion
        #region Letter

        public ActionResult ProposalCall()
        {
            return View();
        }
        public ActionResult ProposalSubmission()
        {
            return View();
        }
        #endregion
        #region RequisitionApproval
        public ActionResult RequisitionApproval()
        {
            //DataSet ds = new DataSet();
            // = comfun.fillDataSet("stpProductdropown_list", null, null);
            return View();
        }
        public ActionResult _RequisitionActionNameddl()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@ActionId", 0);
            list.Add("@UserRole", Session["RoleId"].ToString());
            dt = comfun.fillDataTable("HpsedcHpsedcAction_List", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public ActionResult _RequisitionGetAgency()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@ReqId", Request.Form["ReqId"].ToString());
            // list.Add("@UserRole", Session["RoleId"].ToString());
            dt = comfun.fillDataTable("HpsedcgetSuggestAgency_List", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public ActionResult _RequisitionActionSubmit()
        {

            string mes = "";
            try
            {

                var Requisition_ActionDoc = "";
                var _comPath = "";
                var RequisitionActionImage1 = "";
                var filePathA = "";
                HttpFileCollectionBase files = Request.Files;
                if (System.Web.HttpContext.Current.Request.Files.AllKeys.Any())
                {
                    // string mes2 = string.Empty;
                    for (int i = 0; i < files.Count; i++)
                    {
                        HttpPostedFileBase file = files[i];

                        // Checking for Internet Explorer  
                        if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                        {
                            string[] testfiles = file.FileName.Split(new char[] { '\\' });
                            RequisitionActionImage1 = testfiles[testfiles.Length - 1];

                        }
                        else
                        {
                            RequisitionActionImage1 = file.FileName;
                        }
                        var _ext = Path.GetExtension(RequisitionActionImage1);
                        Requisition_ActionDoc = Path.GetFileNameWithoutExtension(RequisitionActionImage1);
                        string filePath = Path.Combine(Server.MapPath("/RequisitionActionDoc/") + Requisition_ActionDoc + _ext);
                        Requisition_ActionDoc = Requisition_ActionDoc + _ext;
                        RequisitionActionImage1 = filePath;
                        filePathA = filePath;
                        file.SaveAs(filePathA);

                    }

                }

                DataTable dt = new DataTable();
                SortedList list = new SortedList();
                list.Add("@ReqActionId", Request.Form["ReqActionId"].ToString());
                list.Add("@ReqId", Request.Form["ReqId"].ToString());
                list.Add("@ActionNameId", Request.Form["ActionNameId"].ToString());
                list.Add("@Remarks", Request.Form["Remarks"].ToString());
                list.Add("@ActionTakenBy", Session["RoleId"].ToString());
                list.Add("@MoveTo", Request.Form["MoveTo"].ToString());
                list.Add("@Attchement", Requisition_ActionDoc);

                mes = comfun.executeNonQueryWMessage("HpsedcReqActionTaken_AcceptUpdate", "", list).ToString();

            }
            catch (Exception ex) { mes = ex.Message; }

            return Json(mes);
        }




        //public ActionResult _RequisitionActionSubmit()
        //{
        //    //@ReqActionId  ,@ReqId ,@ActionNameId ,@ActionTakenBy ,@MoveTo
        //    DataTable dt = new DataTable();
        //    SortedList list = new SortedList();
        //    list.Add("@ReqActionId", Request.Form["ReqActionId"].ToString());
        //    list.Add("@ReqId", Request.Form["ReqId"].ToString());
        //    list.Add("@ActionNameId", Request.Form["ActionNameId"].ToString());
        //    list.Add("@Remarks", Request.Form["Remarks"].ToString());
        //    list.Add("@ActionTakenBy", Session["RoleId"].ToString());          
        //    list.Add("@MoveTo", Request.Form["MoveTo"].ToString());

        //    string mes = "";
        //    try
        //    {
        //        mes = comfun.executeNonQueryWMessage("HpsedcReqActionTaken_AcceptUpdate", "", list).ToString();

        //    }
        //    catch (Exception ex) { mes = ex.Message; }

        //    return Json(mes);
        //}

        public ActionResult _ActionRemarksList()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@ReqActionId", "0");
            list.Add("@ReqId", Request.Form["ReqId"].ToString());
            dt = comfun.fillDataTable("HpsedcReqActionTaken_List", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _RequisitionApprovalList()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            list.Add("@RequestBy", Session["EmpId"].ToString());
            list.Add("@UserRole", Session["RoleId"].ToString());
            dt = comfun.fillDataTable("HpsedcRequisitionApproval_List1", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _GetRequisitionResource()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@RequisitionId", Request.Form["ReqId"].ToString());
            list.Add("@UserRole", Session["RoleId"].ToString());
            dt = comfun.fillDataTable("HpsedcRequisitonResource_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _GetRequisitionResourceUid()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@RequisitionId", Request.Form["ReqId"].ToString());
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@UserRole", Session["RoleId"].ToString());
            dt = comfun.fillDataTable("HpsedcRequisitonResourceUid_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _GetWorkOfferAgency()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@RequisitionId", Request.Form["ReqId"].ToString());
            list.Add("@UserRole", Session["RoleId"].ToString());
            dt = comfun.fillDataTable("HpsedcWorkOfferAgency_Get", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _InsertWorkOffer(string WorkOrderId, string WorkOfferDetails, string RequisitionId, string ActionId, string Agencyname, string Remarks, string MoveTo)
        {

            string mes = string.Empty;
            try
            {

                var Requisition_ActionDoc = "";
                var _comPath = "";
                var RequisitionActionImage1 = "";
                var filePathA = "";
                HttpFileCollectionBase files = Request.Files;
                if (System.Web.HttpContext.Current.Request.Files.AllKeys.Any())
                {
                    for (int i = 0; i < files.Count; i++)
                    {
                        HttpPostedFileBase file = files[i];

                        if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                        {
                            string[] testfiles = file.FileName.Split(new char[] { '\\' });
                            RequisitionActionImage1 = testfiles[testfiles.Length - 1];

                        }
                        else
                        {
                            RequisitionActionImage1 = file.FileName;
                        }
                        var _ext = Path.GetExtension(RequisitionActionImage1);
                        Requisition_ActionDoc = Path.GetFileNameWithoutExtension(RequisitionActionImage1);
                        string filePath = Path.Combine(Server.MapPath("/RequisitionActionDoc/") + Requisition_ActionDoc + _ext);
                        Requisition_ActionDoc = Requisition_ActionDoc + _ext;
                        RequisitionActionImage1 = filePath;
                        filePathA = filePath;
                        file.SaveAs(filePathA);

                    }

                }

                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(WorkOfferDetails);

                DataTable table = new DataTable();

                table.Columns.Add("RequisitionId", typeof(int));
                table.Columns.Add("ResourceId", typeof(int));
                table.Columns.Add("NoOfResource", typeof(int));
                table.Columns.Add("AgencyId", typeof(int));

                foreach (var Item in jsonData)
                {
                    DataRow dr = table.NewRow();

                    dr["RequisitionId"] = Item.RequisitionId;
                    dr["ResourceId"] = Item.ResourceId;
                    dr["NoOfResource"] = Item.NoOfResource;
                    dr["AgencyId"] = Item.AgencyId;
                    table.Rows.Add(dr);
                }
                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "HpsedcWorkOfferAction1_AcceptUpdate";

                SqlParameter parameter = new SqlParameter();

                com.Parameters.AddWithValue("@WorkOrderId", WorkOrderId);
                com.Parameters.AddWithValue("@RequisitionId", RequisitionId);
                com.Parameters.AddWithValue("@AgencyName", Agencyname);
                com.Parameters.AddWithValue("@DaRemarks", Remarks);
                com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());
                com.Parameters.AddWithValue("@MoveTo", MoveTo);
                com.Parameters.AddWithValue("@ActionId", ActionId);
                com.Parameters.AddWithValue("@ActionTakenBy", Session["RoleId"].ToString());
                com.Parameters.AddWithValue("@Attachement", Requisition_ActionDoc);

                SqlParameter parameter1 = new SqlParameter();
                parameter1.ParameterName = "@tblWorkOrderDetails";
                parameter1.SqlDbType = System.Data.SqlDbType.Structured;
                parameter1.Value = table;
                com.Parameters.Add(parameter1);
                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);
                com.CommandTimeout = 0;

                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();
                com.ExecuteNonQuery();
                conObj.con.Close();

                mes = "Record saved successfully";
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("POList", "Error: " + ex.Message);
            }
            return Json(mes);
        }
        public ActionResult WorkOfferEmp()
        {
            SortedList list = new SortedList();
            list.Add("@CreatedBy", Session["Empid"].ToString());
            list.Add("@UserRole", Session["RoleId"].ToString());
            DataTable dt = comfun.fillDataTable("HpsedcAgencyWorkOffer_View", "", list);

            return View(dt);
        }
        public ActionResult HpsedcWorkOrder_Dll()
        {
            SortedList list = new SortedList();
            list.Add("@WorkOrder", Request.Form["WorkOrder"].ToString());
            DataTable dt = comfun.fillDataTable("HpsedcWorkOrder_Dll", "", list);

            return View(dt);
        }
        public ActionResult HpsedcWorkOrderRUId_Dll()
        {
            SortedList list = new SortedList();
            list.Add("@WorkOrder", Request.Form["WorkOrder"].ToString());
            list.Add("@DesignationId", Request.Form["WorkOrder"].ToString());
            DataTable dt = comfun.fillDataTable("HpsedcWorkOrderPostUid_Dll", "", list);
            return View(dt);
        }
        public ActionResult _HpsedcWorkOrderDeployedSubmit()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@ReqActionId", Request.Form["ReqActionId"].ToString());
            list.Add("@ReqId", Request.Form["ReqId"].ToString());
            list.Add("@ActionNameId", Request.Form["ActionNameId"].ToString());
            list.Add("@Remarks", Request.Form["Remarks"].ToString());
            list.Add("@ActionTakenBy", Session["RoleId"].ToString());
            list.Add("@MoveTo", Request.Form["MoveTo"].ToString());
            string mes = "";
            try
            {
                mes = comfun.executeNonQueryWMessage("HpsedcDeployment_AcceptUpdate", "", list).ToString();

            }
            catch (Exception ex) { mes = ex.Message; }

            return Json(mes);
        }
        public ActionResult UserDepartment()
        {
            var hpdeseInfo = (List<Payroll.portal.Models.Department.Root>)Session["HPSEDCLoginInformation"];
            return View(hpdeseInfo);
        }
        public ActionResult ChooseDepartment(string id)
        {
            Session["SelectedDepartmentId"] = id;

            SortedList list = new SortedList();
            list.Add("@DeptId", id);
            list.Add("@UserRole", "0");

            DataTable dt = comfun.fillDataTable("HpsedcGetDeptName_list", "", list);
            if (dt.Rows.Count > 0)
            {
                Session["DeptName"] = dt.Rows[0]["departmentName"].ToString();
                Session["DeptId"] = dt.Rows[0]["departmentID"].ToString();

                Response.Cookies["DeptName"].Value = dt.Rows[0]["departmentName"].ToString();

                if (dt.Rows[0]["UserRole"].ToString() != "0")
                {
                    Session["RoleId"] = dt.Rows[0]["UserRole"].ToString();
                }



            }

            return RedirectToAction("Dashboard", "TimeSheet");
        }
        public ActionResult WorkOffer()
        {
            SortedList list = new SortedList();
            list.Add("@CreatedBy", Session["Empid"].ToString());
            list.Add("@UserRole", Session["RoleId"].ToString());

            DataTable dt = comfun.fillDataTable("HpsedcAgencyWorkOffer_View", "", list);
            return View(dt);
        }
        public ActionResult _WorkOfferList()
        {

            SortedList list = new SortedList();
            list.Add("@CreatedBy", Session["Empid"].ToString());
            list.Add("@UserRole", Session["RoleId"].ToString());
            list.Add("@WorkOrder", Request.Form["WorkOrder"].ToString());
            //DataTable dt = comfun.fillDataTable("HpsedcAgencyWorkOffer_View", "", list);
            DataTable dt = comfun.fillDataTable("HpsedcAgencyWorkOffer_View2", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _RequisitionActionSubmit1()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@ReqActionId", Request.Form["ReqActionId"].ToString());
            list.Add("@ReqId", Request.Form["ReqId"].ToString());
            list.Add("@ActionNameId", Request.Form["ActionNameId"].ToString());
            list.Add("@Remarks", Request.Form["Remarks"].ToString());
            list.Add("@ActionTakenBy", Session["RoleId"].ToString());
            list.Add("@MoveTo", Request.Form["MoveTo"].ToString());
            string mes = "";
            try
            {
                mes = comfun.executeNonQueryWMessage("HpsedcReqActionTaken_AcceptUpdate", "", list).ToString();

            }
            catch (Exception ex) { mes = ex.Message; }

            return Json(mes);
        }


        public ActionResult _WorkOfferAgencyAction()
        {
            string mes = "";
            try
            {

                var UploadSignedOffer = "";
                var _comPath = "";
                var UploadSignedOffer1 = "";
                var filePathA = "";
                HttpFileCollectionBase files = Request.Files;
                if (System.Web.HttpContext.Current.Request.Files.AllKeys.Any())
                {
                    for (int i = 0; i < files.Count; i++)
                    {
                        HttpPostedFileBase file = files[i];

                        if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                        {
                            string[] testfiles = file.FileName.Split(new char[] { '\\' });
                            UploadSignedOffer1 = testfiles[testfiles.Length - 1];

                        }
                        else
                        {
                            UploadSignedOffer1 = file.FileName;
                        }
                        var _ext = Path.GetExtension(UploadSignedOffer1);
                        UploadSignedOffer = Path.GetFileNameWithoutExtension(UploadSignedOffer1);
                        string filePath = Path.Combine(Server.MapPath("/RequisitionDept_Attachement/") + UploadSignedOffer + _ext);
                        UploadSignedOffer = UploadSignedOffer + _ext;
                        UploadSignedOffer1 = filePath;
                        filePathA = filePath;
                        file.SaveAs(filePathA);

                    }

                }


                SortedList list1 = new SortedList();
                list1.Add("@WorkOrderId", Request.Form["WorkOrderId"].ToString());
                list1.Add("@RequisitionId", Request.Form["ReqId"].ToString());
                list1.Add("@AgencyActionStatus", Request.Form["AgencyActionStatus"].ToString());
                list1.Add("@CreatedBy", Session["EmpId"].ToString());
                list1.Add("@UploadSignedOffer", UploadSignedOffer);


                SortedList list = new SortedList();
                list.Add("@ReqActionId", Request.Form["ReqActionId"].ToString());
                list.Add("@ReqId", Request.Form["ReqId"].ToString());
                list.Add("@ActionNameId", Request.Form["ActionNameId"].ToString());
                list.Add("@Remarks", Request.Form["Remarks"].ToString());
                list.Add("@ActionTakenBy", Session["RoleId"].ToString());
                list.Add("@MoveTo", Request.Form["MoveTo"].ToString());
                try
                {
                    //mes = comfun.executeNonQueryWMessage("HpsedcWorkOfferAgencyAction_AcceptUpdate", "", list1).ToString();
                    mes = comfun.executeNonQueryWMessage("HpsedcWorkOfferAgencyAction_AcceptUpdate2", "", list1).ToString();


                    mes = comfun.executeNonQueryWMessage("HpsedcReqActionTaken_AcceptUpdate", "", list).ToString();

                }


                catch (Exception ex) { mes = ex.Message; }
            }
            catch (Exception ex) { mes = ex.Message; }

            return Json(mes);

        }

        #endregion
        #region Test
        public ActionResult Test()
        {

            //DataSet ds = new DataSet();
            // = comfun.fillDataSet("stpProductdropown_list", null, null);
            return View();
        }
        public ActionResult GeneratePdf()
        {
            // Retrieve your data from a data source
            List<string> data = GetDataFromSource();

            // Create a new PDF document
            Document document = new Document();

            // Set the response content type and headers
            Response.ContentType = "application/pdf";
            Response.AddHeader("content-disposition", "attachment;filename=example.pdf");
            Response.Cache.SetCacheability(HttpCacheability.NoCache);

            // Create a PdfWriter to write the content to the response stream
            PdfWriter writer = PdfWriter.GetInstance(document, Response.OutputStream);

            // Open the document
            document.Open();

            // Add data to the document
            foreach (var item in data)
            {
                document.Add(new Paragraph(item));
            }

            // Close the document
            document.Close();

            // Return the PDF file
            return new EmptyResult();
        }

        private List<string> GetDataFromSource()
        {
            // Simulated data retrieval from a source
            return new List<string> { "Data 1", "Data 2", "Data 3" };
        }
        #endregion
        #region admin

        public ActionResult RequisitionEdit(int RequisitionId)
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();

            list.Add("@RequisitionId", RequisitionId);
            dt = comfun.fillDataTable("HpsedcDeptRequisition_getbyId", "", list);

            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult MasterApproval()
        {

            //DataSet ds = new DataSet();
            // = comfun.fillDataSet("stpProductdropown_list", null, null);
            return View();
        }

        public ActionResult _MasterApprovalSubmit()
        {
            SortedList list = new SortedList();
            list.Add("@ApprovalId", Request.Form["ApprovalId"].ToString());
            list.Add("@IsResourceLtEq10", Request.Form["IsResourceLtEq10"].ToString());
            list.Add("@IsSuggestAgency", Request.Form["IsSuggestAgency"].ToString());
            list.Add("@IsDAApproval", Request.Form["IsDAApproval"].ToString());
            list.Add("@IsDCApproval", Request.Form["IsDCApproval"].ToString());
            list.Add("@IsGMApproval", Request.Form["IsGMApproval"].ToString());
            list.Add("@IsMDApproval", Request.Form["IsMDApproval"].ToString());
            list.Add("@IsAgencyAction", Request.Form["IsAgencyAction"].ToString());
            list.Add("@IsCommitteeApproval", Request.Form["IsCommitteeApproval"].ToString());
            //ApprovalId	IsResourceLtEq10	IsSuggestAgency	IsDAApproval	IsDCApproval	IsGMApproval	IsMDApproval
            //IsAgencyAction	IsCommitteeApproval	IsActive	CreatedBy	CreatedDate	ModifiedBy	ModifiedDate
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@IsActive", Request.Form["IsActive"].ToString());
            string mes = comfun.executeNonQueryWMessage("HpsedRequisitionApprovalcon_AcceptUpdate", "", list).ToString();

            return Json(mes);
        }
        public ActionResult _MasterApprovalList()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId", Session["Customer"]);
            list.Add("@ApprovalId", Request.Form["ApprovalId"].ToString());
            dt = comfun.fillDataTable("HpsedRequisitionApprovalcon_List", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult MasterCommitteeMember()
        {

            //DataSet ds = new DataSet();
            // = comfun.fillDataSet("stpProductdropown_list", null, null);
            return View();
        }
        public ActionResult _AddCommitteeMemberist()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId", Session["Customer"]);
            list.Add("@CmtMemberId", Request.Form["MemberId"].ToString());
            dt = comfun.fillDataTable("HpsedcCommitteeMember_List", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _CommitteeMemberSubmit()
        {
            SortedList list = new SortedList();
            list.Add("@CmtMemberId", Request.Form["MemberId"].ToString());
            list.Add("@MemberName", Request.Form["Name"].ToString());
            list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@Mobileno", Request.Form["MobileNo"].ToString());
            list.Add("@EmailId", Request.Form["email"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@DepartmentAddress", Request.Form["DepartmentAddress"].ToString());
            list.Add("@IsActive", Request.Form["IsActive"].ToString());
            string mes = comfun.executeNonQueryWMessage("HpsedcCommitteeMember_AcceptUpdate", "", list).ToString();

            return Json(mes);
        }
        public ActionResult AddRequisitionAction()
        {

            //DataSet ds = new DataSet();
            // = comfun.fillDataSet("stpProductdropown_list", null, null);
            return View();
        }
        public ActionResult _UserRoleddl()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId", Session["Customer"]);
            list.Add("@RoleId ", 0);
            dt = comfun.fillDataTable("HpsedcUserRoleDdl_List", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public ActionResult _ActionNameddl()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId", Session["Customer"]);
            list.Add("@ActionId", 0);
            dt = comfun.fillDataTable("HpsedcActionName_list", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }

        public ActionResult _AddActionList()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId", Session["Customer"]);
            list.Add("@ActionId", Request.Form["ActionId"].ToString());
            dt = comfun.fillDataTable("HpsedcHpsedcAction_List", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _AddActionSubmit()
        {
            SortedList list = new SortedList();
            list.Add("@ActionId", Request.Form["ActionId"].ToString());
            list.Add("@ActionName", Request.Form["ActionName"].ToString());
            list.Add("@ActionTakenBy", Request.Form["ActionTakenBy"].ToString());
            list.Add("@MoveTo", Request.Form["MoveTo"].ToString());
            list.Add("@IsActive", Request.Form["IsActive"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            string mes = comfun.executeNonQueryWMessage("HpsedcHpsedcAction_AcceptUpdate", "", list).ToString();

            return Json(mes);
        }
        public ActionResult Dashboard()
        {
            SortedList list = new SortedList();
            list.Add("@UserId", Session["Empid"].ToString());
            list.Add("@UserRole", Session["RoleId"].ToString());
            DataTable dt = comfun.fillDataTable("HpsedcDashBoardModule_List", "", list);
            return View(dt);
            //DataSet ds = new DataSet();
            // = comfun.fillDataSet("stpProductdropown_list", null, null);
            //  return View();
        }
        public JsonResult DashboardList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@ModuleId", Request.Form["ModuleId"].ToString());
            list.Add("@UserRole", Session["RoleId"]);
            dt = comfun.fillDataTable("HpsedcDashBoardSubModule_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        #endregion
        //CommitteeMemberSubmit
        #region LoginError
        public ActionResult LoginError()
        {
            return View();
        }
        #endregion
        //#region test Vishu

        //public ActionResult HpsedcModule()
        //{
        //    return View();
        //}
        //public JsonResult _ModuleList()
        //{
        //    DataTable dt = new DataTable();
        //    SortedList list = new SortedList();
        //    list.Add("@ModuleId", Request.Form["ModuleId"].ToString());
        //    dt = comfun.fillDataTable("HpsedcModule_List", "", list);
        //    var json = JsonConvert.SerializeObject(dt);
        //    return Json(json);
        //}

        //public JsonResult _ModuleEdit()
        //{
        //    DataTable dt = new DataTable();
        //    SortedList list = new SortedList();
        //    list.Add("@ModuleId", Request.Form["ModuleId"].ToString());
        //    dt = comfun.fillDataTable("HpsedcModule_List", "", list);
        //    var json = JsonConvert.SerializeObject(dt);
        //    return Json(json);
        //}
        //public JsonResult _ModuleSubmit()
        //{
        //    //comfun.saveformname("HiringExecutiveSubmit", "/HRNew/HiringExecutiveSubmit", "HRNew/HiringExecutive", "HiringExecutive  Add", "N", "HiringExecutive", "HiringExecutive", "Add", 2);
        //    //comfun.saveformname("HiringExecutiveSubmit", "/HRNew/HiringExecutiveSubmit", "HRNew/HiringExecutive", "HiringExecutive", "N", "HiringExecutive", "HiringExecutive", "Submit", 3);
        //    string mes = string.Empty;
        //    SortedList list = new SortedList();
        //    try


        //    {
        //        var Icon_image = "";
        //        var _comPath = "";
        //        var IconImage1 = "";
        //        var filePathA = "";
        //        HttpFileCollectionBase files = Request.Files;
        //        if (System.Web.HttpContext.Current.Request.Files.AllKeys.Any())
        //        {
        //            // string mes2 = string.Empty;
        //            for (int i = 0; i < files.Count; i++)
        //            {
        //                HttpPostedFileBase file = files[i];

        //                // Checking for Internet Explorer  
        //                if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
        //                {
        //                    string[] testfiles = file.FileName.Split(new char[] { '\\' });
        //                    IconImage1 = testfiles[testfiles.Length - 1];

        //                }
        //                else
        //                {
        //                    IconImage1 = file.FileName;
        //                }
        //                var _ext = Path.GetExtension(IconImage1);
        //                Icon_image = Path.GetFileNameWithoutExtension(IconImage1);
        //                string filePath = Path.Combine(Server.MapPath("/DashboardLogo/") + Icon_image + _ext);
        //                Icon_image = Icon_image + _ext;
        //                IconImage1 = filePath;
        //                filePathA = filePath;
        //                file.SaveAs(filePathA);

        //            }


        //        }
        //        list.Add("@ModuleId", Request.Form["ModuleId"].ToString());
        //        list.Add("@ModuleName", Request.Form["ModuleName"].ToString());
        //        list.Add("@MenuType", Request.Form["MenuType"].ToString());
        //        list.Add("@Description", Request.Form["Description"].ToString());
        //        list.Add("@link", Request.Form["link"].ToString());
        //        list.Add("@IsActive", Request.Form["IsActive"].ToString());
        //        list.Add("@Icon ", Icon_image);

        //        mes = comfun.executeNonQueryWMessage("HpsedcModule_AcceptUpdate", "", list).ToString();
        //    }

        //    catch (Exception ex)
        //    {

        //        mes = "Error:" + ex.Message;
        //    }

        //    return Json(mes);
        //}


        //#endregion

        //#region MapModule Vishu

        //public ActionResult MapModule()
        //{
        //    return View();
        //}
        //public JsonResult _MapModuleSubmit()
        //{

        //    string mes = string.Empty;
        //    SortedList list = new SortedList();
        //    try
        //    {
        //        list.Add("@MapId", Request.Form["MapId"].ToString());
        //        list.Add("@ModuleId", Request.Form["ModuleId"].ToString());
        //        list.Add("@SubModuleId", Request.Form["SubModuleId"].ToString());
        //        list.Add("@IsActive", Request.Form["IsActive"].ToString());

        //        mes = comfun.executeNonQueryWMessage("HpsedcModuleMap_AcceptUpdate", "", list).ToString();
        //    }

        //    catch (Exception ex)
        //    {

        //        mes = "Error:" + ex.Message;
        //    }

        //    return Json(mes);
        //}
        //public JsonResult _MapModuleList()
        //{
        //    DataTable dt = new DataTable();
        //    SortedList list = new SortedList();
        //    list.Add("@MapId", Request.Form["MapId"].ToString());
        //    dt = comfun.fillDataTable("HpsedcModuleMap_List", "", list);
        //    var json = JsonConvert.SerializeObject(dt);
        //    return Json(json);
        //}

        //public JsonResult _MapModuleEdit()
        //{
        //    DataTable dt = new DataTable();
        //    SortedList list = new SortedList();
        //    list.Add("@MapId", "0");
        //    dt = comfun.fillDataTable("HpsedcModuleMap_List", "", list);
        //    var json = JsonConvert.SerializeObject(dt);
        //    return Json(json);
        //}

        //public ActionResult _MapModuleddl()
        //{

        //    DataTable dt = new DataTable();
        //    SortedList list = new SortedList();
        //    // list.Add("@CustomerId", Session["Customer"]);
        //    list.Add("@ModuleId", Request.Form["ModuleId"].ToString());
        //    list.Add("@MenuType", Request.Form["MenuType"].ToString());
        //    dt = comfun.fillDataTable("HpsedcModule_ddl", "", list);
        //    //return PartialView("_Employeelist", dt);
        //    var json = JsonConvert.SerializeObject(dt);
        //    return Json(json);

        //}

        //public ActionResult _SubModuleddl()
        //{

        //    DataTable dt = new DataTable();
        //    SortedList list = new SortedList();
        //    // list.Add("@CustomerId", Session["Customer"]);
        //    list.Add("@ModuleId", Request.Form["ModuleId"].ToString());
        //    list.Add("@MenuType", Request.Form["MenuType"].ToString());
        //    dt = comfun.fillDataTable("HpsedcModule_ddl", "", list);
        //    //return PartialView("_Employeelist", dt);
        //    var json = JsonConvert.SerializeObject(dt);
        //    return Json(json);

        //}


        //#endregion
        //public ActionResult AssingSubMenu()
        //{
        //    return View();
        //}

        //public JsonResult AssingSubMenuList()
        //{
        //    DataTable dt = new DataTable();
        //    SortedList list = new SortedList();
        //    list.Add("@AccessId", Request.Form["AccessId"].ToString());
        //    dt = comfun.fillDataTable("HpsedcModuleAccess_List", "", list);
        //    var json = JsonConvert.SerializeObject(dt);
        //    return Json(json);
        //}
        //public JsonResult AssingSubMenuSubmit()
        //{
        //    string mes = string.Empty;
        //    SortedList list = new SortedList();
        //    try
        //    {
        //        list.Add("@AccessId", "0");
        //        list.Add("@ModuleId", "0");
        //        list.Add("@AccessBy", "0");
        //        list.Add("@IsActive", Request.Form["IsActive"].ToString());

        //        mes = comfun.executeNonQueryWMessage("HpsedcModuleAccess_AcceptUpdate", "", list).ToString();
        //    }

        //    catch (Exception ex)
        //    {

        //        mes = "Error:" + ex.Message;
        //    }

        //    return Json(mes);
        //}
        //public JsonResult AssingSubMenuEdit()
        //{
        //    DataTable dt = new DataTable();
        //    SortedList list = new SortedList();

        //    list.Add("@AccessId", "0");
        //    dt = comfun.fillDataTable("HpsedcModuleAccess_List", "", list);
        //    var json = JsonConvert.SerializeObject(dt);
        //    return Json(json);
        //}

        //public ActionResult AssingSubMenuddl()
        //{

        //    DataTable dt = new DataTable();
        //    SortedList list = new SortedList();
        //    // list.Add("@CustomerId", Session["Customer"]);
        //    list.Add("@ModuleId", Request.Form["ModuleId"].ToString());
        //    list.Add("@MenuType", Request.Form["MenuType"].ToString());
        //    dt = comfun.fillDataTable("HpsedcModule_ddl", "", list);
        //    //return PartialView("_Employeelist", dt);
        //    var json = JsonConvert.SerializeObject(dt);
        //    return Json(json);

        //}
        //public ActionResult Roleddl()
        //{
        //    SortedList list = new SortedList();
        //    list.Add("@UserRole", Request.Form["UserRole"].ToString());
        //    DataTable dt = comfun.fillDataTable("HpsedcGetRole_List", "", list);
        //    var json1 = JsonConvert.SerializeObject(dt);
        //    return Json(json1);
        //}

        public ActionResult ProposalCall2()
        {
            return View();
        }

        public ActionResult WorkOffer2()
        {
            return View();
        }

        public ActionResult AgencyWorkOffer()
        {
            return View();
        }
        public ActionResult TripartyAgreement()
        {
            return View();
        }
        public JsonResult _AgencyWorkOfferList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@CreatedBy", Session["Empid"].ToString());
            list.Add("@UserRole", Session["RoleId"].ToString());

            //dt = comfun.fillDataTable("HpsedcAgencyWorkOffer1_List1", "", list);
            dt = comfun.fillDataTable("HpsedcAgencyWorkOffer1_List2", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        #region Attandance
        public ActionResult Attandance()
        {
            return View();
        }
        public ActionResult _Attendance()
        {

            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@FromDate", Request.Form["FromDate"].ToString());
            list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
            list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
            list.Add("@BranchId", Request.Form["BranchId"].ToString());
            list.Add("@CustomerId", Session["Customer"].ToString());
            DataTable dt = comfun.fillDataTable("stpEmployeeAttendanceApprovedList", "", list);
            return PartialView("_Attendance", dt);
        }
        public ActionResult OpenMyCalendar()
        {
            return PartialView("_CalendarView");
        }

        public ActionResult _MyCalendar()
        {
            try
            {
                int startMonth = 0;
                int startYear = 0;
                if (Request.Form["StartMonth"] != null)
                {
                    startMonth = Convert.ToInt32(Request.Form["StartMonth"]);
                }
                if (Request.Form["StartYear"] != null)
                {
                    startYear = Convert.ToInt32(Request.Form["StartYear"]);
                }
                ViewData["StartYear"] = startYear;
                ViewData["StartMonth"] = startMonth;
                ViewData["EmpId"] = Request.Form["EmpId"].ToString();

                DateTime startDate = new DateTime(startYear, startMonth, 1);
                DateTime endDate = startDate.AddMonths(1).AddDays(-1);
                DateTime endMDate = startDate.AddMonths(1).AddMonths(1);
                return PartialView(GetAttendances(startDate, endDate, Request.Form["EmpId"].ToString()));
            }
            catch (Exception ex)
            {
                ViewData["Error"] = ex.Message;
                return PartialView("_error");
            }

        }

        private DataTable GetAttendances(DateTime fromDate, DateTime toDate, string empId)
        {
            SortedList list = new SortedList();
            list.Add("@EmpCode", empId);
            list.Add("@EmpId", empId);
            list.Add("@FromDate", fromDate.ToString("dd-MMM-yyyy"));
            list.Add("@ToDate", toDate.ToString("dd-MMM-yyyy"));
            list.Add("@CostCenterId", 1);
            DataTable dt = comfun.fillDataTable("stpEmployeeAtendanceCalendar", "", list);
            return dt;
        }
        public ActionResult _attendanceChangeDayStatusnew()
        {

            SortedList list = new SortedList();
            list.Add("@EmpID", Request.Form["EmpID"].ToString());
            list.Add("@AttendanceDate", Convert.ToDateTime(Request.Form["AttendanceDate"]).ToString("dd-MMM-yyyy"));

            DataTable dt = comfun.fillDataTable("stpEmployeeAttendance_RegulizeDayID", "", list);
            var json = JsonConvert.SerializeObject(dt);

            return Json(json);
        }


        #endregion

        #region Vishu
        public ActionResult PaymentView()
        {
            return View();
        }

        public ActionResult _PaymentViewSubmit()
        {
            SortedList list = new SortedList();
            list.Add("@PaymentId", Request.Form["PaymentId"].ToString());
            list.Add("@AgainstBillNo", Request.Form["AgainstBillNo"].ToString());
            list.Add("@Narration", Request.Form["Narration"].ToString());
            list.Add("@PaymentDate", Request.Form["PaymentDate"].ToString());
            //list.Add("@PaymentDate", Convert.ToDateTime(Request.Form["PaymentDate"].ToString()).ToString("MMyyyy"));

            list.Add("@ModeOfPayment", Request.Form["ModeOfPayment"].ToString());
            list.Add("@TransactionId", Request.Form["TransactionId"].ToString());
            list.Add("@PaymentAmount", Request.Form["PaymentAmount"].ToString());
            list.Add("@Balance", Request.Form["Balance"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            string mes = comfun.executeNonQueryWMessage("HpsedcAgencyPayment_AcceptUpdate", "", list).ToString();

            return Json(mes);

        }

        public JsonResult _HpsedcAgencyPaymentList()

        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@ID", Request.Form["ID"].ToString());
            list.Add("@RoleId", Request.Form["RoleId"].ToString());
            list.Add("@CreatedBy", Session["Empid"].ToString());
            dt = comfun.fillDataTable("HpsedcAgencyPayment_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
            //return Json(json, JsonRequestBehavior.AllowGet);
        }

        public ActionResult _HpsedcAgencyInvoiceDetail_get()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();

            list.Add("@AccountId", Request.Form["AccountId"].ToString());
            list.Add("@RoleId", Request.Form["RoleId"].ToString());
            list.Add("@CreatedBy", Session["Empid"].ToString());

            dt = comfun.fillDataTable("HpsedcAgencyInvoiceDetail_get", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);



        }

        public ActionResult EmployeeAdd1()
        {
            return View();
        }

        public ActionResult Dist()
        {
            return View();
        }

        public ActionResult AssingUserRole()
        {
            return View();
        }
        public ActionResult _HpsedcUserRoleddl()
        {
            SortedList list = new SortedList();
            list.Add("@RoleId", Request.Form["RoleId"].ToString());
            DataTable dt = comfun.fillDataTable("HpsedcUserRoleDdl_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _HpsedcUserRoleList()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@UserId", Request.Form["UserId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            dt = comfun.fillDataTable("HpsedcUserRole_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _HpsedcUserRoleSubmit()
        {
            SortedList list = new SortedList();
            list.Add("@UserId", Request.Form["UserId"].ToString());
            list.Add("@RoleId", Request.Form["RoleId"].ToString());

            string mes = comfun.executeNonQueryWMessage("HpsedcAssignRole_AcceptUpdate", "", list).ToString();

            return Json(mes);
        }
        #endregion


        public ActionResult SelectPortal()
        {
            //var hpdeseInfo = (List<Payroll.portal.Models.Department.EmployeeSelectPortal>)Session["HPSEDCLoginInformation"];

            //var hpdeseInfo = (List<Payroll.portal.Models.Department.Root>)Session["HPSEDCLoginInformation"];
            //var hpdeseInfo = (List<Payroll.portal.Models.Department.Root>)Session["HPSEDCLoginInformationNew"];
            // Payroll.portal.Models.Department.EmployeeSelectPortal x = new Payroll.portal.Models.Department.EmployeeSelectPortal();
            // var hpdeseInfo1 = (Payroll.portal.Models.Department.EmployeeSelectPortal)(hpdeseInfo);
            //  x = JsonConvert.SerializeObject(hpdeseInfo1.);
           
            // Payroll.portal.Models.Department.EmployeeSelectPortal x = new Payroll.portal.Models.Department.EmployeeSelectPortal();
            
            // x = Payroll.portal.Models.Department.EmployeeSelectPortal>hpdeseInfo;
            // x.EmployeeName

            return View();
        }
        public ActionResult SelectPortalSubmit()
        {

            Session["RoleId"] = 30;
            return Json("df");
        }


        #region EducationalDetail 
        public ActionResult EducationalDetail()
        {

            if (Session["Empid"] == null)
            {

                Session["Empid"] = GetEmpId();
            }
            return View();
        }

        private string GetEmpId()
        {

            return "12345";
        }

        public JsonResult _EducationalDetailSubmit(string RequestDetails, string EmpId, string EmpName, string Fathername, string Mothername, string Gender,
             string Dob, string Employeeaddress, string Panchayat, string state, string Distirct, string Block, string Pin, string Mobileno,
             string Email, string AadharNo)
        {
            string mes = string.Empty;
            try
            {
                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(RequestDetails);

                DataTable table = new DataTable();

                table.Columns.Add("EducationLavel", typeof(int));
                table.Columns.Add("InstitutationName", typeof(string));
                table.Columns.Add("PassingYear", typeof(int));
                table.Columns.Add("Percentage", typeof(int));

                foreach (var Item in jsonData)
                {

                    DataRow dr = table.NewRow();
                    dr["EducationLavel"] = Item.EducationalLevelId;
                    dr["InstitutationName"] = Item.InstituateName;
                    dr["PassingYear"] = Item.YearOfPassing;
                    dr["Percentage"] = Item.Percentage;

                    table.Rows.Add(dr);
                }
                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "HpsedcDeirctApplyEmp_AcceptUpdate";
                SqlParameter parameter = new SqlParameter();

                com.Parameters.AddWithValue("@EmpId", EmpId);
                com.Parameters.AddWithValue("@EmpName", EmpName);
                com.Parameters.AddWithValue("@Fathername", Fathername);
                com.Parameters.AddWithValue("@Mothername", Mothername);
                com.Parameters.AddWithValue("@Gender", Gender);
                com.Parameters.AddWithValue("@Dob", Dob);
                com.Parameters.AddWithValue("@Education", "");
                com.Parameters.AddWithValue("@Employeeaddress", Employeeaddress);
                com.Parameters.AddWithValue("@Panchayat", Panchayat);
                com.Parameters.AddWithValue("@state", state);

                com.Parameters.AddWithValue("@Distirct", Distirct);
                com.Parameters.AddWithValue("@Block", Block);
                com.Parameters.AddWithValue("@Pin", Pin);
                com.Parameters.AddWithValue("@Mobileno", Mobileno);
                com.Parameters.AddWithValue("@Email", Email);
                com.Parameters.AddWithValue("@AadharNo", AadharNo);
                com.Parameters.AddWithValue("@CreatedBY", Session["EmpId"]);

                SqlParameter parameter1 = new SqlParameter();
                parameter1.ParameterName = "@EmployeeEducationDetail";
                parameter1.SqlDbType = System.Data.SqlDbType.Structured;
                parameter1.Value = table;
                com.Parameters.Add(parameter1);
                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);
                com.CommandTimeout = 0;

                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();
                com.ExecuteNonQuery();
                conObj.con.Close();

                mes = "Record saved successfully";
                mes = sp.Value.ToString();
                string retunvalue = (string)com.Parameters["@Mes"].Value;
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("POList", "Error: " + ex.Message);
            }
            return Json(mes);
        }

        public ActionResult _EducationDetailList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@DeptId", Session["DeptId"]?.ToString() ?? "0");
            dt = comfun.fillDataTable("HpsedcNewEmpApply_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _EducationList()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@EmpId", "4938");
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@DeptId", Session["DeptId"]?.ToString() ?? "0");
            dt = comfun.fillDataTable("HpsedcNewEmpApplyEducation_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }


        #endregion

        public ActionResult AvoMapping()
        {
            return View();
        }

        #region MeanModule Vishu

        public ActionResult HpsedcModule()
        {
            return View();
        }
        public JsonResult _ModuleList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@ModuleId", Request.Form["ModuleId"].ToString());
            dt = comfun.fillDataTable("HpsedcModule_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public JsonResult _ModuleEdit()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@ModuleId", Request.Form["ModuleId"].ToString());
            dt = comfun.fillDataTable("HpsedcModule_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _ModuleSubmit()
        {
            //comfun.saveformname("HiringExecutiveSubmit", "/HRNew/HiringExecutiveSubmit", "HRNew/HiringExecutive", "HiringExecutive  Add", "N", "HiringExecutive", "HiringExecutive", "Add", 2);
            //comfun.saveformname("HiringExecutiveSubmit", "/HRNew/HiringExecutiveSubmit", "HRNew/HiringExecutive", "HiringExecutive", "N", "HiringExecutive", "HiringExecutive", "Submit", 3);
            string mes = string.Empty;
            SortedList list = new SortedList();
            try


            {
                var Icon_image = "";
                var _comPath = "";
                var IconImage1 = "";
                var filePathA = "";
                HttpFileCollectionBase files = Request.Files;
                if (System.Web.HttpContext.Current.Request.Files.AllKeys.Any())
                {
                    // string mes2 = string.Empty;
                    for (int i = 0; i < files.Count; i++)
                    {
                        HttpPostedFileBase file = files[i];

                        // Checking for Internet Explorer  
                        if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                        {
                            string[] testfiles = file.FileName.Split(new char[] { '\\' });
                            IconImage1 = testfiles[testfiles.Length - 1];

                        }
                        else
                        {
                            IconImage1 = file.FileName;
                        }
                        var _ext = Path.GetExtension(IconImage1);
                        Icon_image = Path.GetFileNameWithoutExtension(IconImage1);
                        string filePath = Path.Combine(Server.MapPath("/DashboardLogo/") + Icon_image + _ext);
                        Icon_image = Icon_image + _ext;
                        IconImage1 = filePath;
                        filePathA = filePath;
                        file.SaveAs(filePathA);

                    }


                }
                list.Add("@ModuleId", Request.Form["ModuleId"].ToString());
                list.Add("@ModuleName", Request.Form["ModuleName"].ToString());
                list.Add("@MenuType", Request.Form["MenuType"].ToString());
                list.Add("@Description", Request.Form["Description"].ToString());
                list.Add("@link", Request.Form["link"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                list.Add("@Icon ", Icon_image);

                mes = comfun.executeNonQueryWMessage("HpsedcModule_AcceptUpdate", "", list).ToString();
            }

            catch (Exception ex)
            {

                mes = "Error:" + ex.Message;
            }

            return Json(mes);
        }


        #endregion

        #region MapModule Vishu

        public ActionResult MapModule()
        {
            return View();
        }
        public JsonResult _MapModuleSubmit()
        {

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@MapId", Request.Form["MapId"].ToString());
                list.Add("@ModuleId", Request.Form["ModuleId"].ToString());
                list.Add("@SubModuleId", Request.Form["SubModuleId"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());

                mes = comfun.executeNonQueryWMessage("HpsedcModuleMap_AcceptUpdate", "", list).ToString();
            }

            catch (Exception ex)
            {

                mes = "Error:" + ex.Message;
            }

            return Json(mes);
        }
        public JsonResult _MapModuleList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@MapId", Request.Form["MapId"].ToString());
            dt = comfun.fillDataTable("HpsedcModuleMap_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public JsonResult _MapModuleEdit()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@MapId", "0");
            dt = comfun.fillDataTable("HpsedcModuleMap_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _MapModuleddl()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId", Session["Customer"]);
            list.Add("@ModuleId", Request.Form["ModuleId"].ToString());
            list.Add("@MenuType", Request.Form["MenuType"].ToString());
            dt = comfun.fillDataTable("HpsedcModule_ddl", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }

        public ActionResult _SubModuleddl()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId", Session["Customer"]);
            list.Add("@ModuleId", Request.Form["ModuleId"].ToString());
            list.Add("@MenuType", Request.Form["MenuType"].ToString());
            dt = comfun.fillDataTable("HpsedcModule_ddl", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }

        #endregion

        #region RoleAccessModule Vishu
        public ActionResult RoleAccess()
        {
            return View();
        }

        public JsonResult _RoleAccessList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AccessId", Request.Form["AccessId"].ToString());
            dt = comfun.fillDataTable("HpsedcModuleAccess_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _RoleAccessSubmit()
        {
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@AccessId", "0");
                list.Add("@ModuleId", Request.Form["ModuleId"].ToString());
                list.Add("@AccessBy", Request.Form["AccessBy"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());

                mes = comfun.executeNonQueryWMessage("HpsedcModuleAccess_AcceptUpdate", "", list).ToString();
            }

            catch (Exception ex)
            {

                mes = "Error:" + ex.Message;
            }

            return Json(mes);
        }
        public JsonResult _RoleAccessEdit()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();

            list.Add("@AccessId", "0");
            dt = comfun.fillDataTable("HpsedcModuleAccess_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _MapAccessModuleddl()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@ModuleId", Request.Form["ModuleId"].ToString());
            list.Add("@MenuType", Request.Form["MenuType"].ToString());
            dt = comfun.fillDataTable("HpsedcModule_ddl", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }

        public ActionResult _Roleddl()
        {
            SortedList list = new SortedList();
            list.Add("@UserRole", Request.Form["UserRole"].ToString());
            DataTable dt = comfun.fillDataTable("HpsedcGetRole_List", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        #endregion


    }
}
