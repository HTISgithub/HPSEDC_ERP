using DocumentFormat.OpenXml.Bibliography;
using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Helpers;
using System.Web.Mvc;
using Payroll.portal.Models;
using System.Net;
using System.Text;
using System.Web.Script.Serialization;
using ClosedXML.Excel;
using Excelcon = Microsoft.Office.Interop.Excel;
using ExcelDataReader;


namespace Payroll.portal.Controllers
{
    //[Authorize]
    public class AdminController : Controller
    {
        #region Payroll Management System
        public const string _roles = payrollFunctions._rolesGlobal;
        commonFunctions comfun = new commonFunctions();
        payrollFunctions payfun = new payrollFunctions();

        /// <summary>
        /// GeoFence Master
        /// </summary>
        /// <returns></returns>
        #region GeoFence Master
        public ActionResult GeoFenceList()
        {
            comfun.saveformname("GeofenceMaster", "/admin/GeoFenceList", "Master/Geofence/Master&nbsp;", "Geofence Master form", "Y", "GeofenceMaster", "GeofenceMaster", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("GeofenceMaster");
            DataTable dt = comfun.fillDataTable("Country_Select", "", null);
            return View(dt);
        }
        public ActionResult _GeoFeneceList()
        {
            comfun.saveformname("_GeoFeneceList", "/admin/_GeoFeneceList", "Master/Geofence/Master&nbsp;", "Geofence Master List", "N", "GeofenceMaster", "GeofenceMaster", "List", 4);
            SortedList list = new SortedList();
            DataTable dt = comfun.fillDataTable("sp_GeofenceListDisplay", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult _GeofeneceAdd()
        {
            DataTable dt = comfun.fillDataTable("Country_Select", "", null);
            return PartialView("_GeoFeneceAdd", dt);
        }
        public ActionResult _StateListDisplay()
        {
            SortedList list = new SortedList();
            list.Add("@CountryId", Request.Form["CountryId"].ToString());
            DataTable dt = comfun.fillDataTable("State_SelectByCountryId", "", list);
            return PartialView("_StateListDisplay", dt);
        }
        public ActionResult _CityListDisplay()
        {
            SortedList list = new SortedList();
            list.Add("@fiStateId", Request.Form["StateId"].ToString());
            DataTable dt = comfun.fillDataTable("stpVIKSATCity_List", "", list);
            return PartialView("_CityListDisplay", dt);
        }
        public JsonResult _GeoFenceSaveUpdateSubmit()
        {
            comfun.saveformname("_GeoFenceSaveUpdateSubmit", "/admin/_GeoFenceSaveUpdateSubmit", "Master/Geofence/Master&nbsp;", "Geofence Master Add", "N", "GeofenceMaster", "GeofenceMaster", "Add", 2);
            comfun.saveformname("_GeoFenceSaveUpdateSubmit", "/admin/_GeoFenceSaveUpdateSubmit", "Master/Geofence/Master&nbsp;", "Geofence Master Edit", "N", "GeofenceMaster", "GeofenceMaster", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@GeoFenceId", Request.Form["GeoFenceId"].ToString());
                list.Add("@CountryId", Request.Form["GeofenceCountryId"].ToString());
                list.Add("@StateId", Request.Form["GeofenceStateId"].ToString());
                list.Add("@CityId", Request.Form["GeofenceCityId"].ToString());
                list.Add("@Location", Request.Form["GeofenceLocation"].ToString());
                list.Add("@Address", Request.Form["GeofenceAddress"].ToString());
                list.Add("@Latitude", Request.Form["GeofenceLatitude"].ToString());
                list.Add("@Longitude", Request.Form["GeofenceLongitude"].ToString());
                list.Add("@Radius", Request.Form["GeofenceRadius"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_Geofence_SaveUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        public ActionResult _GeoFeneceEdit()
        {
            string geofenceid = "0";
            geofenceid = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@GeoFenceId", geofenceid);
            DataSet ds = comfun.fillDataSet("sp_GeofenceDisplayViaId", "", list);
            var json1 = JsonConvert.SerializeObject(ds.Tables[0]);
            return Json(json1);
        }
        public JsonResult GeoFeneceStatusUpdate()
        {
            comfun.saveformname("GeoFeneceStatusUpdate", "/admin/GeoFeneceStatusUpdate", "Master/Geofence/Master&nbsp;", "Geofence Master Status", "N", "GeofenceMaster", "GeofenceMaster", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_Geofence_Active_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        #endregion GeoFence Master

        /// <summary>
        /// GeoFence Mapping
        /// </summary>
        /// <returns></returns>
        #region GeoFence Mapping
        public ActionResult GeoFenceMappingList()
        {
            comfun.saveformname("GeofenceMapping", "/admin/GeoFenceMappingList", "Master/Geofence/User Mapping", "GeoFence Mapping form", "Y", "GeofenceMapping", "GeofenceMapping", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("GeofenceMapping");
            DataSet ds = comfun.fillDataSet("sp_GeoFenceMapping_ddl", "", null);
            return View(ds);
        }
        public ActionResult _GeoFenceMappingList()
        {
            comfun.saveformname("_GeoFenceMappingList", "/admin/_GeoFenceMappingList", "Master/Geofence/User Mapping", "GeoFence Mapping form", "Y", "GeofenceMapping", "GeofenceMapping", "List", 4);
            DataTable dt = comfun.fillDataTable("sp_GeoFenceMappingList", "", null);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public JsonResult _GeoFenceMappingDisplayViaId()
        {
            SortedList list = new SortedList();
            list.Add("@GeoFenceMappingId", Request.Form["GeoFenceMappingId"].ToString());
            DataSet ds = comfun.fillDataSet("sp_GeoFenceMappingDetailsViaId", "", list);
            var json1 = JsonConvert.SerializeObject(ds.Tables[0]);
            return Json(json1);
        }
        public JsonResult _GeoFenceMappingListByEmp()
        {
            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            DataTable dt = comfun.fillDataTable("sp_GeoFenceMappingDetailByEmpId", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }

        public ActionResult _GeoFeneceUserMappingEdit()
        {
            SortedList list = new SortedList();
            list.Add("@GeoFenceMappingId", Request.Form["GeoFenceMappingId"].ToString());
            DataSet ds = comfun.fillDataSet("sp_GeoFenceMappingDetailsViaId", "", list);
            var json1 = JsonConvert.SerializeObject(ds.Tables[0]);
            return Json(json1);

        }
        public ActionResult _GeoFenceMapping()
        {
            DataSet ds = comfun.fillDataSet("sp_GeoFenceMapping_ddl", "", null);
            return PartialView("_GeoFenceMapping", ds);
        }
        public ActionResult _GeoFenceSitesStateList()
        {
            SortedList list = new SortedList();
            list.Add("@CountryId", Request.Form["CountryId"].ToString());
            DataTable dt = comfun.fillDataTable("sp_GeoFenceSitesSatesList", "", list);
            return PartialView("_GeoFenceSitesStateList", dt);
        }
        public ActionResult _GeoFenceSitesCityList()
        {
            SortedList list = new SortedList();
            list.Add("@StateId", Request.Form["StateId"].ToString());
            DataTable dt = comfun.fillDataTable("sp_GeoFenceSitesCityList", "", list);
            return PartialView("_GeoFenceSitesCityList", dt);
        }
        public class GeoFenceSites
        {
            public string GeoFenceId { get; set; }
            public string StateId { get; set; }
            public string StateName { get; set; }
            public string CityId { get; set; }
            public string CityName { get; set; }
            public string GeoFenceLocation { get; set; }
            public string GeoFenceAddress { get; set; }
            public string GeoFenceLatitude { get; set; }
            public string GeoFenceLongitude { get; set; }
            public string GeoFenceRadius { get; set; }

        }
        public ActionResult GeoFenceSitesListDisplay()
        {
            SortedList list = new SortedList();
            list.Add("@CityId", Request.Form["GeoFenceCityId"].ToString());
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            DataTable dt = comfun.fillDataTable("sp_GeoFenceSitesDisplayViaCityId", "", list);
            //List<GeoFenceSites> GFSitesList = new List<GeoFenceSites>();
            //GFSitesList = (from DataRow dr in dt.Rows
            //               select new GeoFenceSites()
            //               {
            //                   GeoFenceId = dr["GeoFenceId"].ToString(),
            //                   StateId = dr["StateId"].ToString(),
            //                   StateName = dr["StateName"].ToString(),
            //                   CityId = dr["CityId"].ToString(),
            //                   CityName = dr["CityName"].ToString(),
            //                   GeoFenceLocation = dr["GeoFenceLocation"].ToString(),
            //                   GeoFenceAddress = dr["GeoFenceAddress"].ToString(),
            //                   GeoFenceLatitude = dr["GeoFenceLatitude"].ToString(),
            //                   GeoFenceLongitude = dr["GeoFenceLongitude"].ToString(),
            //                   GeoFenceRadius = dr["GeoFenceRadius"].ToString(),
            //               }).ToList();
            // return Json(GFSitesList, JsonRequestBehavior.AllowGet);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult GeoFenceSitesListDisplayMapped()
        {
            SortedList list = new SortedList();
            list.Add("@CityId", Request.Form["GeoFenceCityId"].ToString());
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            DataTable dt = comfun.fillDataTable("sp_GeoFenceSitesDisplayViaCityIdMapped", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _GeoFenceMappedSites()
        {
            SortedList list = new SortedList();
            list.Add("@GeoFenceMappingId", Request.Form["GeoFenceMappingId"].ToString());
            DataTable dt = comfun.fillDataTable("sp_GeofenceEmpMappingListViaId", "", list);
            return PartialView("_GeoFenceMappedSites", dt);
        }
        public JsonResult _GeoFenceMappingSaveUpdateSubmit()
        {
            comfun.saveformname("_GeoFenceMappingSaveUpdateSubmit", "/admin/_GeoFenceMappingSaveUpdateSubmit", "Master/Geofence/User Mapping", "GeoFence Mapping Add", "N", "GeofenceMapping", "GeofenceMapping", "Add", 2);
            comfun.saveformname("_GeoFenceMappingSaveUpdateSubmit", "/admin/_GeoFenceMappingSaveUpdateSubmit", "Master/Geofence/User Mapping", "GeoFence Mapping Edit", "N", "GeofenceMapping", "GeofenceMapping", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                //list.Add("@GeoFenceMappingId", ""Request.Form["GeoFenceMappingId"].ToString());
                list.Add("@EmpId", Request.Form["EmpId"].ToString());
                list.Add("@GeoFenceId", Request.Form["GeoFenceId"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_GeoFenceMapping_SaveUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        #endregion GeoFence Mapping


        /// <summary>
        /// Clearance Group Master
        /// </summary>
        /// <returns></returns>
        #region Clearance Group Master
        public ActionResult CGList()
        {
            comfun.saveformname("ClearanceGroup", "/admin/CGList", "Master/Clearance/Group", "Clearance Group form", "Y", "ClearanceGroup", "ClearanceGroup", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("ClearanceGroup");
            return View();
        }

        public ActionResult _CGList()
        {
            comfun.saveformname("_CGList", "/admin/_CGList", "Master/Clearance/Group", "Clearance Group List", "Y", "ClearanceGroup", "ClearanceGroup", "List", 4);
            DataTable dt = comfun.fillDataTable("sp_CGListDisplay", "", null);
            //return PartialView("_CGList", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }


        public JsonResult CGSubmit()
        {
            comfun.saveformname("CGSubmit", "/admin/CGSubmit", "Master/Clearance/Group", "Clearance Group Add", "N", "ClearanceGroup", "ClearanceGroup", "Add", 2);
            comfun.saveformname("CGSubmit", "/admin/CGSubmit", "Master/Clearance/Group", "Clearance Group Edit", "N", "ClearanceGroup", "ClearanceGroup", "Edit", 3);
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
        public JsonResult CGAStatusUpdate()
        {
            comfun.saveformname("CGAStatusUpdate", "/admin/CGAStatusUpdate", "Master/Clearance/Group", "Clearance Group Status", "N", "ClearanceGroup", "ClearanceGroup", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_CG_Active_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        #endregion Clearance Group Master

        /// <summary>
        /// Group and Clearance Head Mapping
        /// </summary>
        /// <returns></returns>
        #region Group and Clearance Head Mapping
        public ActionResult GroupandClearanceHeadMapping()
        {
            comfun.saveformname("ClearanceGroupHeadMapping", "/admin/GroupandClearanceHeadMapping", "Master/Clearance/Mapping", "Clearance Group Head Mapping form", "Y", "ClearanceGroupHeadMapping", "ClearanceGroupHeadMapping", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("ClearanceGroupHeadMapping");
            DataTable dt = comfun.fillDataTable("stpLevellist", "", null);

            return View(dt);
        }
        public JsonResult _Clearancegroupmappingstatusupdate()
        {
            // comfun.saveformname("_LoanTypeActivationStatusUpdate", "/admin/_LoanTypeActivationStatusUpdate", "Payroll Management/Master/Loan Type", "Loan Type Status", "N", "LoanType", "LoanType", "Status", 5);
            comfun.saveformname("_Clearancegroupmappingstatusupdate", "/admin/_Clearancegroupmappingstatusupdate", "Master/Clearance/Mapping", "Clearance Group Head Mapping form", "N", "ClearanceGroupHeadMapping", "ClearanceGroupHeadMapping", "Status", 5);

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
        public ActionResult _GroupandClearanceHeadMapping()
        {
            comfun.saveformname("_GroupandClearanceHeadMapping", "/admin/_GroupandClearanceHeadMapping", "Master/Clearance/Mapping", "Clearance Group Head Mapping List", "N", "ClearanceGroupHeadMapping", "ClearanceGroupHeadMapping", "List", 4);
            DataTable dt = comfun.fillDataTable("sp_CGListDisplay", "", null);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult _UpdateGroupandClearanceHeadMapping()
        {
            SortedList list = new SortedList();
            list.Add("@CGId", Request.Form["CGId"].ToString());
            DataSet ds = comfun.fillDataSet("sp_CGDisplayViaId", "", list);
            return PartialView("_UpdateGroupandClearanceHeadMapping", ds);
        }
        public ActionResult _CGMappedEmployees()
        {
            SortedList list = new SortedList();
            list.Add("@CGId", Request.Form["CGId"].ToString());
            //list.Add("@LevelId", Request.Form["LevelId"].ToString());
            DataTable dt = comfun.fillDataTable("sp_ClearanceGroupMappedEmpList", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
            //return PartialView("_CGMappedEmployees", dt);
        }
        public class Employees
        {
            public string EmployeeId { get; set; }
            public string EmployeeCode { get; set; }
            public string EmployeeName { get; set; }
            public string EmployeeEmail { get; set; }
            public string EmployeeImage { get; set; }
            public string EmployeeDepartment { get; set; }
            public string EmployeeDesignation { get; set; }
            public string EmployeeBranch { get; set; }

        }
        public ActionResult _EmpListDepartmentWise()
        {
            SortedList list = new SortedList();
            list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
            list.Add("@MainCategoryId", Request.Form["MainCategoryId"].ToString());
            list.Add("@CategoryId", Request.Form["CategoryId"].ToString());

            if (Request.Form["GroupId"] != null)
            {
                list.Add("@GroupId", Request.Form["GroupId"].ToString());
            }
            DataTable dt = comfun.fillDataTable("sp_CGroupEmpListDepartmentId", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _Levelbox()
        {

            DataTable dt = comfun.fillDataTable("stpLevellist", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }


        public ActionResult _EmpListDisplay()
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
                           EmployeeName = dr["EmployeeName"].ToString(),
                           EmployeeEmail = dr["Email"].ToString(),
                           EmployeeImage = dr["EmployeeImage"].ToString(),
                           EmployeeDepartment = dr["EmployeeDepartment"].ToString(),
                           EmployeeDesignation = dr["EmployeeDesignation"].ToString(),
                           EmployeeBranch = dr["EmployeeBranch"].ToString(),
                       }).ToList();
            return Json(empList, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GroupandClearanceHeadMappingSaveUpdateSubmit()
        {
            comfun.saveformname("GroupandClearanceHeadMappingSaveUpdateSubmit", "/admin/GroupandClearanceHeadMappingSaveUpdateSubmit", "Master/Clearance/Mapping", "Clearance Group Head Edit", "N", "ClearanceGroupHeadMapping", "ClearanceGroupHeadMapping", "Add", 2);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();

                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject(Request.Form["Data"].ToString());
                list.Add("@GroupId", Request.Form["GroupId"].ToString());
                comfun.executeNonQuery("sp_ClearanceGroupMapping_SetNoMapping", "", list);
                foreach (var item in jsonData)
                {
                    list.Clear();
                    list.Add("@CGId", item.GroupId);
                    list.Add("@LevelId", item.LevelId);
                    list.Add("@EmpId", item.EmpId);
                    list.Add("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                    list.Add("@CreatedBy", "1");
                    mes = comfun.executeNonQueryWMessage("sp_ClearanceGroupMapping_SaveUpdate", "", list).ToString();
                }
            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
            }
            return Json(mes);
        }
        #endregion Group and Clearance Head Mapping


        /// <summary>
        /// Device Connection Type Master
        /// </summary>
        /// <returns></returns>
        #region Device Connection Type Master
        public ActionResult DeviceConnectionTypeList()
        {
            comfun.saveformname("DeviceConnectionType", "/admin/DeviceConnectionTypeList", "Device Management/Device Connection Type", "Device Connection Type form", "Y", "DeviceConnectionType", "DeviceConnectionType", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("DeviceConnectionType");
            return View();
        }
        public ActionResult _DeviceConnectionTypeList()
        {
            comfun.saveformname("_DeviceConnectionTypeList", "/admin/_DeviceConnectionTypeList", "Device Management/Device Connection Type", "Device Connection Type List", "N", "DeviceConnectionType", "DeviceConnectionType", "List", 4);

            DataTable dt = comfun.fillDataTable("sp_DeviceConnectionTypeListDisplay", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _DeviceConnectionTypeAdd()
        {
            return PartialView("_DeviceConnectionTypeAdd");
        }
        public ActionResult _DeviceConnectionTypeEdit()
        {
            SortedList list = new SortedList();
            list.Add("@DeviceConnectionTypeId", Request.Form["DeviceConnectionTypeId"].ToString());
            DataTable dt = comfun.fillDataTable("sp_DeviceConnectionTypeDisplayViaId", "", list);
            return PartialView("_DeviceConnectionTypeEdit", dt);
        }
        public JsonResult DeviceConnectionTypeSaveUpdateSubmit()
        {
            comfun.saveformname("DeviceConnectionTypeSaveUpdateSubmit", "/admin/DeviceConnectionTypeSaveUpdateSubmit", "Device Management/Device Connection Type", "Device Connection Type Add", "N", "DeviceConnectionType", "DeviceConnectionType", "Add", 2);
            comfun.saveformname("DeviceConnectionTypeSaveUpdateSubmit", "/admin/DeviceConnectionTypeSaveUpdateSubmit", "Device Management/Device Connection Type", "Device Connection Type Edit", "N", "DeviceConnectionType", "DeviceConnectionType", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@DeviceConnectionTypeId", Request.Form["DeviceConnectionTypeId"].ToString());
                list.Add("@DeviceConnectionTypeName", Request.Form["DeviceConnectionTypeName"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_DeviceConnectionType_Save_Update", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        public JsonResult DeviceConnectionTypeStatusUpdate()
        {
            comfun.saveformname("DeviceConnectionTypeStatusUpdate", "/admin/DeviceConnectionTypeStatusUpdate", "Device Management/Device Connection Type", "Device Connection Type Status", "N", "DeviceConnectionType", "DeviceConnectionType", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@DeviceConnectionTypeId", Request.Form["Id"].ToString());
                list.Add("@DeviceConnectionTypeStatus", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_DeviceConnectionType_Active_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        #endregion Device Connection Type Master


        /// <summary>
        /// Device Model Number Master
        /// </summary>
        /// <returns></returns>
        #region Device Model Number Master
        public ActionResult DeviceModelNumberList()
        {
            comfun.saveformname("DeviceModelNumber", "/admin/DeviceModelNumberList", "Device Management/Model Number", "Device Model Number form", "Y", "DeviceModelNumber", "DeviceModelNumber", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("DeviceModelNumber");
            return View();
        }
        public ActionResult _DeviceModelNumberList()
        {
            comfun.saveformname("_DeviceModelNumberList", "/admin/_DeviceModelNumberList", "Device Management/Model Number", "Device Model Number List", "N", "DeviceModelNumber", "DeviceModelNumber", "List", 4);

            DataTable dt = comfun.fillDataTable("sp_DeviceModelNoListDisplay", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _DeviceModelNumberAdd()
        {
            return PartialView("_DeviceModelNumberAdd");
        }
        public ActionResult _DeviceModelNumberEdit()
        {
            SortedList list = new SortedList();
            list.Add("@DeviceModelNoId", Request.Form["DeviceModelNoId"].ToString());
            DataTable dt = comfun.fillDataTable("sp_DeviceModelNoDisplayViaId", "", list);
            return PartialView("_DeviceModelNumberEdit", dt);
        }
        public JsonResult DeviceModelNumberSaveUpdateSubmit()
        {
            comfun.saveformname("DeviceModelNumberSaveUpdateSubmit", "/admin/DeviceModelNumberSaveUpdateSubmit", "Device Management/Model Number", "Device Model Number Add", "N", "DeviceModelNumber", "DeviceModelNumber", "Add", 2);
            comfun.saveformname("DeviceModelNumberSaveUpdateSubmit", "/admin/DeviceModelNumberSaveUpdateSubmit", "Device Management/Model Number", "Device Model Number Edit", "N", "DeviceModelNumber", "DeviceModelNumber", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@DeviceModelNoId", Request.Form["DeviceModelNoId"].ToString());
                list.Add("@DeviceModelNo", Request.Form["DeviceModelNo"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_DeviceModelNo_SaveUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        public JsonResult DeviceModelNumberStatusUpdate()
        {
            comfun.saveformname("DeviceModelNumberSaveUpdateSubmit", "/admin/DeviceModelNumberSaveUpdateSubmit", "Device Management/Model Number", "Device Model Number Status", "N", "DeviceModelNumber", "DeviceModelNumber", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@DeviceModelNoId", Request.Form["Id"].ToString());
                list.Add("@DeviceModelNoStatus", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_DeviceModelNo_Active_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        #endregion Device Model Number Master


        /// <summary>
        /// Device Details Master
        /// </summary>
        /// <returns></returns>
        #region Device Details Master
        public ActionResult DeviceList()
        {
            comfun.saveformname("DeviceList", "/admin/DeviceList", "Device Management/Device", "DeviceList form", "Y", "DeviceList", "DeviceList", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("DeviceList");
            DataSet ds = comfun.fillDataSet("sp_DeviceDetails_ddl", "", null);
            return View(ds);
        }


        private double ToUnixTime(DateTime input)
        {
            return Math.Floor(input.Subtract(new DateTime(1970, 1, 1)).TotalSeconds);
        }

        public ActionResult _DeviceJsonList()
        {
            comfun.saveformname("_DeviceJsonList", "/admin/_DeviceJsonList", "Device Management/Device", "Device form list", "N", "DeviceList", "DeviceList", "List", 4);

            DataTable dt = comfun.fillDataTable("sp_DeviceListDisplay", "", null);
            string deviceCode = "", baseUrl = "", qry = "";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                deviceCode = dt.Rows[i]["DeviceSerialNumber"].ToString();
                var timestamp1 = ToUnixTime(Convert.ToDateTime(DateTime.UtcNow.AddMinutes(325).ToString("dd-MMM-yyyy HH:mm:ss")));
                var timestamp2 = ToUnixTime(Convert.ToDateTime(DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss")));

                baseUrl = "http://" + dt.Rows[i]["IpAddress"] + ":" + dt.Rows[i]["PortNo"] + "/LAPI/V1.0/PACS/Controller/PassRecord";
                qry = "{\"Num\":\"2\",\"QueryInfoList\":[{\"QryType\":\"4\",\"QryCondition\":\"3\",\"QryData\":\"timestamp1\"},{\"QryType\":\"4\",\"QryCondition\":\"4\",\"QryData\":\"timestamp2\"}],\"Limit\":\"150\",\"Offset\":\"0\"}";

                qry = qry.Replace("timestamp1", timestamp1.ToString());
                qry = qry.Replace("timestamp2", timestamp2.ToString());

                WebClient client = new WebClient();
                client.Headers["Content-type"] = "application/json";
                client.Encoding = Encoding.UTF8;

                try
                {
                    var jsonlog = client.UploadString(baseUrl, qry);
                    RootLog log = (new JavaScriptSerializer()).Deserialize<RootLog>(jsonlog);
                    if (log.Response.Data != null)
                    {
                        DataRow row = dt.Select("DeviceSerialNumber='" + deviceCode + "'").FirstOrDefault();
                        row["McStatus"] = "Online";
                    }

                }
                catch (Exception ex)
                {

                }
                finally
                {
                }

            }

            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _DeviceAdd()
        {
            DataSet ds = comfun.fillDataSet("sp_DeviceDetails_ddl", "", null);
            return PartialView("_DeviceAdd", ds);
        }

        public ActionResult _DeviceEdit()
        {
            SortedList list = new SortedList();
            list.Add("@DeviceId", Request.Form["DeviceId"].ToString());
            DataSet ds = comfun.fillDataSet("sp_DeviceDisplayViaId", "", list);
            var json = JsonConvert.SerializeObject(ds.Tables[0]);
            return Json(json);
        }

        public JsonResult DeviceDetailsSaveUpdateSubmit()
        {
            comfun.saveformname("DeviceDetailsSaveUpdateSubmit", "/admin/DeviceDetailsSaveUpdateSubmit", "Device Management/Device", "Device form Add", "N", "DeviceList", "DeviceList", "Add", 2);
            comfun.saveformname("DeviceDetailsSaveUpdateSubmit", "/admin/DeviceDetailsSaveUpdateSubmit", "Device Management/Device", "Device form Edit", "N", "DeviceList", "DeviceList", "Edit", 3);
            string mes = string.Empty;
            try
            {
                string Invoice_image = "";
                if (System.Web.HttpContext.Current.Request.Files.AllKeys.Any())
                {
                    var DeviceInvoiceImage = System.Web.HttpContext.Current.Request.Files["InvoiceImage"];
                    if (DeviceInvoiceImage.ContentLength > 0)
                    {
                        var fileName = Path.GetFileName(DeviceInvoiceImage.FileName);
                        var _ext = Path.GetExtension(DeviceInvoiceImage.FileName);
                        Invoice_image = Guid.NewGuid().ToString();
                        var _comPath = Server.MapPath("/DeviceInvoices/") + Invoice_image + _ext;
                        Invoice_image = Invoice_image + _ext;

                        ViewBag.Msg = _comPath;
                        var path = _comPath;

                        // Saving Image in Original Mode
                        DeviceInvoiceImage.SaveAs(path);
                        // resizing image
                        MemoryStream ms = new MemoryStream();
                        WebImage img = new WebImage(_comPath);

                        if (img.Width > 800)
                            img.Resize(800, 800);
                        img.Save(_comPath);
                        // end resize
                    }
                }
                SortedList list = new SortedList();
                list.Add("@DeviceId", Request.Form["DeviceId"].ToString());
                list.Add("@DeviceName", Request.Form["DeviceName"].ToString());
                list.Add("@DeviceSerialNumber", Request.Form["DeviceSerialNo"].ToString());
                list.Add("@DeviceLocation", Request.Form["DeviceLocation"].ToString());
                list.Add("@DeviceNumber", Request.Form["DeviceNo"].ToString());
                list.Add("@DeviceConnectionTypeId", Request.Form["DeviceConnectionTypeId"].ToString());
                list.Add("@DeviceModelNumberId", Request.Form["DeviceModelNoId"].ToString());
                list.Add("@DeviceBranchId", Request.Form["DeviceBranchId"].ToString());
                list.Add("@DeviceInchargeId", Request.Form["DeviceInchargeId"].ToString());
                list.Add("@DeviceShortName", Request.Form["DeviceShortName"].ToString());
                list.Add("@DeviceModelName", Request.Form["DeviceModelName"].ToString());
                list.Add("@DevicePassword", Request.Form["DevicePassword"].ToString());
                list.Add("@DevicePasswordSize", Request.Form["DevicePasswordSize"].ToString());
                list.Add("@DeviceUserSize", Request.Form["DeviceUserSize"].ToString());
                list.Add("@DeviceCardSize", Request.Form["DeviceCardSize"].ToString());
                list.Add("@DeviceUsedUser", Request.Form["DeviceUsedUser"].ToString());
                list.Add("@DeviceUsedCard", Request.Form["DeviceUsedCard"].ToString());
                list.Add("@DeviceUsedNewLog", Request.Form["DeviceUsedNewLog"].ToString());
                list.Add("@DeviceTime", Request.Form["DeviceTime"].ToString());
                list.Add("@DeviceVolume", Request.Form["DeviceVolume"].ToString());
                list.Add("@DeviceVerifyModeId", Request.Form["DeviceVerifyModeId"].ToString());
                list.Add("@DeviceUserFingerPrintNo", Request.Form["DeviceUserFingerPrintNo"].ToString());
                list.Add("@DeviceReVerifyTime", Request.Form["DeviceReVerifyTime"].ToString());
                list.Add("@DeviceIPAddress", Request.Form["DeviceIPAddress"].ToString());
                list.Add("@DevicePort", Request.Form["DevicePort"].ToString());
                list.Add("@DeviceUsedPassword", Request.Form["DeviceUsedPassword"].ToString());
                list.Add("@DeviceLicence", Request.Form["DeviceLicence"].ToString());
                list.Add("@DeviceFingerPrintSize", Request.Form["DeviceFingerPrintSize"].ToString());
                list.Add("@DeviceLogSize", Request.Form["DeviceLogSize"].ToString());
                list.Add("@DeviceUsedFingerPrint", Request.Form["DeviceUsedFingerPrint"].ToString());
                list.Add("@DeviceUsedLog", Request.Form["DeviceUsedLog"].ToString());
                list.Add("@DeviceFirmware", Request.Form["DeviceFirmware"].ToString());
                list.Add("@DeviceLanguageId", Request.Form["DeviceLangauageId"].ToString());
                list.Add("@DeviceScreenSaver", Request.Form["DeviceScreenSaver"].ToString());
                list.Add("@DeviceSleep", Request.Form["DeviceSleep"].ToString());
                list.Add("@DeviceLogHint", Request.Form["DeviceLogHint"].ToString());
                list.Add("@DeviceInvoice", Invoice_image);
                list.Add("@DeviceInvoicePath", Invoice_image);
                mes = comfun.executeNonQueryWMessage("sp_DeviceDetails_SaveUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        public JsonResult DeviceStatusUpdate()
        {
            comfun.saveformname("DeviceStatusUpdate", "/admin/DeviceStatusUpdate", "Device Management/Device", "Device form Status", "N", "DeviceList", "DeviceList", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@DeviceId", Request.Form["Id"].ToString());
                list.Add("@DeviceStatus", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_Device_Active_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        #endregion Device Details Master


        /// <summary>
        /// Tier Master
        /// </summary>
        /// <returns></returns>
        #region Tier Master
        public ActionResult TierList()
        {
            comfun.saveformname("Tier", "/admin/TierList", "Master/Address/Tier", "Tier  Main form", "Y", "Tier", "Tier", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("Tier");
            //DataTable dt = comfun.fillDataTable("sp_CityListForTier", "", null);
            //return View(dt);
            DataTable dt = comfun.fillDataTable("CountryState_Select", "", null);
            return View(dt);
        }

        public ActionResult _TierCityList()
        {
            SortedList list = new SortedList();
            list.Add("@StateId", Request.Form["StateId"].ToString());
            list.Add("@Action", Request.Form["Action"].ToString());
            DataTable dt = comfun.fillDataTable("sp_CityListForTier", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _Tierlist()
        {
            comfun.saveformname("_Tierlist", "/admin/_Tierlist", "Master/Address/Tier", "Tier List", "N", "Tier", "Tier", "List", 4);
            //decimal total_records = 0;
            //decimal pageno = 1;
            //decimal pagesize = 10;
            //if (Request.Form["PageNo"] != null)
            //{
            //    pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            //}
            SortedList list = new SortedList();
            //list.Add("@PageNo", pageno);
            //list.Add("@PageSize", pagesize);
            DataTable dt = comfun.fillDataTable("sp_TierListDisplay", "", list);
            //if (dt.Rows.Count > 0)
            //    total_records = Convert.ToDecimal(dt.Rows[0]["total"]);
            //string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "admin", "_TierList", "_Tierlist");
            //HtmlString htm = new HtmlString(paging);
            //ViewData["paging"] = htm;
            //return PartialView(dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _TierAdd()
        {
            DataTable dt = comfun.fillDataTable("sp_CityListForTier", "", null);
            return PartialView("_TierAdd", dt);
        }
        public JsonResult _TierSaveUpdateSubmit()
        {
            comfun.saveformname("_TierSaveUpdateSubmit", "/admin/_TierSaveUpdateSubmit", "Master/Address/Tier", "Tier Add", "N", "Tier", "Tier", "Add", 2);
            comfun.saveformname("_TierSaveUpdateSubmit", "/admin/_TierSaveUpdateSubmit", "Master/Address/Tier", "Tier Edit", "N", "Tier", "Tier", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@TierId", Request.Form["TierId"].ToString());
                list.Add("@TierCode", Request.Form["TierCode"].ToString());
                list.Add("@TierName", Request.Form["TierName"].ToString());
                list.Add("@TierCities", Request.Form["TierCities"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_Tier_SaveUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_TierAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        public JsonResult _TierStatusUpdate()
        {
            comfun.saveformname("_TierStatusUpdate", "/admin/_TierStatusUpdate", "Master/Address/Tier", "Tier Status", "N", "Tier", "Tier", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_Tier_Active_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }
        public ActionResult _TierEdit()
        {
            SortedList list = new SortedList();
            list.Add("@TierId", Request.Form["Id"].ToString());
            DataTable dt = comfun.fillDataTable("sp_TierListDisplayViaId1", "", list);
            //return PartialView("_TierEdit", ds);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public class CityList
        {
            public string CityId { get; set; }
            public string CityName { get; set; }
        }

        public ActionResult TierCityList()
        {
            SortedList list = new SortedList();
            DataTable dt = comfun.fillDataTable("sp_CityListForTier", "", list);
            List<CityList> cityList = new List<CityList>();
            cityList = (from DataRow dr in dt.Rows
                        select new CityList()
                        {
                            CityId = dr["CityId"].ToString(),
                            CityName = dr["CityName"].ToString(),
                        }).ToList();
            return Json(cityList, JsonRequestBehavior.AllowGet);
        }

        #endregion Tier Master


        /// <summary>
        /// Employee Category 
        /// </summary>
        /// <returns></returns>
        #region Employee Category
        public ActionResult EmployeeCategoryList()
        {
            comfun.saveformname("EmployeeCategory", "/admin/EmployeeCategoryList", "Master/General/Category", "Employee Category form", "Y", "EmployeeCategory", "EmployeeCategory", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("EmployeeCategory");

            return View();
        }
        public JsonResult _EmployeeCategoryList()
        {
            comfun.saveformname("_EmployeeCategoryList", "/admin/_EmployeeCategoryList", "Master/General/Category", "Employee Category List", "N", "EmployeeCategory", "EmployeeCategory", "List", 4);

            DataTable dt = comfun.fillDataTable("sp_EmpCategoryListDisplay", "", null);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1); ;
        }
        public ActionResult EmployeeCategoryAdd()
        {
            return PartialView("EmployeeCategoryAdd");
        }
        public JsonResult EmployeeCategorySaveUpdateSubmit()
        {
            comfun.saveformname("EmployeeCategorySaveUpdateSubmit", "/admin/EmployeeCategorySaveUpdateSubmit", "Master/General/Category", "Employee Category  Add", "N", "EmployeeCategory", "EmployeeCategory", "Add", 2);
            comfun.saveformname("EmployeeCategorySaveUpdateSubmit", "/admin/EmployeeCategorySaveUpdateSubmit", "Master/General/Category", "Employee Category  Edit", "N", "EmployeeCategory", "EmployeeCategory", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@EmpCategoryId", Request.Form["EmpCategoryId"].ToString());
                list.Add("@EmpCategoryCode", Request.Form["EmpCategoryCode"].ToString());
                list.Add("@EmpCategoryName", Request.Form["EmpCategoryName"].ToString());
                list.Add("@IsOverTime", Request.Form["IsOverTime"].ToString());
                list.Add("@IsLateComingApplied", Request.Form["IsLateComingApplied"].ToString());
                list.Add("@LCHalfDayAbsentMins", Request.Form["LCHalfDayAbsentMins"].ToString());
                list.Add("@LCFullDayAbsentMins", Request.Form["LCFullDayAbsentMins"].ToString());
                list.Add("@IsContinuousLate", Request.Form["IsContinuousLate"].ToString());
                list.Add("@IsWorkDurationApplied", Request.Form["IsWorkDurationApplied"].ToString());
                list.Add("@WDHalfDayAbsentMins", Request.Form["WDHalfDayAbsentMins"].ToString());
                list.Add("@WDFulldayAbsentMins", Request.Form["WDFulldayAbsentMins"].ToString());
                list.Add("@IsOutsider", Request.Form["IsOutsider"].ToString());
                list.Add("@calonWorking", Request.Form["calonWorking"].ToString());

                mes = comfun.executeNonQueryWMessage("sp_EmpCategory_SaveUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        public JsonResult EmployeeCategoryEdit()
        {
            SortedList list = new SortedList();
            list.Add("@EmpCategoryId", Request.Form["Id"].ToString());
            DataTable dt = comfun.fillDataTable("sp_EmpCategoryDisplayViaId", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public JsonResult EmployeeCategoryStatus()
        {
            comfun.saveformname("EmployeeCategoryStatus", "/admin/EmployeeCategoryStatus", "Master/General/Category", "Employee Category  Status", "N", "EmployeeCategory", "EmployeeCategory", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@EmpCategoryId", Request.Form["Id"].ToString());
                list.Add("@EmpCategoryStatus", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_EmpCategory_Active_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        #endregion Employee Category


        /// <summary>
        /// Notice Board
        /// </summary>
        /// <returns></returns>
        #region Notice Board
        public ActionResult NoticeBoardList()
        {
            comfun.saveformname("NoticeBoard", "/admin/NoticeBoardList", "Master/Advance/Notice Board", "Notice Board form", "Y", "NoticeBoard", "NoticeBoard", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("NoticeBoard");
            DataSet ds = comfun.fillDataSet("sp_DeviceDetails_ddl", "", null);
            return View(ds);
        }
        public ActionResult _NoticeBoardJsonList()
        {
            comfun.saveformname("_NoticeBoardJsonList", "/admin/_NoticeBoardJsonList", "Master/Advance/Notice Board", "Notice Board List", "N", "NoticeBoard", "NoticeBoard", "List", 4);
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
            DataTable dt = comfun.fillDataTable("sp_NoticeBoardList", "", null);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        //public ActionResult _NoticeBoardList()
        //{
        //    DataTable dt = comfun.fillDataTable("sp_NoticeBoardList", "", null);
        //    return PartialView("_NoticeBoardList", dt);
        //}
        public ActionResult _NoticeBoardAdd()
        {
            DataSet ds = comfun.fillDataSet("sp_DeviceDetails_ddl", "", null);
            return PartialView("_NoticeBoardAdd", ds);
        }
        public ActionResult _NoticeBoardEdit()
        {
            SortedList list = new SortedList();
            list.Add("@NoticeBoardId", Request.Form["NoticeBoardId"].ToString());
            DataTable dt = comfun.fillDataTable("sp_NoticeBoardDisplayViaId", "", list);
            //return PartialView("_NoticeBoardEdit", ds);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public JsonResult _NoticeBoardSubmit()
        {
            comfun.saveformname("_NoticeBoardSubmit", "/admin/_NoticeBoardSubmit", "Master/Advance/Notice Board", "Notice Board Add", "N", "NoticeBoard", "NoticeBoard", "Add", 2);
            comfun.saveformname("_NoticeBoardSubmit", "/admin/_NoticeBoardSubmit", "Master/Advance/Notice Board", "Notice Board Edit", "N", "NoticeBoard", "NoticeBoard", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@NoticeBoardId", Request.Form["NoticeBoardId"].ToString());
                list.Add("@NoticeBoardTitle", Request.Form["NoticeBoardTitle"].ToString());
                list.Add("@NoticeBoardBranchId", Request.Form["NoticeBoardBranchId"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                var NBMesaage = Request.Unvalidated["NoticeBoardMessage"];
                list.Add("@NoticeBoardMessage", NBMesaage);
                mes = comfun.executeNonQueryWMessage("sp_NoticeBoard_SaveUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        public JsonResult NoticeBoardStatus()
        {
            comfun.saveformname("NoticeBoardStatus", "/admin/NoticeBoardStatus", "Master/Advance/Notice Board", "Notice Board Add", "N", "NoticeBoard", "NoticeBoard", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@NoticeBoardId", Request.Form["Id"].ToString());
                list.Add("@NoticeBoardStatus", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_NoticeBoardId_Active_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        #endregion


        /// <summary>
        /// Helpdesk Group
        /// </summary>
        /// <returns></returns>
        #region HelpDesk Group
        public ActionResult HelpdeskGroupList()
        {
            comfun.saveformname("HelpDeskGroup", "/admin/HelpdeskGroupList", "Master/Help Desk/Group&nbsp;", "HelpDeskGroup form", "Y", "HelpDeskGroup", "HelpDeskGroup", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("HelpDeskGroup");
            return View();
        }
        public ActionResult _helpdeskgroups()
        {
            comfun.saveformname("_helpdeskgroups", "/admin/_helpdeskgroups", "Master/Help Desk/Group&nbsp;", "HelpDeskGroup List", "N", "HelpDeskGroup", "HelpDeskGroup", "List", 4);
            SortedList list = new SortedList();
            DataTable dt = comfun.fillDataTable("sp_helpdeskGroup_Select", "", list);
            //return PartialView("_helpdeskgroups", dt);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult _helpdeskgroupAdd()
        {
            return PartialView("_helpdeskgroupAdd");
        }


        public JsonResult HelpdeskGroupSubmit()
        {
            comfun.saveformname("HelpdeskGroupSubmit", "/admin/HelpdeskGroupSubmit", "Master/Help Desk/Group&nbsp;", "HelpDeskGroup Add", "N", "HelpDeskGroup", "HelpDeskGroup", "Add", 2);
            comfun.saveformname("HelpdeskGroupSubmit", "/admin/HelpdeskGroupSubmit", "Master/Help Desk/Group&nbsp;", "HelpDeskGroup Edit", "N", "HelpDeskGroup", "HelpDeskGroup", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@HelpdeskGroupId", Request.Form["GroupId"].ToString());
                list.Add("@HelpdeskGroupCode", Request.Form["GroupCode"].ToString());
                list.Add("@HelpdeskGroupName", Request.Form["GroupName"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_HelpDeskGroup_Save_Update", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_helpdeskgroupSubmit", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult _helpdeskgroupEdit()
        {
            SortedList list = new SortedList();
            list.Add("@GroupId", Request.Form["GroupId"].ToString());
            DataTable dt = comfun.fillDataTable("sp_helpdeskGroup_SelectByID", "", list);
            return PartialView("_helpdeskgroupEdit", dt);
        }
        public JsonResult _helpdeskgroupStatusUpdate()
        {
            comfun.saveformname("_helpdeskgroupStatusUpdate", "/admin/_helpdeskgroupStatusUpdate", "Master/Help Desk/Group&nbsp;", "HelpDeskGroup Status", "N", "HelpDeskGroup", "HelpDeskGroup", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@ID", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stphepdeskStatusUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }
        #endregion HelpDesk Group


        /// <summary>
        /// Helpdesk Group Mapping
        /// </summary>
        /// <returns></returns>
        #region Helpdesk Group Mapping
        public ActionResult HelpdeskGroupMappingList()
        {
            comfun.saveformname("HelpDeskMapping", "/admin/HelpdeskGroupMappingList", "Master/Help Desk/Mapping&nbsp;&nbsp;", "HelpDesk Mapping form", "Y", "HelpDeskMapping", "HelpDeskMapping", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("HelpDeskMapping");
            SortedList list = new SortedList();
            list.Add("@PageNo", 1);
            list.Add("@PageSize", 1000);
            DataTable dt = comfun.fillDataTable("Department_Select", "", list);
            return View(dt);
        }
        public JsonResult _HelpdeskGroupMappingList()
        {
            comfun.saveformname("_HelpdeskGroupMappingList", "/admin/_HelpdeskGroupMappingList", "Master/Help Desk/Mapping&nbsp;&nbsp;", "HelpDesk Mapping List", "N", "HelpDeskMapping", "HelpDeskMapping", "List", 4);
            SortedList list = new SortedList();
            DataTable dt = comfun.fillDataTable("sp_helpdeskGroup_Select", "", list);
            //return PartialView("_HelpdeskGroupMappingList", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _helpdeskGroupMappingBranch()
        {
            SortedList list = new SortedList();
            list.Add("@PageNo", 1);
            list.Add("@PageSize", 1000);
            DataTable dt = comfun.fillDataTable("Branch_Select", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _helpDeskDesignation()
        {
            SortedList list = new SortedList();
            list.Add("@PageNo", 1);
            list.Add("@PageSize", 1000);
            DataTable dt = comfun.fillDataTable("Designation_Select", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _helpDeskMappingEmployee()
        {
            SortedList list = new SortedList();
            list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
            list.Add("@BranchId", Request.Form["BranchId"].ToString());
            list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
            list.Add("@CategoryID", Request.Form["CategoryId"].ToString());
            list.Add("@MainCategory", Request.Form["MainCategory"].ToString());
            DataTable dt = comfun.fillDataTable("sp_HelpdeskGroupMappingEmpList", "", list);
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

        public JsonResult _HelpdeskMappingSaveUpdateSubmit()
        {
            comfun.saveformname("_HelpdeskMappingSaveUpdateSubmit", "/admin/_HelpdeskMappingSaveUpdateSubmit", "Master/Help Desk/Mapping&nbsp;&nbsp;", "HelpDesk Mapping Add", "N", "HelpDeskMapping", "HelpDeskMapping", "Add", 2);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@HelpdeskGroupId", Request.Form["HelpdeskGroupId"].ToString());
                list.Add("@MappedEmpId", Request.Form["Emps"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_HelpdeskGroupMapping_SaveUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        #endregion Helpdesk Group Mapping


        /// <summary>
        /// Application Settings
        /// </summary>
        /// <returns></returns>
        #region Application Settings
        public ActionResult ApplicationSettings()
        {
            comfun.saveformname("ApplicationSettings", "/admin/ApplicationSettings", "App Settings", "App Setting form", "Y", "AppSetting", "AppSetting", "View", 1);
            return View();
        }
        #endregion App Settings

        /// <summary>
        /// /Conveyance Group List
        /// </summary>
        /// <returns></returns>
        #region Conveyance Group List
        public ActionResult ConveyanceGroupList()
        {
            comfun.saveformname("ConveyanceGroupList", "/admin/ConveyanceGroupList", "Master/Conveyance&nbsp;/Group&nbsp;", "Conveyance Group Main form", "Y", "ConveyanceGroupList", "ConveyanceGroupList", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("ConveyanceGroupList");
            return View();
        }
        public ActionResult _ConveyanceGroupList()
        {
            comfun.saveformname("_ConveyanceGroupList", "/admin/_ConveyanceGroupList", "Master/Conveyance&nbsp;/Group&nbsp;", "Conveyance Group  List", "N", "ConveyanceGroupList", "ConveyanceGroupList", "List", 4);

            SortedList list = new SortedList();
            DataTable dt = comfun.fillDataTable("sp_ConveyanceGroupListDisplay", "", null);
            //return PartialView(dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _ConveyanceGroupAdd()
        {
            return PartialView("_conveyanceGroupAdd");
        }
        public JsonResult ConveyanceGroupSaveUpdateSubmit()
        {
            comfun.saveformname("ConveyanceGroupSaveUpdateSubmit", "/admin/ConveyanceGroupSaveUpdateSubmit", "Master/Conveyance&nbsp;/Group&nbsp;", "Conveyance Group Add", "N", "ConveyanceGroupList", "ConveyanceGroupList", "Add", 2);
            comfun.saveformname("ConveyanceGroupSaveUpdateSubmit", "/admin/ConveyanceGroupSaveUpdateSubmit", "Master/Conveyance&nbsp;/Group&nbsp;", "Conveyance Group Edit", "N", "ConveyanceGroupList", "ConveyanceGroupList", "Edit", 3);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@ConveyanceGroupId", Request.Form["GroupId"].ToString());
                list.Add("@ConveyanceGroupCode", Request.Form["GroupCode"].ToString());
                list.Add("@ConveyanceGroupName", Request.Form["GroupName"].ToString());
                list.Add("@PricePerKM", Request.Form["GroupPricePerKM"].ToString());
                list.Add("@MonthlyLimit", Request.Form["GroupMonthlyLimit"].ToString());
                mes = comfun.executeNonQueryWMessage("ConveyanceGroupSaveUpdateSubmit", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_CompnayNatureAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        public JsonResult ConveyanceGroupStatusUpdate()
        {
            comfun.saveformname("ConveyanceGroupStatusUpdate", "/admin/ConveyanceGroupStatusUpdate", "Master/Conveyance&nbsp;/Group&nbsp;", "Conveyance Group Main Status", "N", "ConveyanceGroupList", "ConveyanceGroupList", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_ConveyanceGroup_Active_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        public ActionResult _ConveyanceGroupEdit()
        {
            SortedList list = new SortedList();
            list.Add("@ConveyanceGroupId", Request.Form["ConveyanceGroupId"].ToString());
            DataTable dt = comfun.fillDataTable("sp_ConveyanceGroupDisplayViaId", "", list);
            return PartialView("_conveyanceGroupEdit", dt);
        }
        #endregion Conveyance Group List




        // GET: Admin
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
        public ActionResult _Paycharges()
        {
            SortedList list = new SortedList();
            list.Add("@PartId", Request.Form["PartId"].ToString());
            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("Charges_SelectM", "", list);
            return PartialView("_Paycharges", dt);

        }
        [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult PayChargesAdd()
        {
            //comfun.saveformname("PayChargesAdd", "/admin/PayChargesAdd", "Pay Charges Add", "Pay Charges view", "N", "PayChargesAdd", "PayCharges", "Add");
            return View();
        }
        public JsonResult _PayChargesAddSubmit()
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
                list.Add("@ChargeName", Request.Form["ChargeName"].ToString());
                list.Add("@ChargeType", Request.Form["ChargeType"].ToString());
                list.Add("@CtcType", Request.Form["CtcType"].ToString());
                list.Add("@CalculationType", Request.Form["CalculationType"].ToString());
                list.Add("@IsPF", Request.Form["IsPF"].ToString());
                list.Add("@IsESI", Request.Form["IsESI"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                list.Add("@Part", Request.Form["Part"].ToString());
                list.Add("@ChargeWEF", Request.Form["ChargeWEF"].ToString());
                mes = comfun.executeNonQueryWMessage("Charges_Save", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_PayChargesAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult PayChargesEdit()
        {
            //comfun.saveformname("PayChargesEdit", "/admin/PayChargesEdit", "Pay Charges Edit", "Pay Charges Edit", "N", "PayChargesEdit", "PayCharges", "Edit");

            string chargeid = "0";
            if (Request.QueryString["id"] == null)
            {
                return RedirectToAction("PayCharges");
            }
            chargeid = comfun.decryptString(Request.QueryString["Id"].ToString());
            SortedList list = new SortedList();
            list.Add("@ChargesId", chargeid);
            DataTable dt = comfun.fillDataTable("Charges_SelectWithId", "", list);
            return View(dt);
        }
        public JsonResult _PayChargesEditSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@ChargeName", Request.Form["ChargeName"].ToString());
                list.Add("@ChargeType", Request.Form["ChargeType"].ToString());
                list.Add("@CtcType", Request.Form["CtcType"].ToString());
                list.Add("@CalculationType", Request.Form["CalculationType"].ToString());
                list.Add("@IsPF", Request.Form["IsPF"].ToString());
                list.Add("@IsESI", Request.Form["IsESI"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                list.Add("@ChargesId", Request.Form["ChargeId"].ToString());
                list.Add("@ChargeWEF", Request.Form["ChargeWEF"].ToString());
                list.Add("@Part", Request.Form["Part"].ToString());
                mes = comfun.executeNonQueryWMessage("Charges_Edit", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_PayChargesEditSubmit", "Error:" + ex.Message);
            }
            return Json(mes);
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






        [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult PayChargeApplied()
        {
            //comfun.saveformname("PayChargeApplied", "/admin/PayChargeApplied", "Employee CTC", "Employee CTC", "Y", "EmployeeCTC", "EmployeeCTC", "List");
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            DataSet ds = comfun.fillDataSet("Department_PayChargesSelectDS", "", null);
            return View(ds);
        }
        public ActionResult _pay_charge_applied()
        {
            SortedList list = new SortedList();
            list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
            DataTable dt = comfun.fillDataTable("Employee_SelectDepartment", "", list);
            return PartialView("_pay_charge_applied", dt);
        }
        public JsonResult _PayChargesSubmit()
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
                list.Add("@ChargeAppliedDate", Request.Form["ChargeAppliedDate"].ToString());
                list.Add("@EmployeeId", Request.Form["EmployeeId"].ToString());
                list.Add("@GrossPay", Request.Form["GrossPay"].ToString());
                mes = comfun.executeNonQueryWTranOutMes("payChargeAppliedPart1_Save", list, conObj.con, tran).ToString();
                string partId = mes;
                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject(Request.Form["PayChargeAppliedPart2"].ToString());

                foreach (var item in jsonData)
                {
                    list.Clear();
                    list.Add("@ChargeAppliedID", partId);
                    list.Add("@ChargePercentage", "0.00");
                    list.Add("@ChargeAmount", item.ChargeAmount);
                    list.Add("@EmployeeId", Request.Form["EmployeeId"].ToString());
                    list.Add("@ChargeAppliedDate", Request.Form["ChargeAppliedDate"].ToString());
                    list.Add("@PayChargeId", item.PayChargeId);
                    mes = comfun.executeNonQueryWTranOutMes("payChargeAppliedPart2_Save", list, conObj.con, tran).ToString();
                }
                mes = "Record saved successfully";
                tran.Commit();
                conObj.con.Close();
            }
            catch (Exception ex)
            {
                tran.Rollback();
                conObj.con.Close();
                mes = comfun.errorMessage("Admin _PayChargesSubmit", ex.Message);
            }


            return Json(mes);
        }






        public ActionResult ShiftPlanning()
        {
            // comfun.saveformname("ShiftPlanning", "/admin/ShiftPlanning", "Shift Planning", "ShiftPlaning view", "N", "ShiftPlanning", "ShiftPlaning", "Add");

            DataSet ds = comfun.fillDataSet("Branch_Select", "", null);
            return View(ds);
        }
        public ActionResult _shiftPlanning_display()
        {
            SortedList list = new SortedList();
            list.Add("@BranchId", Request.Form["BranchId"].ToString());
            DataTable dt = comfun.fillDataTable("ShiftSelect_Branch", "", list);
            return PartialView("_shiftPlanning_display", dt);
        }
        public ActionResult shiftPlanningAdd()
        {
            //comfun.saveformname("shiftPlanningAdd", "/admin/shiftPlanningAdd", "Shift Plan Add", "Shift Plan", "N", "shiftPlanningAdd", "ShiftPlanning", "Add");

            DataSet ds = comfun.fillDataSet("Branch_Select", "", null);
            return View(ds);
        }
        public JsonResult _ShiftPlanningAddSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@BranchId", Request.Form["BranchId"].ToString());
                list.Add("@ShiftId", Request.Form["ShiftId"].ToString());
                list.Add("@FromDate", Request.Form["FromDate"].ToString());
                list.Add("@ToDate", Request.Form["ToDate"].ToString());
                list.Add("@StartTime", Request.Form["StartTime"].ToString());
                list.Add("@EndTime", Request.Form["EndTime"].ToString());

                list.Add("@BreakStartTime", Request.Form["BreakStartTime"].ToString());
                list.Add("@BreakEndTime", Request.Form["BreakEndTime"].ToString());
                list.Add("@CreatedBy", "1");
                list.Add("@CreatedDate", comfun.dateISTstr());

                mes = comfun.executeNonQueryWMessage("ShiftPlanning_Save", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _ShiftPlanningAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult ShiftEdit()
        {
            //comfun.saveformname("ShiftEdit", "/admin/ShiftEdit", "Shift Edit", "Shift view", "N", "ShiftEdit", "Shift", "Edit");

            string Id = "0";
            if (Request.QueryString["id"] == null)
            {
                return RedirectToAction("Shift");
            }
            Id = comfun.decryptString(Request.QueryString["Id"].ToString());
            SortedList list = new SortedList();
            list.Add("@Id", Id);
            DataSet ds = comfun.fillDataSet("ShiftSelectWithId_DS", "", list);
            return View(ds);
        }
        public JsonResult _ShiftEditSubmit()
        {
            comfun.saveformname("_ShiftEditSubmit", "/admin/_ShiftEditSubmit", "Shift Management/Shift Assign", "Shift Assign  edit", "N", "ShiftAssign", "ShiftAssign", "Edit", 3);

            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@ShiftID", Request.Form["ShiftId"].ToString());
                list.Add("@StartTime", Request.Form["StartTime"].ToString());
                list.Add("@EndTime", Request.Form["EndTime"].ToString());
                list.Add("@ShiftID", Request.Form["ShiftID"].ToString());
                list.Add("@BreakStartTime", Request.Form["BreakStartTime"].ToString());
                list.Add("@BreakEndTime", Request.Form["BreakEndTime"].ToString());
                list.Add("@BranchId", Request.Form["BranchId"].ToString());
                mes = comfun.executeNonQueryWMessage("ShiftDetail_Edit", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _ShiftEditSubmit", ex.Message);
            }
            return Json(mes);
        }



        public ActionResult _Shift_select_branch()
        {
            SortedList list = new SortedList();
            list.Add("@BranchId", Request.Form["BranchId"].ToString());
            DataTable dt = comfun.fillDataTable("ShiftSelect_Branch", "", list);
            return PartialView("_Shift_select_branch", dt);
        }

        //  [customAuthorize(Roles = payrollFunctions.rolePayrollTeam)]
        public ActionResult Allowances()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return PartialView("_sessionExpired");

            }
            SortedList list = new SortedList();
            list.Add("@ctcType", "A");
            DataTable dt = comfun.fillDataTable("PayCharges_GetHeadsByCtcType", "", list);
            return View(dt);
        }
        public ActionResult _Employeeddl()
        {
            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("stpEmployeedropdown", null, null);
            return PartialView("_Employeeddl", dt);
        }
        public JsonResult _allowancesSubmit()
        {

            string mes = string.Empty;
            try
            {
                DateTime monthYear = Convert.ToDateTime(Request.Form["AllowanceDate"]);
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@fiEmployeeId", Request.Form["EmployeeId"].ToString());
                list.Add("@fiMonthId", monthYear.ToString("MM"));
                list.Add("@fdAllowanceDate", Convert.ToDateTime(Request.Form["AllowanceDate"]).ToString("dd-MMM-yyyy"));
                list.Add("@fiChargeID", Request.Form["Allowances"].ToString());
                list.Add("@fnAmount", Request.Form["Amount"].ToString());
                list.Add("@fiSessionID", Session["SessionId"].ToString());
                list.Add("@fnPercentage", "0.00");
                list.Add("@fiMonthYear", monthYear.ToString("MMyyyy"));
                mes = comfun.executeNonQueryWMessage("stpVIKSATAllowanceEntryPart1_Accept", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _allowancesSubmit", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult _allowancesList()
        {
            SortedList list = new SortedList();
            string fromDate = Request.Form["StartDate"].ToString();
            int startYear = Convert.ToDateTime(fromDate).Year;
            int startMonth = Convert.ToDateTime(fromDate).Month;

            DateTime startDate = new DateTime(startYear, startMonth, 1);
            DateTime endDate = startDate.AddMonths(1).AddDays(-1);
            list.Add("@StartDate", startDate.ToString("dd/MMM/yyyy"));
            list.Add("@EndDate", endDate.ToString("dd/MMM/yyyy"));
            DataTable dt = comfun.fillDataTable("stpVIKSATAllowanceGridWithoutid", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult _AllowanceEdit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return PartialView("_sessionExpired");
            }
            //comfun.saveformname("AllowanceEdit", "/admin/AllowanceEdit", "Allowance Edit", "Allowance Edit", "N", "AllowanceEdit", "Allowance", "Edit");

            string AllowanceId = "0";
            if (Request.Form["Id"].ToString() == null)
            {
                return RedirectToAction("Allowances");
            }
            AllowanceId = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@AllowanceId", AllowanceId);
            DataTable dt = comfun.fillDataTable("AllowanceSelectById", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);

        }
        public JsonResult _AllowanceEditSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@AllowanceDate", Convert.ToDateTime(Request.Form["AllowanceDate"]).ToString("dd-MMM-yyyy"));
                list.Add("@AllowanceId", Request.Form["AllowanceID"].ToString());
                list.Add("@Allowances", Request.Form["Allowances"].ToString());
                list.Add("@Amount", Request.Form["Amount"].ToString());

                mes = comfun.executeNonQueryWMessage("Allowances_Edit", "", list).ToString();
            }
            catch (Exception ex)

            {
                mes = comfun.errorMessage("Admin _ShiftEditSubmit", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult AllowanceDelete()
        {
            //comfun.saveformname("AllowanceDelete", "/admin/AllowanceDelete", "Allowance Delete", "Allowance Dlete", "N", "AllowanceDelete", "Allowance", "Delete");

            string AllowanceId = "0";
            if (payfun.sessionRecreate() == "expires")
            {
                return PartialView("_sessionExpired");
            }
            if (Request.QueryString["id"] == null)
            {
                return RedirectToAction("Allowances");
            }
            AllowanceId = comfun.decryptString(Request.QueryString["Id"].ToString());
            SortedList list = new SortedList();
            list.Add("@AllowanceId", AllowanceId);
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@Createddate", comfun.dateISTstr());
            DataTable dt = comfun.fillDataTable("stpviksatAllowanceDelete", "", list);
            return View();

        }


        // [customAuthorize(Roles = payrollFunctions.rolePayrollTeam)]
        public ActionResult Deductions()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return PartialView("_sessionExpired");
            }
            SortedList list = new SortedList();
            list.Add("@ctcType", "D");
            DataTable dt = comfun.fillDataTable("PayCharges_GetHeadsByCtcType", "", list);
            return View(dt);
        }
        public ActionResult _DeductionList()
        {
            DateTime monthYear = Convert.ToDateTime(Request.Form["DeductionDate"]);

            SortedList list = new SortedList();
            string fromDate = Request.Form["StartDate"].ToString();
            int startYear = Convert.ToDateTime(fromDate).Year;
            int startMonth = Convert.ToDateTime(fromDate).Month;

            DateTime startDate = new DateTime(startYear, startMonth, 1);
            DateTime endDate = startDate.AddMonths(1).AddDays(-1);
            list.Add("@StartDate", startDate.ToString("dd/MMM/yyyy"));
            list.Add("@EndDate", endDate.ToString("dd/MMM/yyyy"));
            DataTable dt = comfun.fillDataTable("stpVIKSATDeductionGridWithoutid", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public JsonResult _deductionsSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            try
            {
                DateTime monthYear = Convert.ToDateTime(Request.Form["DeductionDate"]);
                SortedList list = new SortedList();
                list.Add("@fiEmployeeId", Request.Form["EmployeeId"].ToString());
                list.Add("@fiMonthId", monthYear.ToString("MM"));
                list.Add("@fdDeductionDate", Convert.ToDateTime(Request.Form["DeductionDate"]).ToString("dd-MMM-yyyy"));
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@fiChargeID", Request.Form["DeductionId"].ToString());
                list.Add("@fnAmount", Request.Form["Amount"].ToString());
                list.Add("@fiSessionID", Session["SessionId"].ToString());
                list.Add("@fnPercentage", "0.00");
                list.Add("@fiMonthYear", monthYear.ToString("MMyyyy"));
                mes = comfun.executeNonQueryWMessage("stpVIKSATDeductionEntryPart1_Accept", "", list).ToString();
            }
            catch (Exception ex)
            {

                mes = comfun.errorMessage("Admin _deductionsSubmit", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult _DeductionEdit()
        {
            //comfun.saveformname("DeductionEdit", "/admin/DeductionEdit", "Deduction Edit", "Deduction Edit", "N", "DeductionEdit", "Deduction", "Edit");
            if (payfun.sessionRecreate() == "expires")
            {
                return PartialView("_sessionExpired");
            }
            string DeductionId = "0";

            DeductionId = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@DeductionId", DeductionId);
            DataTable dt = comfun.fillDataTable("DeductionSelectById", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public JsonResult _DeductionEditSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@DeductionDate", Convert.ToDateTime(Request.Form["DeductionDate"]).ToString("dd-MMM-yyyy"));
                list.Add("@DeductionId", Request.Form["DeductionId"].ToString());
                list.Add("@Deductions", Request.Form["Deductions"].ToString());
                list.Add("@Amount", Request.Form["Amount"].ToString());

                mes = comfun.executeNonQueryWMessage("Deductions_Edit", "", list).ToString();
            }
            catch (Exception ex)
            {

                mes = comfun.errorMessage("Admin _deductionsEditSubmit", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult DeductionDelete()
        {
            //comfun.saveformname("DeductionDelete", "/admin/DeductionDelete", "Deduction Delete", "Deduction Dlete", "N", "DeductionDelete", "Deduction", "Delete");

            string DeductionId = "0";
            if (Request.QueryString["id"] == null)
            {
                return RedirectToAction("Deductions");
            }
            DeductionId = comfun.decryptString(Request.QueryString["Id"].ToString());
            SortedList list = new SortedList();
            list.Add("@DeductionId", DeductionId);
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@Createddate", comfun.dateISTstr());
            DataTable dt = comfun.fillDataTable("stpviksatDeductionDelete", "", list);
            return View();

        }
        public ActionResult DayDeduction()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            //comfun.saveformname("DayDeduction", "/admin/DayDeduction", "DayDeduction", "DayDeduction", "Y", "DayDeduction", "DayDeduction", "List");
            return View();
        }
        public JsonResult _DayDeductionsSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            try
            {
                DateTime monthYear = Convert.ToDateTime(Request.Form["LeaveDate"]);
                SortedList list = new SortedList();
                list.Add("@fiEmployeeId", Request.Form["EmployeeId"].ToString());

                list.Add("@LeaveDate", Convert.ToDateTime(Request.Form["LeaveDate"]).ToString("dd-MMM-yyyy"));

                list.Add("@Days", Request.Form["Days"].ToString());
                list.Add("@fiSessionID", Session["SessionId"].ToString());
                list.Add("@Remarks", Request.Form["Remarks"].ToString());
                list.Add("@fiMonthYear", monthYear.ToString("MMyyyy"));
                mes = comfun.executeNonQueryWMessage("stpVIKSATDeductionDays_Accept", "", list).ToString();
            }
            catch (Exception ex)
            {

                mes = comfun.errorMessage("Admin _deductionsSubmit", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult _DaysdeductionsShow()
        {
            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@DeductionDate", Convert.ToDateTime(Request.Form["LeaveDate"]).ToString("dd-MMM-yyyy"));
            list.Add("@fiSessionId", Session["SessionId"].ToString());
            list.Add("@CostcenterId", Session["CostCenterId"].ToString());
            DataTable dt = comfun.fillDataTable("stpviksatleaveWithoutPay", "", list);
            return PartialView("_DayDeductionshow", dt);
        }
        public ActionResult DayDeductionEdit()
        {
            //comfun.saveformname("DayDeductionEdit", "/admin/DayDeductionEdit", "Day Deduction Edit", "Day Deduction Edit", "N", "DayDeductionEdit", "Day Deduction", "Edit");

            string DayDeductionId = "0";
            if (Request.QueryString["id"] == null)
            {
                return RedirectToAction("DayDeduction");
            }
            DayDeductionId = comfun.decryptString(Request.QueryString["Id"].ToString());
            SortedList list = new SortedList();
            list.Add("@Id", DayDeductionId);
            DataTable dt = comfun.fillDataTable("DayDeductionSelectById", "", list);
            return View(dt);
        }
        public JsonResult _DayDeductionEditSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@LeaveDate", Convert.ToDateTime(Request.Form["LeaveDate"]).ToString("dd-MMM-yyyy"));
                list.Add("@Days", Request.Form["Days"].ToString());
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@Remarks", Request.Form["Remarks"].ToString());

                mes = comfun.executeNonQueryWMessage("DayDeductions_Edit", "", list).ToString();
            }
            catch (Exception ex)
            {

                mes = comfun.errorMessage("Admin _deductionsEditSubmit", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult DayDeductionDelete()
        {
            //comfun.saveformname("DayDeductionDelete", "/admin/DayDeductionDelete", "Day Deduction Delete", "Day Deduction Dlete", "N", "DayDeductionDelete", "DayDeduction", "Delete");

            string DayDeductionId = "0";
            if (Request.QueryString["id"] == null)
            {
                return RedirectToAction("DayDeduction");
            }
            DayDeductionId = comfun.decryptString(Request.QueryString["Id"].ToString());
            SortedList list = new SortedList();
            list.Add("@Id", DayDeductionId);
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@Createddate", comfun.dateISTstr());
            DataTable dt = comfun.fillDataTable("stpviksatDayDeductionDelete", "", list);
            return View();

        }
        //used Loan Type in HTIS Project//
        public ActionResult LoanTypeList()
        {
            // comfun.saveformname("LoanType", "/admin/LoanTypeList", "Payroll Management/Master/Loan Type", "Loan Type form", "Y", "LoanType", "LoanType", "View", 1);
            //ViewData["AccessRights"] = payfun.getAccessRights("LoanType");
            return View();
        }
        public ActionResult _LoanTypeList()
        {
            //comfun.saveformname("_LoanTypeList", "/admin/_LoanTypeList", "Payroll Management/Master/Loan Type", "Loan Type form", "N", "LoanType", "LoanType", "List", 4);

            //comfun.saveformname("LoanType", "/admin/LoanTypeList", "LoanType", "LoanType", "Y", "Loan", "Loan", "List");
            DataTable dt = comfun.fillDataTable("stpLoanType_List", "", null);

            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public JsonResult _LoanTypeAddSubmit()
        {
            // comfun.saveformname("_LoanTypeAddSubmit", "/admin/_LoanTypeAddSubmit", "Payroll Management/Master/Loan Type", "Loan Type Add", "N", "LoanType", "LoanType", "Add", 2);
            // comfun.saveformname("_LoanTypeAddSubmit", "/admin/_LoanTypeAddSubmit", "Payroll Management/Master/Loan Type", "Loan Type Edit", "N", "LoanType", "LoanType", "Edit", 3);

            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            SortedList list = new SortedList();
            string mes = string.Empty;
            try
            {
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@fvLoanType ", Request.Form["loantype"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                list.Add("@CreatedDate", comfun.dateISTstr());

                mes = comfun.executeNonQueryWMessage("stpLoan_UpdateAccept", "", list).ToString();

            }
            catch (Exception ex)
            {

                mes = comfun.errorMessage("Error:", ex.Message);
            }
            return Json(mes);
        }


        public JsonResult _LoanTypeActivationStatusUpdate()
        {
            // comfun.saveformname("_LoanTypeActivationStatusUpdate", "/admin/_LoanTypeActivationStatusUpdate", "Payroll Management/Master/Loan Type", "Loan Type Status", "N", "LoanType", "LoanType", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_LoanTypeStatusActiveDeactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }
        // [customAuthorize(Roles = payrollFunctions.rolePayrollTeam)]
        public ActionResult Loanissued()
        {
            //comfun.saveformname("Loanissued", "/admin/Loanissued", "Loanissued", "Loanissued", "Y", "Loan Issue", "Loan Issue", "List");
            DataTable dt = comfun.fillDataTable("stpLoanTypeddl", "", null);
            return View(dt);
        }
        public JsonResult _loanissuedSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@fiEmployeeId", Request.Form["employeeId"].ToString());
                list.Add("@fdLoanIssueDate", Request.Form["IssueDate"].ToString());
                list.Add("@fiLoanTtypeID", Request.Form["loanType"].ToString());
                list.Add("@fdInstallmentStartDate", Request.Form["InstallmentstartDate"].ToString());
                list.Add("@fnLoanAmount", Request.Form["loanamount"].ToString());
                //list.Add("@fnInstallmentAmount", Request.Form["InstallmentAmount"].ToString());
                list.Add("@fnNoofInstalment", Request.Form["Installment"].ToString());
                mes = comfun.executeNonQueryWMessage("stpVIKSATLoanIssuedPart1_Accept", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _loanissuedSubmit", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult LoanList()
        {
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            // SortedList list = new SortedList();
            // // list.Add("@Employeeid", Session["EmpId"].ToString());
            //// list.Add("@CostCenterId", Session["CostCenterId"].ToString());
            //// list.Add("@CompanyId", Session["CompanyId"].ToString());

            DataTable dt = comfun.fillDataTable("stpLoanType_List", "", null);
            return View(dt);
        }
        public ActionResult _LoanList()
        {
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            SortedList list = new SortedList();
            // list.Add("@Employeeid", Session["EmpId"].ToString());
            // list.Add("@CostCenterId", Session["CostCenterId"].ToString());
            // list.Add("@CompanyId", Session["CompanyId"].ToString());

            DataTable dt = comfun.fillDataTable("stpLoanIssued_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult LoanPaid()
        {
            return View();
        }

        public ActionResult _LoanIssueDetail()
        {

            SortedList list = new SortedList();
            list.Add("@fiLoanId", Request.Form["fiLoanID"].ToString());
            DataTable dt = comfun.fillDataTable("stpViksatLoandetailByID", "", list);
            return PartialView("_LoanDetail", dt);

        }
        public ActionResult LoanEdit()
        {
            string LoanId = "0";
            if (Request.Form["Id"] == null)
            {
                return RedirectToAction("LoanList");
            }
            LoanId = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@LoanId", LoanId);
            DataTable dt = comfun.fillDataTable("stpLoanSelectByLoanId", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public ActionResult _LoanEditSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                //list.Add("@fiEmployeeID", Request.Form["employeeId"].ToString());
                list.Add("@fiLoanTtypeID", Request.Form["loanType"].ToString());
                list.Add("@fnLoanAmount", Request.Form["loanamount"].ToString());
                list.Add("@fnNoofInstalment", Request.Form["Installment"].ToString());
                //list.Add("@fnInstallmentAmount",Request.Form[""].ToString());
                list.Add("@fdLoanIssueDate", Request.Form["IssueDate"].ToString());
                list.Add("@fdInstallmentStartDate", Request.Form["InstallmentstartDate"].ToString());
                list.Add("@fiLoanID", Request.Form["LoanId"].ToString());
                mes = comfun.executeNonQueryWMessage("stpViksatLoanissuedPart1_Update", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _LoanEditSubmit", ex.Message);
            }
            return Json(mes);
        }
        public JsonResult _LoanDetail(string LoanData)
        {
            string mes = string.Empty;
            try
            {
                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject(LoanData);
                DataTable table = new DataTable();
                table.Columns.Add("fiLoanId", typeof(int));
                table.Columns.Add("fiEmployee", typeof(int));
                table.Columns.Add("fiMonthYearId", typeof(int));
                table.Columns.Add("fiLoanDetailId", typeof(int));
                table.Columns.Add("InstallmentAmount", typeof(string));
                table.Columns.Add("InstallmentDate", typeof(string));
                table.Columns.Add("IsStop", typeof(string));
                table.Columns.Add("Remarks", typeof(string));
                foreach (var item in jsonData)
                {
                    DataRow dr = table.NewRow();
                    dr["fiLoanId"] = item.LoanId;
                    dr["fiEmployee"] = item.Employee;
                    dr["fiMonthYearId"] = item.MonthYearId;
                    dr["fiLoanDetailId"] = item.LoandetailId;
                    dr["InstallmentAmount"] = item.LoanAmount;
                    dr["InstallmentDate"] = item.InstallmentDate;
                    dr["IsStop"] = item.Status;
                    dr["Remarks"] = item.Remarks;
                    table.Rows.Add(dr);
                }

                SqlCommand com = new SqlCommand();
                connection conObj = new connection();

                com.Connection = conObj.con;

                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "stpViksatLoanIssuedpart2_Accept";
                com.Parameters.AddWithValue("@CreateDate", comfun.dateISTstr());
                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@LoaninsertTable5";

                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table;
                com.Parameters.Add(parameter);


                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();
                com.ExecuteNonQuery();
                conObj.con.Close();

                mes = "Record saved successfully";
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _LoanDetail", "Error: " + ex.Message);
            }
            return Json(mes);
        }

        [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult Contractor()
        {
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            //comfun.saveformname("Contractor", "/admin/Contractor", "Contractor", "Contractor  form", "Y", "", "Contractor", "List");

            return View();
        }
        public ActionResult _ContractorList()
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

            DataTable dt = comfun.fillDataTable("stpviksatContactorList", "", list);
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "admin", "_ContractorList", "_ContractorList");

            HtmlString htm = new HtmlString(paging);
            ViewData["paging"] = htm;
            return PartialView("_ContractorList", dt);
        }
        [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult ContractorEdit()
        {
            //comfun.saveformname("ContractorEdit", "/admin/ContractorEdit", "Contractor Edit", "Contractor Edit", "N", "ContractorEdit", "Contractor", "Edit");
            string Contractorid = "0";
            if (Request.QueryString["id"] == null)
            {
                return RedirectToAction("Contractor");
            }
            Contractorid = comfun.decryptString(Request.QueryString["Id"].ToString());
            SortedList list = new SortedList();
            list.Add("@ContractorID", Contractorid);
            DataSet ds = comfun.fillDataSet("stpviksatContactorWithID", "", list);
            return View(ds);

        }
        [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult ContractorAdd()
        {
            //comfun.saveformname("ContractorAdd", "/admin/ContractorAdd", "Contractor Add", "Contractor Add view", "N", "ContractorAdd", "Contractor", "Add");

            DataSet ds = comfun.fillDataSet("stpViksatBranchCompanyStateCity", "", null);
            return View(ds);
        }
        public JsonResult _contractorAdd()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@ContractorName", Request.Form["ContractorName"].ToString());
                list.Add("@ContractorMobileNo", Request.Form["ContractorMobileNo"].ToString());
                list.Add("@ContractorAddress", Request.Form["ContractorAddress"].ToString());
                //list.Add("@Gender", Request.Form["Gender"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                list.Add("@Surcharge", Request.Form["Surcharge"].ToString());
                list.Add("@CreatedDate", comfun.dateISTstr());
                list.Add("@AadhaarNo", Request.Form["AadhaarNo"].ToString());
                list.Add("@PanNumber", Request.Form["PanNumber"].ToString());
                list.Add("@PFNO", Request.Form["PFNO"].ToString());
                list.Add("@ESINo", Request.Form["ESINo"].ToString());
                list.Add("@City", Request.Form["City"].ToString());
                list.Add("@State", Request.Form["State"].ToString());
                list.Add("@GSTNO", Request.Form["GSTNO"].ToString());
                list.Add("@ContractorCode", Request.Form["ContractorCode"].ToString());
                list.Add("@ContractorCompanyName", Request.Form["ContractorCompanyName"].ToString());
                list.Add("@BranchID", Request.Form["BranchID"].ToString());
                list.Add("@CompanyID", Request.Form["CompantID"].ToString());
                list.Add("@ContractorERPBPCode", Request.Form["ContractorERPBPCode"].ToString());
                mes = comfun.executeNonQueryWMessage("stpviksatContractorSave", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _contractorAdd", ex.Message);
            }
            return Json(mes);
        }
        //public JsonResult _contractorEdit()
        //{
        //    if (payfun.sessionRecreate() == "expires")
        //    {
        //        return Json("Session expires");
        //    }
        //    string mes = string.Empty;
        //    try
        //    {
        //        SortedList list = new SortedList();
        //        list.Add("@ContractorID", Request.Form["ContractorID"].ToString());
        //        list.Add("@ContractorName", Request.Form["ContractorName"].ToString());
        //        list.Add("@ContractorMobileNo", Request.Form["ContractorMobileNo"].ToString());
        //        list.Add("@ContractorAddress", Request.Form["ContractorAddress"].ToString());
        //        //  list.Add("@Gender", Request.Form["Gender"].ToString());
        //        list.Add("@AadhaarNo", Request.Form["AadhaarNo"].ToString());
        //        list.Add("@PanNumber", Request.Form["PanNumber"].ToString());
        //        list.Add("@CityID", Request.Form["City"].ToString());
        //        list.Add("@ESINo", Request.Form["ESINO"].ToString());
        //        list.Add("@PFNO", Request.Form["PFNo"].ToString());
        //        list.Add("@Surcharge", Request.Form["Surcharge"].ToString());
        //        list.Add("@StateID", Request.Form["State"].ToString());
        //        list.Add("@GSTNO", Request.Form["GSTNO"].ToString());
        //        list.Add("@ContractorCompanyName", Request.Form["ContractorCompanyName"].ToString());
        //        list.Add("@BranchID", Request.Form["BranchID"].ToString());
        //        list.Add("@CompanyID", Request.Form["CompanyID"].ToString());
        //        list.Add("@ContractorERPBPCode", Request.Form["ContractorERPBPCode"].ToString());
        //        mes = comfun.executeNonQueryWMessage("stpviksatContractorUpdate", "", list).ToString();
        //    }
        //    catch (Exception ex)
        //    {
        //        mes = comfun.errorMessage("Admin _contractorEdit", ex.Message);
        //    }
        //    return Json(mes);
        //}

        [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult MainCategoryAdd()
        {
            return View();
        }
        public JsonResult MainCategoryInsert()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@MainCategoryName", Request.Form["MainCategoryName"].ToString());
                mes = comfun.executeNonQueryWMessage("stpViksatpayMainCategoryinsert", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin MainCategoryInsert", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult MainCategoryShow()
        {
            SortedList list = new SortedList();
            DataTable dt = comfun.fillDataTable("stpviksatPaymainCategoryShow", "", list);
            return PartialView("_MainCategorylist", dt);
        }

        [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult LeaveTypeList()
        {
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            //comfun.saveformname("LeaveTypeList", "/admin/LeaveTypeList", "Leave Type List", "Leave Type Main form", "Y", "", "LeaveType", "List");
            DataTable dt = comfun.fillDataTable("stpviksatLeaveTypeListShow", "", null);
            return View(dt);
        }
        [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult LeavetypeAdd()
        {
            // comfun.saveformname("LeavetypeAdd", "/admin/LeavetypeAdd", "Leave Type Add", "Leave Type Add view", "N", "LeaveTypeAdd", "LeaveType", "Add");
            return View();
        }
        public JsonResult _leaveTypeSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@fvLeaveTypeName", Request.Form["LeaveType"].ToString());
                list.Add("@Code", Request.Form["LeaveTypeCode"].ToString());

                list.Add("@Value", Request.Form["LeaveTypeValue"].ToString());

                // list.Add("@ftLeaveTypeId", 0);

                mes = comfun.executeNonQueryWMessage("stpVIKSATLeaveTypes_Accept", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _leaveTypeSubmit", ex.Message);
            }
            return Json(mes);
        }
        [customAuthorize(Roles = _roles)]
        public ActionResult LeavetypeEdit()
        {

            //comfun.saveformname("LeavetypeEdit", "/admin/LeavetypeEdit", "Leave Type Edit", "Leave Type Edit", "N", "LeavetypeEdit", "LeaveType", "Edit");


            string LeaveTypeID = "0";
            if (Request.QueryString["id"] == null)
            {
                return RedirectToAction("LeaveTypeList");
            }
            LeaveTypeID = comfun.decryptString(Request.QueryString["Id"].ToString());
            SortedList list = new SortedList();
            list.Add("@LeaveTypeID", LeaveTypeID);
            DataTable dt = comfun.fillDataTable("stpviksatLeaveTypeListSelectByID", "", list);
            return View(dt);
        }
        public JsonResult _LeaveTypeEditSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@ftLeaveTypeId", Request.Form["LeaveTypeId"].ToString());
                list.Add("@fvLeaveTypeName", Request.Form["LeaveType"].ToString());
                list.Add("@Code", Request.Form["LeaveTypeCode"].ToString());
                list.Add("@Value", Request.Form["LeaveTypeValue"].ToString());
                list.Add("@branchId", Session["CostCenterId"].ToString());
                mes = comfun.executeNonQueryWMessage("stpViksatLeaveTypes_Update", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _LeaveTypeEditSubmit", ex.Message);
            }
            return Json(mes);
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

        public ActionResult Leaveopeningbalance()
        {
            //comfun.saveformname("Leaveopeningbalance", "/admin/Leaveopeningbalance", "Leave Opening Balance", "Leave Opening Balance", "Y", "Leaveopeningbalance", "LeaveBalance", "Add");

            DataSet ds = comfun.fillDataSet("stpviksatQ1Leavetype", "", null);
            return View(ds);
        }
        public JsonResult _LeaveBalanceAdd()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;

            SortedList list = new SortedList();
            try
            {
                list.Add("@fiEmployeeid", Request.Form["Employeeid"].ToString());
                list.Add("@fiLeaveTypeId", Request.Form["LeaveType"].ToString());
                list.Add("@fvNoofLeaves", Request.Form["NoofLeaves"].ToString());
                list.Add("@fiSessionId", Session["SessionId"].ToString());
                list.Add("@fiBranchID", Session["CostCenterId"]);
                mes = comfun.executeNonQueryWMessage("stpviksatleavebalancesubmit", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _LeaveBalanceAdd", ex.Message);
            }
            return Json(mes);
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

        public ActionResult SessionList()
        {
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            DataTable dt = comfun.fillDataTable("SessionList", "", null);
            return View(dt);
        }
        public ActionResult SessionAdd()
        {
            DataTable dt = comfun.fillDataTable("stpViksatCompanySelect", "", null);
            return View(dt);
        }
        public JsonResult _SessionAddSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            SortedList list = new SortedList();

            string mes = string.Empty;
            try
            {
                list.Add("@fvSessionName", Request.Form["SessionName"].ToString());
                list.Add("@fdSessionStart", Request.Form["SessionStartDate"].ToString());
                list.Add("@fdSessionEnd", Request.Form["SessionEndDate"].ToString());
                list.Add("@fvSessionType", Request.Form["SessionType"].ToString());

                list.Add("@CompanyID", Request.Form["CompanyId"].ToString());
                mes = comfun.executeNonQueryWMessage("stpviksatSessionAccept", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _SessionAddSubmit", ex.Message);
            }
            return Json(mes);
        }

        public ActionResult getGradeByDesignation()
        {
            ViewData["DisplayType"] = "Select";
            ViewData["ActionResult"] = "GradeByDesignation";
            SortedList list = new SortedList();
            list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
            DataTable dt = comfun.fillDataTable("stpViksatGradeByDesignation", "", list);
            return PartialView("_listOrDropDown", dt);
        }

        public ActionResult BloodGroupList()
        {
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            DataTable dt = comfun.fillDataTable("stpViksatBloodGroup", "", null);
            return View(dt);
        }
        public ActionResult BloodGroupAdd()
        {
            return View();
        }
        public JsonResult _BloodGroupAddSubmit()
        {
            string mes = string.Empty;
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            SortedList list = new SortedList();
            try
            {
                list.Add("@fvBloodGroupName", Request.Form["BloodGroup"].ToString());
                list.Add("@CreatedDate", comfun.dateISTstr());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                mes = comfun.executeNonQueryWMessage("stpViksatBloodgroupAdd", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _BloodGroupAddSubmit", ex.Message);
            }

            return Json(mes);
        }
        public ActionResult BloodGroupEdit()
        {
            string BloodGroupId = "0";
            if (Request.QueryString["id"] == null)
            {
                return RedirectToAction("BloodGroupList");
            }
            BloodGroupId = comfun.decryptString(Request.QueryString["Id"].ToString());

            SortedList list = new SortedList();

            list.Add("@fiBloodGroupId", BloodGroupId);
            DataSet dt = comfun.fillDataSet("stpViksatBloodGroupSelectById", "", list);
            return View(dt);

        }
        public JsonResult _BloodGroupEditSubmit()
        {
            string mes = string.Empty;
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            SortedList list = new SortedList();
            try
            {
                list.Add("@fiBloodgroupId", Request.Form["BloodgroupId"].ToString());
                list.Add("@fvBloodGroupName", Request.Form["BloodGroup"].ToString());
                list.Add("@ModifyDate", comfun.dateISTstr());
                list.Add("@ModifyBy", Session["EmpId"].ToString());
                mes = comfun.executeNonQueryWMessage("stpViksatBloodgroupEdit", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _BloodGroupEditSubmit", ex.Message);
            }
            return Json(mes);
        }

        public ActionResult PayChargesApplied()
        {
            return View();
        }
        public ActionResult PayChargesPartList()
        {
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("stpViksatPayEmployeesPartList", "", null);
            return View(dt);
        }
        public ActionResult PayChargesPart()
        {

            return View();
        }
        public JsonResult _PayChargeSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@PartName", Request.Form["PartName"].ToString());
                list.Add("@PartCode", Request.Form["PartCode"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                list.Add("@CreatedDate", comfun.dateISTstr());
                mes = comfun.executeNonQueryWMessage("StpViksatPayPartSubmit", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _PayChargeSubmit", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult PayChargePartEdit()
        {
            string PartID = "0";
            if (Request.QueryString["id"] == null)
            {
                return RedirectToAction("PayCharges");
            }
            PartID = comfun.decryptString(Request.QueryString["Id"].ToString());
            SortedList list = new SortedList();
            list.Add("@PartID", PartID);
            DataTable dt = comfun.fillDataTable("stpviksatpart_SelectWithId", "", list);
            return View(dt);
        }

        //***********************************************Employee Transfer****************************************************************************//
        //Employee Transfer Details Add//
        public ActionResult EmployeeTransfer()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            // comfun.saveformname("EmployeeTransfer", "/admin/EmployeeTransfer", "Employee Transfer", "Employee Transfer", "Y", "EmployeeTransfer", "EmployeeTransfer", "list");
            DataSet ds = comfun.fillDataSet("EmpTransferToOtherLocation_Display", "", null);
            return View(ds);
        }
        //Employee Transfer Details (Partial View)//
        public ActionResult _EmpDetailsGetById()
        {
            SortedList list = new SortedList();
            list.Add("@EmployeeID", Request.Form["EmployeeId"].ToString());
            DataSet ds = comfun.fillDataSet("EmpTransferToOtherLocation_EmpInfoGetById", "", list);
            return PartialView("_EmpDetailsGetById", ds);
        }
        public JsonResult EmpTransferAddSubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@EmpTransferId", 0);
                list.Add("@EmployeeId", Request.Form["EmployeeID"].ToString());
                list.Add("@CompanyDivisionId", Request.Form["NewCompanyId"].ToString());
                list.Add("@CompanyCostCentreId", Request.Form["NewCostCentreId"].ToString());
                list.Add("@CompanyUnitId", Request.Form["NewUnitId"].ToString());
                list.Add("@CompanyLocationId", Request.Form["NewLocationId"].ToString());
                list.Add("@EmpTransferDepartmentId", Request.Form["NewDepartmentId"].ToString());
                list.Add("@EmpTransferDesignationId", Request.Form["NewDesignationId"].ToString());
                list.Add("@EmpTransferCategoryId", Request.Form["NewEmpCategoryId"].ToString());
                list.Add("@EmpTransferGradeId", Request.Form["NewGradeId"].ToString());
                list.Add("@EmpTransferReportingManagerId", Request.Form["NewReportingManagerId"].ToString());
                list.Add("@NewEmpCodeAssigned", Request.Form["NewEmployeeCode"].ToString());
                list.Add("@EmpTransferDate", Request.Form["EmpTransferDate"].ToString());
                list.Add("@EmpTransferCreateDate", comfun.dateISTstr());
                list.Add("@EmpTransferCreatedBy", Session["EmpId"].ToString());
                list.Add("@EmpTransferModifiedDate", comfun.dateISTstr());
                list.Add("@EmpTransferModifiedBy", Session["EmpId"].ToString());
                list.Add("@OldCostCenterId", Request.Form["CurrentCostCentreId"].ToString());
                list.Add("@OldDivisionId", Request.Form["CurrentCompanyId"].ToString());
                list.Add("@OldUnitId", Request.Form["CurrentUnitId"].ToString());
                list.Add("@OldLocationId", Request.Form["CurrentLocationId"].ToString());
                mes = comfun.executeNonQueryWMessage("EmpTransferToOtherLocation_Save_Update", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Add", ex.Message);
            }
            return Json(mes);
        }
        //Employee Transfer List Display//
        public ActionResult EmployeeTransferList()
        {
            //comfun.saveformname("EmployeeTransferList", "/admin/EmployeeTransferList", "Employee Transfer List", "Employee Transfer List", "Y", "EmployeeTransferList", "EmployeeTransfer", "Display");
            DataTable dt = comfun.fillDataTable("EmpTransferToOtherLocation_DisplayList", "", null);
            return View(dt);
        }
        //************************************************Employee Transfer Details Edit****************************************************************************************************//

        //[customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult EmployeeTransferDetail()
        {
            //comfun.saveformname("EmployeeTransferDetail", "/admin/EmployeeTransferDetail", "Add Employee Transfer Detail", "Add Employee Transfer Detail", "Y", "AddEmployeeTransferDetail", "AddEmployeeTransferDetail", "Add");
            SortedList list = new SortedList();
            list.Add("@EmpId", comfun.decryptString(Request.QueryString["id"]));
            DataSet ds = comfun.fillDataSet("EmployeeTransferDetails_Edit", "", list);
            ViewData["EmpId"] = comfun.decryptString(Request.QueryString["id"]).Trim();
            return View(ds);
        }
        public JsonResult _EmployeeTransferDetailEditSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@EmpTransferDetailId", 0);
                list.Add("@EmployeeID", Request.Form["EmployeeID"].ToString());
                list.Add("@EmployeeCode", Request.Form["EmployeeCode"].ToString());
                list.Add("@EmployeeName", Request.Form["EmployeeName"].ToString());
                list.Add("@FatherHusbandName", Request.Form["FatherHusbandName"].ToString());
                list.Add("@DateofBirth", Request.Form["DateOfBirth"].ToString());
                list.Add("@DateOfJoining", Request.Form["DateOfJoining"].ToString());
                list.Add("@EmpCategoryId", Request.Form["EmpCategoryId"].ToString());
                list.Add("@BranchId", Request.Form["BranchId"].ToString());
                list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
                list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
                list.Add("@Relation", Request.Form["Relation"].ToString());
                list.Add("@Gender", Request.Form["Gender"].ToString());
                list.Add("@BloodGroupId", Request.Form["BloodGroup"].ToString());
                list.Add("@IsMarried", Request.Form["IsMarried"].ToString());
                list.Add("@NoOfChildren ", Request.Form["NoOfChildren"].ToString());
                list.Add("@DateofLeft", Request.Form["DateofLeft"].ToString());
                list.Add("@DateOfResignation", Request.Form["DateOfResignation"].ToString());
                list.Add("@GradeId", Request.Form["GradeID"].ToString());
                list.Add("@IsRegular", Request.Form["IsRegular"].ToString());
                list.Add("@ContractorId", Request.Form["ContractorId"].ToString());
                list.Add("@AadhaarNo", Request.Form["Aadhaar"].ToString());
                list.Add("@Qualification", Request.Form["Qualification"].ToString());
                list.Add("@Experience", Request.Form["Experience"].ToString());
                list.Add("@ReportingManager", Request.Form["ReportingManager"].ToString());
                list.Add("@HRManager", Request.Form["HRManager"].ToString());
                list.Add("@LocationId", Request.Form["LocationId"].ToString());
                list.Add("@UnitId", Request.Form["UnitID"].ToString());
                list.Add("@CompanyId", Request.Form["CompanyID"].ToString());
                list.Add("@DateofMarriage", Request.Form["DOM"].ToString());
                list.Add("@MaritalStatusId", Request.Form["MaritalStatus"].ToString());
                list.Add("@StatusId", Request.Form["Status"].ToString());
                list.Add("@DateofConfirmation", Request.Form["DOC"].ToString());
                list.Add("@EmpTransferModifiedDate", comfun.dateISTstr());
                list.Add("@EmpTransferModifiedBy", Session["EmpId"].ToString());
                //@Mobile
                //@Address1@EmpTransferModifiedDate datetime =null , 

                //@Address2
                //@CityId
                //@Phone
                //@Email
                //@PaymentMode
                //@BankId
                //@BankAccountNo
                //@PFNo
                //@ESINo
                //@EmpImage
                //@EmpImagePath
                mes = comfun.executeNonQueryWMessage("Emp_Transfer_Details_Save_Update", "", list).ToString();
                if (!mes.Contains("Error"))
                {
                    mes = comfun.encryptString(mes);
                }
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("EmployeeTransferDetailEditSubmit", "Error:" + ex.Message);
            }
            return Json(mes);
        }
        //*************************************************Employee Transfer List Report*************************************************************************************************//
        public ActionResult EmployeeTransferListReport()
        {
            //comfun.saveformname("EmployeeTransferListReport", "/admin/EmployeeTransferListReport", "Employee Transfer List Report", "Employee Transfer List Report", "Y", "EmployeeTransferListReport", "EmployeeTransferListReport", "Display");
            return View();
        }
        public ActionResult _GetEmployeeTransferListDisplay()
        {
            SortedList list = new SortedList();
            list.Add("@EmployeeId", Convert.ToInt32(Request.Form["EmployeeId"]));
            list.Add("@Status", Request.Form["Status"].ToString());
            list.Add("@FromDate", Convert.ToDateTime(Request.Form["FromDate"]).ToString("dd-MMM-yyyy"));
            list.Add("@ToDate", Convert.ToDateTime(Request.Form["ToDate"]).ToString("dd-MMM-yyyy"));
            DataTable dt = comfun.fillDataTable("EmployeeTransferListByStatus", "", list);
            return PartialView(dt);
        }
        public JsonResult _TransferApprovalSubmit()
        {
            string mes = string.Empty;
            SortedList list = new SortedList();
            list.Add("@EmployeeId", Convert.ToInt32(Request.Form["EmployeeId"]));
            list.Add("@fvEmpTransferStatus", Request.Form["Status"].ToString());
            list.Add("@EmpTransferId", Request.Form["TransferId"].ToString());
            list.Add("@fdEmpTransferModifiedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy"));
            list.Add("@fiEmpTransferModifiedBy", Session["EmpId"].ToString());
            mes = comfun.executeNonQueryWMessage("Employee_TransferEmployeeSubmit", "", list).ToString();
            return Json(mes);
        }
        //******************************************************************************************************************************************************************************//

        public ActionResult _FamilyDetail()
        {
            if (Request.Form["EmpName"] != null)
            {
                ViewData["EmpName"] = Request.Form["EmpName"].ToString();
                //ViewData["EmpCode"] = Request.Form[""].ToString();
            }
            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmployeeID"].ToString());
            ViewData["EditEmpId"] = Request.Form["EmployeeID"].ToString();
            DataTable dt = comfun.fillDataTable("Employee_FamilyDetailWithEmpId", "", list);
            return PartialView(dt);
            //return PartialView("_employeeFamilyDetail");
        }
        public ActionResult _EmployeeFamilydetailAdd()
        {
            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("StpviksatPayRelationShow", null, null);

            return PartialView("_EmployeeFamilyDetailAdd", dt);
        }
        public JsonResult _EmployeeFamilyDetailAddSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@fiEmployeeId", Request.Form["fiEmployeeId"].ToString());
                list.Add("@Name", Request.Form["Name"].ToString());
                list.Add("@Age", Request.Form["Age"].ToString());
                list.Add("@RelationId", Request.Form["RelationName"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                list.Add("@CreatedDate", comfun.dateISTstr());
                mes = comfun.executeNonQueryWMessage("stpviksatEmployeeFamilyDetailSave", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _EmployeeFamilyDetailAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult _employeeFamilyDetailedit()
        {
            //empcontactid = comfun.decryptString(Request.QueryString["Id"].ToString());

            SortedList list = new SortedList();
            list.Add("@FamilyDetailId", Request.Form["Id"].ToString());
            DataSet ds = comfun.fillDataSet("Employee_FamilyDetailSelectWithId", "", list);
            return PartialView("_employeeFamilyDetailedit", ds);
        }
        public JsonResult _EmployeeFamilyEditSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@fiemployeedetailid", Request.Form["fiemployeedetailid"].ToString());
                list.Add("@Age", Request.Form["Age"].ToString());
                list.Add("@fiRelationId", Request.Form["RelationName"].ToString());
                list.Add("@Name", Request.Form["Name"].ToString());


                mes = comfun.executeNonQueryWMessage("Employee_FamilyDetailEdit", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _EmployeeFamilyEditSubmit", ex.Message);
            }
            return Json(mes);
        }

        public ActionResult ArearShow()
        {
            //comfun.saveformname("ArearShow", "/admin/ArearShow", "Arear List", "Arear List form", "Y", "", "Arear", "List");
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }

            SortedList list = new SortedList();
            list.Add("@CostCenterId", Session["CostCenterId"].ToString());
            DataTable dt = comfun.fillDataTable("ArearPaidShow", "", list);

            return View(dt);

        }
        public ActionResult AddArear()
        {
            //comfun.saveformname("AddArear", "/admin/AddArear", "Arear Add", "Arear Add", "N", "ArearAdd", "Arear", "Add");

            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            return View();
        }
        public JsonResult _ArearSave()
        {

            string Mes = string.Empty;
            //DateTime monthYear = Convert.ToDateTime(Request.Form["ArearDate"]);
            try
            {
                SortedList list = new SortedList();
                list.Add("@fiEmployeeID", Request.Form["fiEmployeeID"].ToString());
                list.Add("@ArearDate", Convert.ToDateTime(Request.Form["ArearDate"]).ToString("dd-MMM-yyyy"));
                list.Add("@NoOfDays", Request.Form["NoOfDays"].ToString());
                list.Add("@Tag", Request.Form["Tag"].ToString());
                list.Add("@CostCenterId", Session["CostCenterId"].ToString());
                //list.Add("@fiMonthYear", monthYear.ToString("MMyyyy"));
                Mes = comfun.executeNonQueryWMessage("ArearPaidSave", "", list).ToString();
            }
            catch (Exception ex)
            {
                Mes = comfun.errorMessage("Admin _ArearSave", ex.Message);
            }
            return Json(Mes);

        }
        public ActionResult ArearEdit()
        {
            //comfun.saveformname("ArearEdit", "/admin/ArearEdit", "Arear Edit", "Arear Edit", "N", "ArearEdit", "Arear", "Edit");
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string ArearId = "0";
            SortedList list = new SortedList();
            ArearId = comfun.decryptString(Request.QueryString["Id"].ToString());
            list.Add("@ID", ArearId);

            DataTable dt = comfun.fillDataTable("ArearPaidShowwithID", "", list);

            return View(dt);

        }
        public JsonResult _ArearEdit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string Mes = string.Empty;
            //  DateTime monthYear = Convert.ToDateTime(Request.Form["ArearDate"]);
            try
            {
                SortedList list = new SortedList();
                list.Add("@ArearID", Request.Form["ID"].ToString());
                list.Add("@ArearDate", Convert.ToDateTime(Request.Form["ArearDate"]).ToString("dd-MMM-yyyy"));
                list.Add("@NoOfDays", Request.Form["NoOfDays"].ToString());

                Mes = comfun.executeNonQueryWMessage("ArearPaidUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                Mes = comfun.errorMessage("Admin _ArearEdit", ex.Message);
            }
            return Json(Mes);

        }

        public ActionResult _DepartmentChange()
        {
            //comfun.saveformname("DepartmentChange", "/admin/DepartmentChange", "Department Change", "Department Change view", "N", "DepartmentChangeAdd", "DepartmentChange", "Add");
            DataTable dt = comfun.fillDataTable("StpViksatDepartment", "", null);
            return PartialView("_DepartmentChange", dt);
        }
        public JsonResult _DepartmentChangeAddSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@fiEmployeeId", Request.Form["EmployeeId"].ToString());
                list.Add("@WEFDate", Request.Form["WEFDate"].ToString());
                list.Add("@fiDepartmentId", Request.Form["DepartmentId"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                list.Add("@CreatedDate", comfun.dateISTstr());
                mes = comfun.executeNonQueryWMessage("stpViksatChangeDepartmentSave", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _DepartmentChangeAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult _GradeBind()
        {
            ViewData["DisplayType"] = "Select";
            ViewData["ActionResult"] = "GradeByDesignation";
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
            dt = comfun.fillDataTable("stpViksatGrade", "", list);
            return PartialView("_listOrDropDown", dt);
        }
        public ActionResult _DesignationChange()
        {
            //comfun.saveformname("_DesignationChange", "/admin/_DesignationChange", "_Designation Change", "_Designation Change view", "N", "_DesignationChange", "_DesignationChange", "Add");
            DataTable dt = comfun.fillDataTable("StpViksatDesignation", "", null);
            return PartialView("_DesignationChange", dt);
        }
        public JsonResult _DesignationChangeAddSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@fiEmployeeId", Request.Form["EmployeeId"].ToString());
                list.Add("@WEFDate", Request.Form["WEFDate"].ToString());
                list.Add("@fiDesignationId", Request.Form["DesignationId"].ToString());
                list.Add("@fiGradeId", Request.Form["GradeId"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                list.Add("@CreatedDate", comfun.dateISTstr());
                mes = comfun.executeNonQueryWMessage("stpViksatChangeDesignationSave", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _DepartmentChangeAddSubmit", ex.Message);
            }
            return Json(mes);
        }

        public ActionResult WorkerCategory()
        {

            if (payfun.sessionRecreate() == "expires")
            {
                return RedirectToAction("Login", "account");
            }
            //comfun.saveformname("WorkerCategoryAdd", "/admin/WorkerCategoryAdd", "Worker Category Add", "Worker Category Add view", "N", "WorkerCategoryAdd", "Worker Category", "Add");
            return View();
        }
        public ActionResult WorkerCategoryList()
        {
            //comfun.saveformname("WorkerCategoryList", "/admin/WorkerCategoryList", "Worker Category List", "Worker Category  view", "N", "WorkerCategoryList", "WorkerCategory", "List");

            DataTable dt = comfun.fillDataTable("stpWorkerCategoryList", "", null);
            return View(dt);
        }
        public ActionResult WorkerCategoryEdit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return RedirectToAction("Login", "account");
            }
            //comfun.saveformname("WorkerCategoryEdit", "/admin/WorkerCategoryEdit", "Worker Category Edit", "Worker Category  view", "N", "WorkerCategoryEdit", "WorkerCategory", "Edit");
            string WorkercategoryId = "0";
            if (Request.QueryString["id"] == null)
            {
                return RedirectToAction("WorkerCategory");
            }
            WorkercategoryId = comfun.decryptString(Request.QueryString["Id"].ToString());
            SortedList list = new SortedList();
            list.Add("@fiWorkercategoryId", WorkercategoryId);
            DataTable dt = comfun.fillDataTable("stpWorkerCategory_SelectWithId", "", list);
            return View(dt);
        }
        public JsonResult _WorkerCategorySubmit()
        {
            string mes = string.Empty;
            try
            {

                SortedList list = new SortedList();
                list.Add("@fvWorkercategoryCode", Request.Form["Workercategorycode"].ToString());
                list.Add("@fvWorkerCategoryName", Request.Form["Workercategory"].ToString());
                list.Add("@fiCreatedBy", Session["EmpId"].ToString());
                list.Add("@fdCreatedDate", comfun.dateISTstr());
                mes = comfun.executeNonQueryWMessage("stpWorkerCategory_Add", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Add", ex.Message);
            }
            return Json(mes);
        }
        public JsonResult _WorkerCategoryEditSubmit()
        {
            string mes = string.Empty;
            try
            {

                SortedList list = new SortedList();
                list.Add("@fvWorkercategoryCode", Request.Form["WorkercategoryCode"].ToString());
                list.Add("@fvWorkerCategoryName", Request.Form["WorkerCategoryName"].ToString());
                list.Add("@fdModifyDate", comfun.dateISTstr());
                list.Add("@fiModifyBy", Session["EmpId"].ToString());
                list.Add("@fiWorkercategoryId", Request.Form["WorkercategoryId"].ToString());


                mes = comfun.executeNonQueryWMessage("stpWorkerCategory_Edit", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Edit", ex.Message);
            }
            return Json(mes);
        }

        public ActionResult InsuranceList()
        {
            //comfun.saveformname("InsuranceList", "/admin/InsuranceList", "Insurance List", "Insurance List", "Y", "InsuranceList", "InsuranceList", "List");


            if (payfun.sessionRecreate() == "expires")
            {
                return RedirectToAction("Login", "account");
            }
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@CostcenterID", Session["CostCenterID"].ToString());
            dt = comfun.fillDataTable("stpViksatInsurancePolicyDetail", "", list);
            return View(dt);
        }
        public ActionResult InsuranceAdd()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return RedirectToAction("Login", "account");
            }
            //comfun.saveformname("InsuranceAdd", "/admin/InsuranceAdd", "Insurance Add", "Insurance Add", "N", "InsuranceAdd", "InsuranceAdd", "Add");

            //if (Request.QueryString["id"] == null)
            //{
            //    return RedirectToAction("InsuranceAdd");
            //}

            return View();
        }
        public JsonResult _InsuranceSubmit()
        {
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@fvPolicyNo", Request.Form["PolicyNo"].ToString());
                list.Add("@fiCostCenterId", Session["CostCenterId"].ToString());
                list.Add("@fvInsuaranceCompanyName", Request.Form["InsuranceCompanyName"].ToString());
                list.Add("@fvInsuranceReferenceNo", Request.Form["InsuranceCompanyrefno"].ToString());
                list.Add("@fvInsuanceAddress", Request.Form["Address"].ToString());
                list.Add("@CreatedDate", comfun.dateISTstr());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                mes = comfun.executeNonQueryWMessage("stpviksatInsuranceDetailSave", "", list).ToString();
            }
            catch (Exception ex)
            {



                mes = comfun.errorMessage("Add", ex.Message);
            }
            return Json(mes);
        }

        //////////////////////////////
        ///Company Nature Master                            (((DONE)))
        /////////////////////////////
        ///

        //public ActionResult AddNewCompanyNature()
        //{
        //    return View();
        //}

        ////To Save/Update Company Nature//
        //public JsonResult CompanyNatureSaveUpdateSubmit()
        //{
        //    string mes = string.Empty;
        //    try
        //    {
        //        SortedList list = new SortedList();
        //        list.Add("@CompanyNatureId", Request.Form["CompanyNatureId"].ToString());
        //        list.Add("@CompanyNature", Request.Form["CompanyNature"].ToString());
        //        mes = comfun.executeNonQueryWMessage("sp_CompanyNature_SaveUpdate", "", list).ToString();
        //    }
        //    catch (Exception ex)
        //    {
        //        mes = ex.Message;
        //    }
        //    return Json(mes);
        //}

        ////To Display Company Nature List//
        //public ActionResult CompanyNatureList()
        //{
        //    DataTable dt = comfun.fillDataTable("sp_CompanyNatureListDisplay", "", null);
        //    return View(dt);
        //}

        ////To Get Company Nature Details Via Id//
        //public ActionResult UpdateCompanyNature()
        //{
        //    string companynatureid = "0";
        //    if (Request.QueryString["id"] == null)
        //    {
        //        return RedirectToAction("CompanyNatureList", "CompanyNatureList");
        //    }
        //    companynatureid = Request.QueryString["Id"].ToString();
        //    SortedList list = new SortedList();
        //    list.Add("@CompanyNatureId", companynatureid);
        //    DataTable dt = comfun.fillDataTable("sp_CompanyNatureListDisplayViaId", "", list);
        //    return View(dt);
        //}

        ////To Delete Existing Company Nature//
        //public JsonResult DeleteCompanyNature()
        //{
        //    string mes = string.Empty;
        //    try
        //    {
        //        SortedList list = new SortedList();
        //        list.Add("@CompanyNatureId", Request.Form["CompanyNatureId"].ToString());
        //        mes = comfun.executeNonQueryWMessage("sp_CompanyNature_Delete", "", list).ToString();
        //    }
        //    catch (Exception ex)
        //    {
        //        mes = ex.Message;
        //    }
        //    return Json(mes, JsonRequestBehavior.AllowGet);
        //}

        ////To Change Active/Deactive Status Company Nature//
        public JsonResult _CompanyNatureStatusUpdate()
        {
            comfun.saveformname("_CompanyNatureStatusUpdate", "/admin/_CompanyNatureStatusUpdate", "Master/Organization/Company Nature", "Company Nature Status", "N", "CompanyNature", "CompanyNature", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_CompanyNature_Active_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }


        //////////////////////////////
        ///Marital Status Master
        /////////////////////////////
        ///

        public ActionResult AddNewMaritalStatus()
        {
            return View();
        }

        //To Save/Update Marital Status//
        public JsonResult MaritalStatusSaveUpdateSubmit()
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
                mes = ex.Message;
            }
            return Json(mes);
        }

        //To Display Marital Status List//
        public ActionResult MaritalStatussList()
        {
            DataTable dt = comfun.fillDataTable("sp_MaritalStatusListDisplay", "", null);
            return View(dt);
        }

        //To Get Marital Status Details Via Id//
        public ActionResult UpdateMaritalStatus()
        {
            string maritalstatusid = "0";
            if (Request.QueryString["id"] == null)
            {
                return RedirectToAction("MaritalStatusList", "MaritalStatusList");
            }
            maritalstatusid = Request.QueryString["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@MaritalStatusId", maritalstatusid);
            DataTable dt = comfun.fillDataTable("sp_MaritalStatusDisplayViaId", "", list);
            return View(dt);
        }

        //To Delete Existing Marital Status//
        public JsonResult DeleteMaritalStatus()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@MaritalStatusId", Request.Form["MaritalStatusId"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_MaritalStatus_Delete", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes, JsonRequestBehavior.AllowGet);
        }

        //To Change Active/Deactive Status Marital Status//
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


        ////////////////////////////////
        /////Employee Category Master ((DONE))
        ///////////////////////////////
        /////

        //public ActionResult AddNewEmployeeCategory()
        //{
        //    return View();
        //}

        ////To Save/Update Employee Category//
        //public JsonResult EmployeeCategorySaveUpdateSubmit()
        //{
        //    string mes = string.Empty;
        //    try
        //    {
        //        SortedList list = new SortedList();
        //        list.Add("@EmpCategoryId", Request.Form["EmpCategoryId"].ToString());
        //        list.Add("@EmpCategoryCode", Request.Form["EmpCategoryCode"].ToString());
        //        list.Add("@EmpCategoryName", Request.Form["EmpCategoryName"].ToString());
        //        list.Add("@IsOverTime", Request.Form["IsOverTime"].ToString());
        //        list.Add("@IsLateComingApplied", Request.Form["IsLateComingApplied"].ToString());
        //        list.Add("@LCHalfDayAbsentMins", Request.Form["LCHalfDayAbsentMins"].ToString());
        //        list.Add("@LCFullDayAbsentMins", Request.Form["LCFullDayAbsentMins"].ToString());
        //        list.Add("@IsContinuousLate", Request.Form["IsContinuousLate"].ToString());
        //        list.Add("@IsWorkDurationApplied", Request.Form["IsWorkDurationApplied"].ToString());
        //        list.Add("@WDHalfDayAbsentMins", Request.Form["WDHalfDayAbsentMins"].ToString());
        //        list.Add("@WDFulldayAbsentMins", Request.Form["WDFulldayAbsentMins"].ToString());
        //        mes = comfun.executeNonQueryWMessage("sp_EmpCategory_SaveUpdate", "", list).ToString();
        //    }
        //    catch (Exception ex)
        //    {
        //        mes = ex.Message;
        //    }
        //    return Json(mes);
        //}

        ////To Display Employee Category List//
        ////public ActionResult EmployeeCategoryList()
        ////{
        ////    DataTable dt = comfun.fillDataTable("sp_EmpCategoryListDisplay", "", null);
        ////    return View(dt);
        ////}

        ////To Get Employee Category Details Via Id//
        //public ActionResult UpdateEmployeeCategory()
        //{
        //    string employeecategoryid = "0";
        //    if (Request.QueryString["id"] == null)
        //    {
        //        return RedirectToAction("EmployeeCategoryList", "EmployeeCategoryList");
        //    }
        //    employeecategoryid = Request.QueryString["Id"].ToString();
        //    SortedList list = new SortedList();
        //    list.Add("@EmpCategoryId", employeecategoryid);
        //    DataTable dt = comfun.fillDataTable("sp_EmpCategoryDisplayViaId", "", list);
        //    return View(dt);
        //}

        ////To Delete Existing Employee Category//
        //public JsonResult DeleteEmployeeCategory()
        //{
        //    string mes = string.Empty;
        //    try
        //    {
        //        SortedList list = new SortedList();
        //        list.Add("@EmpCategoryId", Request.Form["EmpCategoryId"].ToString());
        //        mes = comfun.executeNonQueryWMessage("sp_EmpCategoryDelete", "", list).ToString();
        //    }
        //    catch (Exception ex)
        //    {
        //        mes = ex.Message;
        //    }
        //    return Json(mes, JsonRequestBehavior.AllowGet);
        //}

        ////To Change Active/Deactive Status Employee Category//
        //public JsonResult EmployeeCategoryStatusUpdate()
        //{
        //    string mes = string.Empty;
        //    try
        //    {
        //        SortedList list = new SortedList();
        //        list.Add("@EmpCategoryId", Request.Form["EmpCategoryId"].ToString());
        //        list.Add("@EmpCategoryStatus", Request.Form["EmpCategoryStatus"].ToString());
        //        mes = comfun.executeNonQueryWMessage("sp_EmpCategory_Active_Deactive", "", list).ToString();
        //    }
        //    catch (Exception ex)
        //    {
        //        mes = ex.Message;
        //    }
        //    return Json(mes);
        //}


        //////////////////////////////
        ///Employee Status ((DONE))
        /////////////////////////////
        ///

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

        //To Display Employee Status List//
        //public ActionResult EmpStatusaList()
        //{

        //    DataTable dt = comfun.fillDataTable("sp_EmpStatusListDisplay", "", null);
        //    return View(dt);
        //}

        //To Get Employee Status Details Via Id//
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



        //////////////////////////////
        ///Policy Category ((DONE))
        /////////////////////////////
        ///

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




        /// <summary>
        /// Conveyance Auto Setting ((DONE))
        /// </summary>
        /// <returns></returns>











        #endregion
        #region Munish


        public ActionResult _Pageno()
        {
            DataTable dt = comfun.fillDataTable("stpPageMaster", null, null);

            return PartialView("_PageNo", dt);
        }
        //Zone//





        //Country//
        public ActionResult Country()
        {
            comfun.saveformname("Country", "/admin/Country", "Master/Address/Country", "Country  Main form", "Y", "Country", "Country", "View", 1);
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
            comfun.saveformname("_countryList", "/admin/_countryList", "Master/Address/Country", "Country  list", "N", "Country", "Country", "List", 4);
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
            comfun.saveformname("_CountrySubmit", "/admin/_CountrySubmit", "Master/Address/Country", "Country  add", "N", "Country", "Country", "Add", 2);
            comfun.saveformname("_CountrySubmit", "/admin/_CountrySubmit", "Master/Address/Country", "Country  edit", "N", "Country", "Country", "Edit", 3);

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
            comfun.saveformname("_CountryStatusUpdate", "/admin/_CountryStatusUpdate", "Master/Address/Country", "Country  Status", "N", "Country", "Country", "Status", 5);
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
            comfun.saveformname("State", "/admin/States", "Master/Address/States", "State  Main Form", "Y", "State", "State", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("State");
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            DataTable dt = comfun.fillDataTable("CountryState_Select", "", null);
            return View(dt);
        }
        public ActionResult _state_selectddl()
        {
            SortedList list = new SortedList();
            //  list.Add("@CountryId", Request.Form["CountrStateIdyId"].ToString());
            DataTable dt = comfun.fillDataTable("State_SelectByCountryId", "", list);
            return PartialView("_state_ddl", dt);
        }
        public JsonResult _StateStatusUpdate()
        {
            comfun.saveformname("_StateStatusUpdate", "/admin/_StateStatusUpdate", "Master/Address/States", "State  Status", "N", "State", "State", "Status", 5);

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
            comfun.saveformname("_stateListJson", "/admin/_stateListJson", "Master/Address/States", "State List", "N", "State", "State", "List", 4);
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
            comfun.saveformname("_StateSubmit", "/admin/_StateSubmit", "Master/Address/States", "States Add", "N", "State", "State", "Add", 2);
            comfun.saveformname("_StateSubmit", "/admin/_StateSubmit", "Master/Address/States", "States Edit", "N", "State", "State", "Edit", 3);
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
        public ActionResult Cities()
        {
            comfun.saveformname("City", "/admin/Cities", "Master/Address/City", "City  Main Form", "Y", "City", "City", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("City");

            DataTable dt = comfun.fillDataTable("CountryState_Select", "", null);
            return View(dt);
        }
        //To Display City List Via State Id//
        public ActionResult _city_select()
        {
            comfun.saveformname("_city_select", "/admin/_city_select", "Master/Address/City", "City  List", "N", "City", "City", "List", 4);
            SortedList list = new SortedList();
            list.Add("@StateId", Request.Form["StateId"].ToString());
            DataTable dt = comfun.fillDataTable("City_SelectByStateId", "", list);
            ViewData["StateId"] = Request.Form["StateId"].ToString();

            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        //For City Save/Update Submit//
        public JsonResult _CitiesSaveUpdateSubmit()
        {
            comfun.saveformname("_CitiesSaveUpdateSubmit", "/admin/_CitiesSaveUpdateSubmit", "Master/Address/City", "City  Add", "N", "City", "City", "Add", 2);
            comfun.saveformname("_CitiesSaveUpdateSubmit", "/admin/_CitiesSaveUpdateSubmit", "Master/Address/City", "City  Edit", "N", "City", "City", "Edit", 3);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@CityId", Request.Form["CityId"].ToString());
                list.Add("@CityCode", Request.Form["CityCode"].ToString());
                list.Add("@CityName", Request.Form["CityName"].ToString());
                list.Add("@StateId", Request.Form["StateId"].ToString());
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
        //Zone//
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
            comfun.saveformname("BankMaster", "/admin/Banks", "Master/Bank Detail/Bank", "Bank Master form", "Y", "BankMaster", "BankMaster", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("BankMaster");
            return View();
        }
        public ActionResult _BankList()
        {
            comfun.saveformname("_BankList", "/admin/Banks", "Master/Bank Detail/Bank", "Bank Master List", "N", "BankMaster", "BankMaster", "List", 2);
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
            comfun.saveformname("_BanksAddSubmit", "/admin/_BanksAddSubmit", "Master/Bank Detail/Bank", "Bank Master Add", "N", "BankMaster", "BankMaster", "Add", 2);
            comfun.saveformname("_BankEditSubmit", "/admin/_BankEditSubmit", "Master/Bank Detail/Bank", "Bank Master Edit", "N", "BankMaster", "BankMaster", "Edit", 3);
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@BankName", Request.Form["BankName"].ToString());
                list.Add("@BankCode", Request.Form["BankCode"].ToString());
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
            comfun.saveformname("_BankStatusUpdate", "/admin/_BankStatusUpdate", "Master/Bank Detail/Bank", "Bank Master Status", "N", "BankMaster", "BankMaster", "Status", 5);

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
            comfun.saveformname("Level", "/admin/Level", "Master/Clearance/Level", "Level  Main form", "Y", "Level", "Level", "View", 1);
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
            comfun.saveformname("_LevelList", "/admin/_LevelList", "Master/Clearance/Level", "Level  list", "N", "Level", "Level", "List", 4);


            SortedList list = new SortedList();

            DataTable dt = comfun.fillDataTable("stpHtisLevelsList", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public JsonResult _LevelSubmit()
        {
            comfun.saveformname("_LevelSubmit", "/admin/_LevelSubmit", "Master/Clearance/Level", "Level  add", "N", "Level", "Level", "Add", 2);
            comfun.saveformname("_LevelSubmit", "/admin/_LevelSubmit", "Master/Clearance/Level", "Level  edit", "N", "Level", "Level", "Edit", 3);

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@Title", Request.Form["Title"].ToString());
                list.Add("@ScheduleTime", Request.Form["ScheduleTime"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                list.Add("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                mes = comfun.executeNonQueryWMessage("stpLevel_Save", "", list).ToString();
                if (mes.Contains("Error"))
                {
                    list.Clear();
                    list.Add("Error", mes);
                }
                else
                {
                    list.Clear();
                    list.Add("Error", "");
                    list.Add("Id", mes);
                    list.Add("Title", Request.Form["Title"].ToString());
                    list.Add("ScheduleTime", Request.Form["ScheduleTime"].ToString());
                }
            }
            catch (Exception ex)
            {
                list.Clear();
                list.Add("Error", mes);
            }
            return Json(list);
        }
        public ActionResult _LevelStatusUpdate()
        {
            comfun.saveformname("_LevelStatusUpdate", "/admin/_LevelStatusUpdate", "Master/Clearance/Level", "Level  Status", "N", "Level", "Level", "Status", 4);

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
        public ActionResult department()
        {
            comfun.saveformname("department", "/admin/department", "Master/General/Department", "Department Main form", "Y", "Department", "Department", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("department");

            return View();
        }
        public ActionResult _departmentList()
        {
            // comfun.saveformname("_departmentList", "/admin/_departmentList", "Master/General/Department", "Department  List", "N", "Department", "Department", "List");
            comfun.saveformname("_departmentList", "/admin/_departmentList", "Master/General/Department", "Department List", "N", "Department", "Department", "List", 4);
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
            DataTable dt = comfun.fillDataTable("Department_Select", "", list);
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "admin", "_departmentList", "_departmentList");

            HtmlString htm = new HtmlString(paging);
            ViewData["paging"] = htm;
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        //  [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult _departmentAdd()
        {

            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return RedirectToAction("Login", "account");
            //}
            //comfun.saveformname("departmentAdd", "/admin/departmentAdd", "Department Add", "Department add view", "N", "departmentAdd", "Department", "Add");
            return PartialView("_departmentAdd");
        }
        // [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult _departmentEdit()
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return RedirectToAction("Login", "account");
            //}
            //comfun.saveformname("departmentEdit", "/admin/departmentEdit", "Department Edit", "Department edit view", "N", "departmentEdit", "Department", "Edit");
            string departmentId = "0";
            //if (Request.QueryString["id"] == null)
            //{
            //    return RedirectToAction("department");
            //}
            departmentId = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@DepartmentId", departmentId);

            DataTable dt = comfun.fillDataTable("Department_SelectWithId", "", list);
            return PartialView(dt);
        }
        public JsonResult _departmentAddSubmit()
        {
            string mes = string.Empty;
            try
            {
                comfun.saveformname("_departmentEditSubmit", "/admin/_departmentEditSubmit", "Master/General/Department", "Department Edit", "N", "Department", "Department", "Edit", 3);
                comfun.saveformname("_departmentAddSubmit", "/admin/_departmentAddSubmit", "Master/General/Department", "Department Add", "N", "Department", "Department", "Add", 2);
                //comfun.saveformname("_departmentAddSubmit", "/admin/_departmentAddSubmit", "Master/General/Department", "Department  Add", "N", "Department", "Department", "Add");
                SortedList list = new SortedList();
                list.Add("@DepartmentName", Request.Form["DepartmentName"].ToString());
                list.Add("@DepartmentCode", Request.Form["DepartmentCode"].ToString());
                list.Add("@CreatedBy", "1");
                list.Add("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                mes = comfun.executeNonQueryWMessage("Department_save", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Add", ex.Message);
            }
            return Json(mes);
        }
        public JsonResult _UpdateStatusDepartment()
        {
            comfun.saveformname("_UpdateStatusDepartment", "/admin/_UpdateStatusDepartment", "Master/General/Department", "Department Status", "N", "Department", "Department", "Status", 5);

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpUpdateStatusDepartment", "", list).ToString();
            }
            catch (Exception ex)
            {

                mes = ex.Message;
            }

            return Json(mes);
        }
        public JsonResult _departmentEditSubmit()
        {
            // comfun.saveformname("_departmentEditSubmit", "/admin/_departmentEditSubmit", "Master/General/Department", "Department  Add", "N", "Department", "Department", "Edit");
            string mes = string.Empty;
            try
            {

                SortedList list = new SortedList();
                list.Add("@DepartmentName", Request.Form["DepartmentName"].ToString());
                list.Add("@DepartmentCode", Request.Form["DepartmentCode"].ToString());
                list.Add("DepartmentId", Request.Form["DepartmentId"].ToString());

                mes = comfun.executeNonQueryWMessage("Department_Edit", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Edit", ex.Message);
            }
            return Json(mes);
        }
        public ActionResult Designation()
        {
            comfun.saveformname("Designation", "/admin/Designation", "Master/General/Designation", "Designation Main form", "Y", "Designation", "Designation", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("Designation");
            return View();
        }
        public ActionResult _designationList()
        {
            comfun.saveformname("_designationList", "/admin/_designationList", "Master/General/Designation", "Designation List", "N", "Designation", "Designation", "List", 4);
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
            comfun.saveformname("_DesignationStatusUpdate", "/admin/_DesignationStatusUpdate", "Master/General/Designation", "Designation Status", "N", "Designation", "Designation", "Status", 5);

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
            comfun.saveformname("_designationEditSubmit", "/admin/_designationEditSubmit", "Master/General/Designation", "Designation Edit", "N", "Designation", "Designation", "Edit", 3);
            comfun.saveformname("_designationAddSubmit", "/admin/_designationAddSubmit", "Master/General/Designation", "Designation Add", "N", "Designation", "Designation", "Add", 2);
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
            comfun.saveformname("_designationEditSubmit", "/admin/_designationEditSubmit", "Master/General/Designation", "Designation Edit", "N", "Designation", "Designation", "Edit", 3);
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
            comfun.saveformname("CompanyNature", "/admin/CompanyNatureList", "Master/Organization/Company Nature", "Company Nature Main form", "Y", "CompanyNature", "CompanyNature", "View", 1);
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
            comfun.saveformname("_companyNatureList", "/admin/_companyNatureList", "Master/Organization/Company Nature", "Company Nature List", "N", "CompanyNature", "CompanyNature", "List", 4);
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
            comfun.saveformname("_CompnayNatureEditSubmit", "/admin/_CompnayNatureEditSubmit", "Master/Organization/Company Nature", "Company Nature Add", "N", "CompanyNature", "CompanyNature", "Add", 2);
            comfun.saveformname("_CompnayNatureEditSubmit", "/admin/_CompnayNatureEditSubmit", "Master/Organization/Company Nature", "Company Nature Edit", "N", "CompanyNature", "CompanyNature", "Edit", 3);
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
            comfun.saveformname("ConveyanceSettingList", "/admin/ConveyanceSettingList", "Master/Conveyance&nbsp;/Auto setting", "Conveyance Setting form", "Y", "ConveyanceSetting", "ConveyanceSetting", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("ConveyanceSettingList");
            return View();
        }
        public ActionResult _conveyanceSettingList()
        {
            comfun.saveformname("ConveyanceSettingList", "/admin/_conveyanceSettingList", "Master/Conveyance&nbsp;/Auto setting", "Conveyance Setting List", "N", "ConveyanceSetting", "ConveyanceSetting", "List", 4);

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
            comfun.saveformname("ConveyanceSettingList", "/admin/_ConveyanceSettingAddSubmit", "Master/Conveyance&nbsp;/Auto setting", "Conveyance Setting Add", "N", "ConveyanceSetting", "ConveyanceSetting", "Add", 2);

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
            comfun.saveformname("_ConveyanceSettingStatusUpdate", "/admin/_ConveyanceSettingStatusUpdate", "Master/Conveyance&nbsp;/Auto setting", "Conveyance Setting Status", "N", "ConveyanceSetting", "ConveyanceSetting", "Status", 5);

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
            comfun.saveformname("ConveyanceSettingList", "/admin/ConveyanceAutoSettingSaveUpdateSubmit", "Master/Conveyance&nbsp;/Auto setting", "Conveyance Setting Add", "N", "ConveyanceSetting", "ConveyanceSetting", "Add", 2);
            comfun.saveformname("ConveyanceSettingList", "/admin/ConveyanceAutoSettingSaveUpdateSubmit", "Master/Conveyance&nbsp;/Auto setting", "Conveyance Setting Edit", "N", "ConveyanceSetting", "ConveyanceSetting", "Edit", 3);

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
            comfun.saveformname("_CityStatusUpdate", "/admin/_CityStatusUpdate", "Master/Address/City", "City  Status", "N", "City", "City", "Status", 5);

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
            comfun.saveformname("Branch", "/admin/Branch", "Master/Organization/Branch&nbsp;", "Branch form", "Y", "Branch", "Branch", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("Branch");
            SortedList list = new SortedList();
            list.Add("@Action", "Add");
            DataSet ds = comfun.fillDataSet("stpViksatBranchCompanyStateCity", "", list);
            return View(ds);
        }
        public JsonResult _branchList()
        {
            comfun.saveformname("_branchList", "/admin/_branchList", "Master/Organization/Branch&nbsp;", "Branch List", "N", "Branch", "Branch", "List", 4);
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
            comfun.saveformname("_BranchAddSubmit", "/admin/_BranchAddSubmit", "Master/Organization/Branch&nbsp;", "Branch Add", "N", "Branch", "Branch", "Add", 2);
            comfun.saveformname("_BranchAddSubmit", "/admin/_BranchAddSubmit", "Master/Organization/Branch&nbsp;", "Branch Edit", "N", "Branch", "Branch", "Edit", 3);
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
            comfun.saveformname("_CompanyBranchStatusUpdate", "/admin/_CompanyBranchStatusUpdate", "Master/Organization/Branch&nbsp;", "Branch Status", "N", "Branch", "Branch", "Status", 5);

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
            comfun.saveformname("_DivisionList", "/admin/_DivisionList", "Master/Organization/Company List", "Company Main form", "Y", "CompanyList", "CompanyList", "List", 4);
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
            comfun.saveformname("_DivisionAddSubmit", "/admin/_DivisionAddSubmit", "Master/Organization/Company List", "Company Add", "N", "CompanyList", "CompanyList", "Add", 2);
            comfun.saveformname("_DivisionAddSubmit", "/admin/_DivisionAddSubmit", "Master/Organization/Company List", "Company Edit", "N", "CompanyList", "CompanyList", "Edit", 3);
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
            comfun.saveformname("_CompanyStatusUpdate", "/admin/_CompanyStatusUpdate", "Master/Organization/Company Status", "Company Status", "N", "CompanyList", "CompanyList", "Status", 5);

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
            comfun.saveformname("GradeList", "/admin/GradeList", "Master/General/GradeList", "Grade List form", "Y", "GradeList", "GradeList", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("GradeList");
            return View();
        }

        public ActionResult _GradeList()
        {
            comfun.saveformname("_GradeList", "/admin/_GradeList", "Master/General/GradeList", "Grade List List", "N", "GradeList", "GradeList", "List", 4);
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
            comfun.saveformname("_GradesAddSubmit", "/admin/_GradesAddSubmit", "Master/General/GradeList", "Grade List Add", "N", "GradeList", "GradeList", "Add", 2);
            comfun.saveformname("_GradeEditSubmit", "/admin/_GradeEditSubmit", "Master/General/GradeList", "Grade List Edit", "N", "GradeList", "GradeList", "Edit", 3);
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
            comfun.saveformname("_GradeStatusUpdate", "/admin/_GradeStatusUpdate", "Master/General/GradeList", "Grade List Status", "Y", "GradeList", "GradeList", "Status", 5);

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
            comfun.saveformname("EmployeeType", "/admin/EmptypeList", "Master/General/Employee Type", "Employee Type form", "Y", "EmployeeType", "EmployeeType", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("EmployeeType");
            return View();
        }
        public ActionResult _EmptypeList()
        {
            comfun.saveformname("_EmptypeList", "/admin/_EmptypeList", "Master/General/Employee Type", "Employee Type List", "N", "EmployeeType", "EmployeeType", "List", 4);
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
            comfun.saveformname("_EmptypeAddSubmit", "/admin/_EmptypeAddSubmit", "Master/General/Employee Type", "Employee Type Add", "N", "EmployeeType", "EmployeeType", "Add", 2);
            comfun.saveformname("_EmptypeEditSubmit", "/admin/_EmptypeEditSubmit", "Master/General/Employee Type", "Employee Type Edit", "N", "EmployeeType", "EmployeeType", "Edit", 3);
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
            comfun.saveformname("_EmptypeStatusUpdate", "/admin/_EmptypeStatusUpdate", "Master/General/Employee Type", "Employee Type Status", "N", "EmployeeType", "EmployeeType", "Status", 5);
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
            comfun.saveformname("EmployeeStatus", "/admin/EmpStatusList", "Master/General/Employee Status", "Employee Status form", "Y", "EmployeeStatus", "EmployeeStatus", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("EmployeeStatus");
            return View();
        }
        public ActionResult _EmpstatusList()
        {
            comfun.saveformname("_EmpstatusList", "/admin/_EmpstatusList", "Master/General/Employee Status", "Employee Status List", "N", "EmployeeStatus", "EmployeeStatus", "List", 4);
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
            comfun.saveformname("_EmpstatusAddSubmit", "/admin/_EmpstatusAddSubmit", "Master/General/Employee Status", "Employee Status Add", "N", "EmployeeStatus", "EmployeeStatus", "Add", 2);
            comfun.saveformname("_EmpstatusEditSubmit", "/admin/_EmpstatusEditSubmit", "Master/General/Employee Status", "Employee Status Edit", "N", "EmployeeStatus", "EmployeeStatus", "Edit", 3);
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
            comfun.saveformname("_EmpStatusUpdate", "/admin/_EmpStatusUpdate", "Master/General/Employee Status", "Employee Status", "N", "EmployeeStatus", "EmployeeStatus", "Status", 5);

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
            comfun.saveformname("MaritalStatus", "/admin/MaritalStatusList", "Master/General/Marital Status", "Marital Status form", "Y", "MaritalStatus", "MaritalStatus", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("MaritalStatus");
            return View();
        }
        public ActionResult _MaritalstatusList()
        {
            comfun.saveformname("_MaritalstatusList", "/admin/_MaritalstatusList", "Master/General/Marital Status", "Marital Status List", "N", "MaritalStatus", "MaritalStatus", "List", 4);
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
            comfun.saveformname("_MaritalstatusAddSubmit", "/admin/_MaritalstatusAddSubmit", "Master/General/Marital Status", "Marital Status Add", "N", "MaritalStatus", "MaritalStatus", "Add", 2);
            comfun.saveformname("_MaritalstatusEdit", "/admin/_MaritalstatusEdit", "Master/General/Marital Status", "Marital Status Edit", "N", "MaritalStatus", "MaritalStatus", "Edit", 3);
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
            comfun.saveformname("_MaritalStatusUpdate", "/admin/_MaritalStatusUpdate", "Master/General/Marital Status", "Marital Status ", "N", "MaritalStatus", "MaritalStatus", "Status", 5);

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
            comfun.saveformname("PolicyCategory", "/admin/PolicyCategoryList", "Master/General/Policy Category", "Policy Category form", "Y", "PolicyCategory", "PolicyCategory", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("PolicyCategory");
            return View();
        }
        public ActionResult _PolicyCategoryList()
        {
            comfun.saveformname("_PolicyCategoryList", "/admin/_PolicyCategoryList", "Master/General/Policy Category", "Policy Category form", "N", "PolicyCategory", "PolicyCategory", "List", 4);
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
            comfun.saveformname("_PolicyCategoryAddSubmit", "/admin/_PolicyCategoryAddSubmit", "Master/General/Policy Category", "Policy Category Add", "N", "PolicyCategory", "PolicyCategory", "Add", 2);
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
            comfun.saveformname("_PolicyCategoryEditSubmit", "/admin/_PolicyCategoryEditSubmit", "Master/General/Policy Category", "Policy Category Edit", "N", "PolicyCategory", "PolicyCategory", "Edit", 3);

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
            comfun.saveformname("PolicyCategoryStatusUpdate", "/admin/PolicyCategoryStatusUpdate", "Master/General/Policy Category", "Policy Category Status", "N", "PolicyCategory", "PolicyCategory", "Status", 5);

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
            comfun.saveformname("ShiftAssign", "/admin/ShiftAssign", "Shift Management/Shift Assign", "Shift Assign form", "Y", "ShiftAssign", "ShiftAssign", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("ShiftAssign");
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

        public JsonResult _ShiftAssignJsonList()
        {
            comfun.saveformname("_ShiftAssignJsonList", "/admin/_ShiftAssignJsonList", "Shift Management/Shift Assign", "Shift Assign List", "N", "ShiftAssign", "ShiftAssign", "List", 4);

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
            comfun.saveformname("_ShiftAssignSubmit", "/admin/_ShiftAssignSubmit", "Shift Management/Shift Assign", "Shift Assign  add", "N", "ShiftAssign", "ShiftAssign", "Add", 2);

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

        public ActionResult _ShiftAssignBulkUpload()
        {

            _exportExcel("Shift Assign Upload");
            return View();
        }

        public void _exportExcel(string excelName)
        {
            string date = DateTime.UtcNow.AddMinutes(330).ToString("yyyyMMdd_HHmm");
            excelName = excelName + "_" + date;
            DataTable dt = new DataTable();
            dt.Columns.Add("S.No.", typeof(int));

            dt.Columns.Add("Employee Id", typeof(string));
            dt.Columns.Add("Employee Name", typeof(string));
            dt.Columns.Add("Shift", typeof(string));
            dt.Columns.Add("From Date", typeof(string));
            dt.Columns.Add("To Date", typeof(string));


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


        public ActionResult _ImportFile()
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

            var postedFile = System.Web.HttpContext.Current.Request.Files["ExcelFile"];
            List<string> data = new List<string>();
            if (postedFile != null)
            {

                if (postedFile.ContentType == "application/vnd.ms-excel" || postedFile.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                {
                    HttpPostedFileBase fileBase;
                    DataTable dataTable = new DataTable();
                    if (Request.Files.Count > 0)
                    {
                        foreach (string file in Request.Files)
                        {
                            //DataTable dtshift = comfun.fillDataTable("", "select * from tblshiftmaster where ID='120006' order by InTime", null);

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

                    var json = JsonConvert.SerializeObject(resultCount);

                    return Json(json);

                }


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

            dtExcel.Columns.Add("SNo", typeof(int));
            dtExcel.Columns.Add("EmployeeId", typeof(string));
            dtExcel.Columns.Add("EmployeeName", typeof(string));
            dtExcel.Columns.Add("Shiftname", typeof(string));
            dtExcel.Columns.Add("FromDate", typeof(string));
            dtExcel.Columns.Add("ToDate", typeof(string));



            DataColumnCollection columns = importdt.Columns;
            DataTable dt = importdt;
            DataRow dr;
            for (int i = 0; i < dt.Rows.Count; i++)
            {

                if (dt.Rows[i][0].ToString() != " ")
                {
                    dr = dtExcel.NewRow();

                    if (dt.Rows[i]["S.No."] != System.DBNull.Value)
                    {
                        dr["SNo"] = dt.Rows[i]["S.No."].ToString();

                    }
                    else
                    {
                        dr["SNo"] = "0";
                    }
                    if (dt.Rows[i]["Employee Id"] != System.DBNull.Value)
                    {
                        dr["EmployeeId"] = dt.Rows[i]["Employee Id"].ToString();
                    }
                    else
                    {
                        dr["EmployeeId"] = "";
                    }
                    if (dt.Rows[i]["Employee Name"] != System.DBNull.Value)
                    {
                        dr["EmployeeName"] = dt.Rows[i]["Employee Name"].ToString();
                    }
                    else
                    {
                        dr["EmployeeName"] = "";
                    }

                    if (dt.Rows[i]["Shift"] != System.DBNull.Value)
                    {
                        dr["Shiftname"] = dt.Rows[i]["Shift"].ToString();
                    }
                    else
                    {
                        dr["Shiftname"] = "";
                    }

                    if (dt.Rows[i]["From Date"] != System.DBNull.Value)
                    {
                        dr["FromDate"] = dt.Rows[i]["From Date"].ToString();


                    }
                    else
                    {

                        dr["FromDate"] = "";
                    }

                    if (dt.Rows[i]["To Date"] != System.DBNull.Value)
                    {
                        dr["ToDate"] = dt.Rows[i]["To Date"].ToString();


                    }
                    else
                    {

                        dr["ToDate"] = "";
                    }


                    dtExcel.Rows.Add(dr);

                }

            }
            return dtExcel;
        }

        public JsonResult _InsertDetail()
        {
            string mes = string.Empty;

            try
            {
                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject(Request.Form["Table"].ToString());
                DataTable table = new DataTable();
                table.Columns.Add("EmpId", typeof(int));
                table.Columns.Add("ShiftId", typeof(int));
                table.Columns.Add("FromDate", typeof(string));
                table.Columns.Add("toDate", typeof(string));

                foreach (var item in jsonData)
                {
                    DataTable dtshift = comfun.fillDataTable("", "select  top 1 ID from tblshiftmaster where Shiftname ='" + item.Shift + "'", null);
                    DataRow dr = table.NewRow();
                    dr["EmpId"] = item.EmployeeId;
                    dr["ShiftId"] = dtshift.Rows[0]["ID"];
                    dr["FromDate"] = Convert.ToDateTime(item.FromDate.ToString()).ToString("yyyy-MM-dd");
                    dr["toDate"] = Convert.ToDateTime(item.ToDate.ToString()).ToString("yyyy-MM-dd");
                    table.Rows.Add(dr);
                }


                SqlCommand com = new SqlCommand();
                connection conObj = new connection();

                com.Connection = conObj.con;

                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "stpShiftAssignBulk_upload";
                com.Parameters.AddWithValue("@loginId", Session["EmpId"]);
                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@ShiftbulkEmp";
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
                mes = comfun.errorMessage("List", "Error: " + ex.Message);
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
            comfun.saveformname("ShiftGroup", "/admin/ShiftGroup", "Shift Management/Shift Group", "Shift Group form", "Y", "ShiftGroup", "ShiftGroup", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("ShiftGroup");

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
            comfun.saveformname("_ShiftGroupJsonList", "/admin/_ShiftGroupJsonList", "Shift Management/Shift Group", "Shift Group List", "N", "ShiftGroup", "ShiftGroup", "List", 4);

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
            comfun.saveformname("_ShiftGroupSubmit", "/admin/_ShiftGroupSubmit", "Shift Management/Shift Group", "Shift Group  add", "N", "ShiftGroup", "ShiftGroup", "Add", 2);
            comfun.saveformname("_ShiftGroupSubmit", "/admin/_ShiftGroupSubmit", "Shift Management/Shift Group", "Shift Group  Edit", "N", "ShiftGroup", "ShiftGroup", "Edit", 3);

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
            comfun.saveformname("_ShiftGroupStatusUpdate", "/admin/_ShiftGroupStatusUpdate", "Shift Management/Shift Group", "Shift Group  Status", "N", "ShiftGroup", "ShiftGroup", "Status", 5);

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
            comfun.saveformname("ShiftMaster", "/admin/ShiftMasterDisplay", "Shift Management/Shift Master", "Shift Master form", "Y", "ShiftMaster", "ShiftMaster", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("ShiftMaster");
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            return View();
        }
        public ActionResult _ShiftMasterJsonList()
        {
            comfun.saveformname("ShiftMaster", "/admin/_ShiftMasterJsonList", "Shift Management/ShiftMaster", "Shift Master List", "N", "ShiftMaster", "ShiftMaster", "List", 4);

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
            comfun.saveformname("_ShiftMasterSubmit", "/admin/_ShiftMasterSubmit", "Shift Management/ShiftMaster", "Shift Master  add", "N", "ShiftMaster", "ShiftMaster", "Add", 2);
            comfun.saveformname("_ShiftMasterSubmit", "/admin/_ShiftMasterSubmit", "Shift Management/ShiftMaster", "Shift Master  edit", "N", "ShiftMaster", "ShiftMaster", "Edit", 3);


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
            comfun.saveformname("_ShiftMasterStatusUpdate", "/admin/_ShiftMasterStatusUpdate", "Shift Management/Shift Master", "Shift Master form", "N", "ShiftMaster", "ShiftMaster", "Status", 5);

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
        //public JsonResult _ShiftMasterEditsubmit()
        //{
        //    if (payfun.sessionRecreate() == "expires")
        //    {
        //        return Json("Session expires");
        //    }
        //    string mes = string.Empty;
        //    try
        //    {
        //        SortedList list = new SortedList();
        //        list.Add("@ShiftName", Request.Form["ShiftName"].ToString());
        //        list.Add("@Id", Request.Form["Id"].ToString());
        //        mes = comfun.executeNonQueryWMessage("ShiftName_Edit", "", list).ToString();
        //    }
        //    catch (Exception ex)
        //    {
        //        mes = comfun.errorMessage("Admin _ShiftMasterEditsubmit", ex.Message);
        //    }
        //    return Json(mes);
        //}




        #endregion
        #region Project Management system

        #endregion
        public ActionResult DocumentCatgoryList()
        {
            comfun.saveformname("DocumentCatgoryList", "/admin/DocumentCatgoryList", "DMS/Catgory", "Document Catgory form", "Y", "DocumentCatgory", "DocumentCatgory", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("DocumentCatgoryList");
            return View();
        }
        public JsonResult _DocumentCategoryList()
        {
            comfun.saveformname("_DocumentCategoryList", "/admin/_DocumentCategoryList", "DMS/Catgory", "Document Catgory List", "N", "DocumentCatgory", "DocumentCatgory", "List", 4);

            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("stpDocumentcategoryList", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _DocumentCategorySubmit()
        {
            comfun.saveformname("_DocumentCategorySubmit", "/admin/_DocumentCategorySubmit", "DMS/Catgory", "Document Catgory Add", "N", "DocumentCatgory", "DocumentCatgory", "Add", 2);
            comfun.saveformname("_DocumentCategorySubmit", "/admin/_DocumentCategorySubmit", "DMS/Catgory", "Document Catgory Edit", "N", "DocumentCatgory", "DocumentCatgory", "Edit", 3);

            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@DocumentCategoryId", Request.Form["DocumentCategoryId"].ToString());
                list.Add("@DocumentCategory", Request.Form["DocumentCategory"].ToString());
                list.Add("@DocumentCategoryColor", Request.Form["DocumentCategoryColor"].ToString());
                mes = comfun.executeNonQueryWMessage("stpDocumentCategoryAcceptUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {

                mes = ex.Message;
            }

            return Json(mes);
        }
        public JsonResult _DocumentCategoryStatusUpdate()
        {
            comfun.saveformname("_DocumentCategoryStatusUpdate", "/admin/_DocumentCategoryStatusUpdate", "DMS/Catgory", "Document Catgory Status", "N", "DocumentCatgory", "DocumentCatgory", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_DocumentCategory_Active_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }
        public JsonResult _DocumentCategoryViewUpdate()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsUserView", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("sp_DocumentCategory_View_Deactive", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }
        public ActionResult DocumentRepository()
        {
            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("stpEmployeedropdown", null, null);
            return View(dt);
        }
        public ActionResult RepositoryDocumentList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@EmpId", "0");
            dt = comfun.fillDataTable("stpDocumentCategoryEmployeeList", null, list);
            return PartialView("_DocumentCategoryList", dt);
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

                            ImageName = "DMS" + fname + Request.Form["Category"].ToString() + Request.Form["UserId"].ToString() + _ext;
                            string filePath = Path.Combine(Server.MapPath("~/DMS/") + ImageName);

                            fname = filePath;
                            var filePathA = filePath;
                            //  file.SaveAs(filePathA);
                            file.SaveAs(filePathA);
                            SortedList list = new SortedList();
                            list.Add("@DocumentCategoryId", Request.Form["Category"].ToString());
                            list.Add("@fiEmployeeId", Request.Form["UserId"].ToString());
                            list.Add("@LoginId", Session["EmpId"].ToString());
                            list.Add("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                            list.Add("@fvFileName", ImageName);
                            list.Add("@fvFilePath", filePathA);



                            mes = comfun.executeNonQueryWMessage("stpDocumentRepository_Accept", "", list).ToString();
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
        public JsonResult IsUserViewApply(string type, string Id, string EmpId)
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@type", type);
                list.Add("@CategoryId", Id);
                list.Add("@EmpId", EmpId);
                mes = comfun.executeNonQueryWMessage("sp_DocumentCategory_UserView", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        public JsonResult IsUserView(string type, string Id, string EmpId)
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@type", type);
                list.Add("@SubCategoryId", Id);
                list.Add("@EmpId", EmpId);
                mes = comfun.executeNonQueryWMessage("sp_DocumentSubCategory_UserView", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        public JsonResult removedocument()
        {
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@Id", Request.Form["documentid"].ToString());
                mes = comfun.executeNonQueryWMessage("stpDocumentRemove", "", list).ToString();

            }
            catch (Exception ex)
            {

                mes = ex.Message;
            }

            return Json(mes);
        }
        public ActionResult _DocumentSubCategoryList()
        {
            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            //  list.Add("@CategoryId", Request.Form["CategoryId"].ToString());
            DataTable dt = comfun.fillDataTable("stpDocumentRepositorylist", null, list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);

        }
        public FileResult Downloads(string P, string d)
        {
            SortedList list = new SortedList();
            string id = string.Empty;
            id = Request.QueryString["id"].ToString();
            list.Add("@Id", id);
            DataTable dt = comfun.fillDataTable("stpViksatDepartmentPolicy", "", list);
            if (dt.Rows[0]["fvImage"].ToString() != null)
            {
                P = dt.Rows[0]["fvImage"].ToString();
            }
            else
            {
                P = "";
            }

            /// return File(Path.Combine(Server.MapPath("~/Policies/"), P), P);//(Path.Combine(Server.MapPath("~/MedicalCertificate/"), P), System.Net.Mime.MediaTypeNames.Application.Octet, d);

            string contentType = "*";
            Response.AddHeader("Content-Disposition", "attachment;filename=\"" + P + "\"");
            Response.TransmitFile(Path.Combine(Server.MapPath("~/DMS/"), P));
            Response.End();
            //System.IO.FileInfo[] fileNames = dir.GetFiles("*.*");
            //List<string> items = new List<string>();
            //System.IO.FileInfo[] image = dir.GetFiles("*dt.Rows[0][fvImageName].ToString()*");
            //foreach (var file in fileNames)
            //{
            //    if (fileNames == image)
            //    {
            //        items.Add(file.Name);
            //    }
            //}

            return File(P, contentType, P);
        }
        /// <summary>
        /// District Master
        /// </summary>
        /// <returns></returns>
        public ActionResult District()
        {
            comfun.saveformname("District", "/admin/District", "Master/Address/District", "District  Main Form", "Y", "District", "District", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("City");

            DataTable dt = comfun.fillDataTable("DistrictState_Select", "", null);
            return View(dt);
        }
        //To Display City List Via State Id//
        public ActionResult _district_select()
        {
            comfun.saveformname("_district_select", "/admin/_district_select", "Master/Address/_district_select", "District  List", "N", "District", "District", "List", 4);
            SortedList list = new SortedList();
            list.Add("@DistrictId", Request.Form["DistrictId"].ToString());
            DataTable dt = comfun.fillDataTable("stpDistricts_list", "", list);


            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        //For District Save/Update Submit//
        public JsonResult _DistrictSaveUpdateSubmit()
        {
            comfun.saveformname("_DistrictSaveUpdateSubmit", "/admin/_DistrictSaveUpdateSubmit", "Master/Address/District", "District  Add", "N", "District", "District", "Add", 2);
            comfun.saveformname("_DistrictSaveUpdateSubmit", "/admin/_DistrictSaveUpdateSubmit", "Master/Address/District", "District  Edit", "N", "District", "District", "Edit", 3);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@DistrictId", Request.Form["DistrictId"].ToString());
                list.Add("@DistrictName", Request.Form["DistrictName"].ToString());
                list.Add("@DistrictCode", Request.Form["DistrictCode"].ToString());
                list.Add("@StateId", Request.Form["StateId"].ToString());
                //list.Add("@IsActive", "Y");
                //list.Add("@CreatedBy", Session["EmpId"].ToString());
                //list.Add("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                mes = comfun.executeNonQueryWMessage("stpDistricts_Accept", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
                //  mes = comfun.errorMessage("_CitiesAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        public JsonResult _DistrictStatusUpdate()
        {
            comfun.saveformname("_DistrictStatusUpdate", "/admin/_DistrictStatusUpdate", "Master/Address/District", "District  Status", "N", "District", "District", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpDistrict_Status", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }
        public ActionResult _DistrictEdit()
        {
            SortedList list = new SortedList();
            list.Add("@DistrictId", Request.Form["DistrictId"].ToString());
            DataTable dt = comfun.fillDataTable("District_SelectWithId", "", list);
            return PartialView("_DistrictEdit", dt);
        }
        /// <summary>
        /// Village Master
        /// </summary>
        /// <returns></returns>
        public ActionResult Village()
        {
            comfun.saveformname("Village", "/admin/Village", "Master/Address/Village", "Village  Main Form", "Y", "Village", "Village", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("City");

            DataTable dt = comfun.fillDataTable("VillageDistrict_Select", "", null);
            return View(dt);
        }
        //To Display City List Via State Id//
        public ActionResult _Village_select()
        {
            comfun.saveformname("_Village_select", "/admin/_Village_select", "Master/Address/_Village_select", "Village  List", "N", "Village", "Village", "List", 4);
            SortedList list = new SortedList();
            list.Add("@VillageId", Request.Form["VillageId"].ToString());
            DataTable dt = comfun.fillDataTable("stpVillage_list", "", list);


            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        //For District Save/Update Submit//
        public JsonResult _VillageSaveUpdateSubmit()
        {
            comfun.saveformname("_VillageSaveUpdateSubmit", "/admin/_VillageSaveUpdateSubmit", "Master/Address/Village", "Village  Add", "N", "Village", "Village", "Add", 2);
            comfun.saveformname("_VillageSaveUpdateSubmit", "/admin/_VillageSaveUpdateSubmit", "Master/Address/Village", "Village  Edit", "N", "Village", "Village", "Edit", 3);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@VillageId", Request.Form["VillageId"].ToString());
                list.Add("@VillageName", Request.Form["VillageName"].ToString());
                list.Add("@VillageCode", Request.Form["VillageCode"].ToString());
                list.Add("@DistrictId", Request.Form["DistrictId"].ToString());
                //list.Add("@IsActive", "Y");
                //list.Add("@CreatedBy", Session["EmpId"].ToString());
                //list.Add("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                mes = comfun.executeNonQueryWMessage("stpVillage_Accept", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
                //  mes = comfun.errorMessage("_CitiesAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        public JsonResult _VillageStatusUpdate()
        {
            comfun.saveformname("_VillageStatusUpdate", "/admin/_VillageStatusUpdate", "Master/Address/Village", "Village  Status", "N", "Village", "Village", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpVillage_Status", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }
        public ActionResult _District_selectddl()
        {
            SortedList list = new SortedList();
            //  list.Add("@CountryId", Request.Form["CountrStateIdyId"].ToString());
            DataTable dt = comfun.fillDataTable("District_SelectByStateId", "", list);
            return PartialView("_district_ddl", dt);
        }
        public ActionResult _VillageEdit()
        {
            SortedList list = new SortedList();
            list.Add("@VillageId", Request.Form["VillageId"].ToString());
            DataTable dt = comfun.fillDataTable("Village_SelectWithId", "", list);
            return PartialView("_VillageEdit", dt);
        }

        /// <summary>
        /// Supplier Charge Master
        /// </summary>
        /// <returns></returns>

        public ActionResult SupplierCharge()
        {
            comfun.saveformname("SupplierCharge", "/Admin/SupplierCharge", "Master/General/SupplierCharge", "Supplier Charge  Main Form", "Y", "SupplierCharge", "SupplierCharge", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("SupplierCharge");

            //DataTable dt = comfun.fillDataTable("CountryState_Select", "", null);
            return View();
        }
        //To Display City List Via St_Seldctate Id//
        public ActionResult _SupplierChargeList()
        {
            comfun.saveformname("_SupplierChargeList", "/Admin/_SupplierChargeList", "Master/General/SupplierCharge", "SupplierCharge  List", "N", "SupplierCharge", "SupplierCharge", "List", 4);

            DataTable dt = comfun.fillDataTable("stpSupplierChargeList", "", null);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        //For City Save/Update Submit//
        public JsonResult _SupplierChargeSaveUpdateSubmit()
        {
            comfun.saveformname("_SupplierChargeSaveUpdateSubmit", "/Admin/_SupplierChargeSaveUpdateSubmit", "Master/General/SupplierCharge", "SupplierCharge  Add", "N", "SupplierCharge", "SupplierCharge", "Add", 2);
            comfun.saveformname("_SupplierChargeSaveUpdateSubmit", "/Admin/_SupplierChargeSaveUpdateSubmit", "Master/General/SupplierCharge", "SupplierCharge  Edit", "N", "SupplierCharge", "SupplierCharge", "Edit", 3);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@SupplierChargeId", Request.Form["SupplierChargeId"].ToString());
                list.Add("@Charges", Request.Form["Charges"].ToString());
                list.Add("@AccountId", Request.Form["AccountId"].ToString());
                list.Add("@IsActive", "Y");
                //list.Add("@CreatedBy", Session["EmpId"].ToString());
                list.Add("@FromDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                list.Add("@UptoDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                mes = comfun.executeNonQueryWMessage("sp_SupplierChargeSaveUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
                //  mes = comfun.errorMessage("_CitiesAddSubmit", ex.Message);
            }
            return Json(mes);
        }
        public JsonResult _SupplierChargeStatusUpdate()
        {
            comfun.saveformname("_SupplierChargeStatusUpdate", "/Admin/_SupplierChargeStatusUpdate", "Master/General/SupplierCharge", "SupplierCharge  Status", "N", "SupplierCharge", "SupplierCharge", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpSupplierChargeStatusUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }

        //
        //Main DDL
        //public JsonResult _Locationddl()
        //{
        //    SortedList list = new SortedList();
        //    // list.Add("@CircleId", Request.Form["CircleId"].ToString());
        //    DataTable dt = comfun.fillDataTable("stpLocationddl", "", null);

        //    //return PartialView("_city_select", dt);
        //    var json = JsonConvert.SerializeObject(dt);
        //    return Json(json);

        //}

        //
        //Main DDL
        //public JsonResult _SupplierAccountddl()
        //{
        //    SortedList list = new SortedList();
        //   list.Add("@AccountId", Request.Form["AccountId"].ToString());
        //    DataTable dt = comfun.fillDataTable("stpSupplierAccountddl", "", null);

        //    //return PartialView("_city_select", dt);
        //    var json = JsonConvert.SerializeObject(dt);
        //    return Json(json);

        //}
        //public ActionResult __SupplierAccount_ddl()
        //{
        //    DataTable dt = comfun.fillDataTable("stpSupplierAccountddl", "", null);
        //    return PartialView(dt);
        //}

        public ActionResult _SupplierAccountddl()
        {
            DataTable dt = comfun.fillDataTable("stpSupplierAccountddl", "", null);
            return PartialView(dt);
        }


        public JsonResult _SupplierChargeEdit()
        {
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            DataTable dt = comfun.fillDataTable("stpSupplierChargeEdit", "", list);

            //return PartialView("_city_select", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public ActionResult LoginPassword()
        {
            return View();
        }
        public ActionResult _LoginPasswordlist()
        {
            SortedList list = new SortedList();
            list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
            list.Add("@EmpCategoryID", Request.Form["CategoryId"].ToString());
            //list.Add("@MainCategoryId", Request.Form["MainCategoryId"].ToString());
            list.Add("@DesignationID", Request.Form["DesignationId"].ToString());
            DataTable dt = comfun.fillDataTable("stpSignupManual", "", list);

            //return PartialView("_city_select", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public JsonResult _EmployeelistSubmit(string data)
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                DataTable table = (DataTable)JsonConvert.DeserializeObject(data, (typeof(DataTable)));
                //dynamic jsonData1 = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(jsonData);
                DataTable dt = new DataTable();
                string subject = "Welcome to WFMS";

                HtmlString htmlString = new HtmlString(mes);



                foreach (DataRow row in table.Rows)
                {
                    Random rnd = new Random();
                    string password = rnd.Next(1111, 9999).ToString();
                    StreamReader reader = new StreamReader(Server.MapPath("~/EmailTemplates/welcome.htm"));
                    mes = string.Empty;
                    list.Clear();
                    list.Add("@EmpID", row["EmpId"]);
                    list.Add("@Email", row["Email"]);
                    list.Add("@Password", password);
                    list.Add("@LoginID", Session["loginID"]);
                    mes = reader.ReadToEnd();
                    mes = mes.Replace("#EmployeeName#", row["EmpName"].ToString());
                    mes = mes.Replace("#Password#", password);
                    mes = mes.Replace("#Email#", row["Email"].ToString());
                    mes = mes.Replace("#Portal#", "wfms.htistelecom.in");
                    mes = mes.Replace("#Link#", "https://wfms.htistelecom.in");
                    htmlString = new HtmlString(mes);
                    reader.Close();
                    reader.Dispose();
                    payfun.SendEmail("Wfms LogIn", row["Email"].ToString(), "Login Detail", mes);
                    mes = string.Empty;
                    mes = comfun.executeNonQueryWMessage("stpSignupManual_Accept", "", list).ToString();

                }


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
        public ActionResult ContractorMaster()
        {
            return View();
        }
        public JsonResult _ContractorMasterList()
        {
            SortedList list = new SortedList();
            list.Add("@CustomerId", Session["Customer"]);

            DataTable dt = comfun.fillDataTable("stphtisContactorList1", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);

        }

        public JsonResult _ContractorEditSubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@ContractorId", Request.Form["Id"].ToString());
                list.Add("@ContractorName", Request.Form["ContractorName"].ToString());
                list.Add("@ContractorMobileNo", Request.Form["ContractorMobileNo"].ToString());
                list.Add("@ContractorAddress", Request.Form["ContractorAddress"].ToString());
                //list.Add("@Gender", Request.Form["Gender"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                list.Add("@Surcharge", Request.Form["Surcharge"].ToString());
                list.Add("@CreatedDate", comfun.dateISTstr());
                list.Add("@AadhaarNo", Request.Form["AadhaarNo"].ToString());
                list.Add("@PanNumber", Request.Form["PanNumber"].ToString());
                list.Add("@PFNO", Request.Form["PFNO"].ToString());
                list.Add("@ESINo", Request.Form["ESINo"].ToString());
                list.Add("@City", Request.Form["City"].ToString());
                list.Add("@State", Request.Form["State"].ToString());
                list.Add("@GSTNO", Request.Form["GSTNO"].ToString());
                list.Add("@ContractorCode", Request.Form["ContractorCode"].ToString());
                list.Add("@ContractorCompanyName", Request.Form["ContractorCompanyName"].ToString());
                list.Add("@BranchID", Request.Form["BranchID"].ToString());
                list.Add("@CompanyID", Request.Form["CompantID"].ToString());
                list.Add("@ContractorERPBPCode", Request.Form["ContractorERPBPCode"].ToString());
                mes = comfun.executeNonQueryWMessage("stphtisContractorSave", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _contractorAdd", ex.Message);
            }
            return Json(mes);



        }
        public ActionResult _ContractorEdit()
        {

            //string Contractorid = "0";
            if (Request.Form["Id"] == null)
            {
                return RedirectToAction("ContractorMaster");
            }
            SortedList list = new SortedList();
            list.Add("@ContractorID", Request.Form["Id"].ToString());
            DataTable dt = comfun.fillDataTable("stpviksatContactorWithID", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult _ContractorStatusUpdate()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("StpContractorStatusUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }

    }
}




