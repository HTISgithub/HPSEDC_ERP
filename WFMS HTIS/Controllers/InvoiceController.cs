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
using MasterIndia;
using MasterIndia.Model;
using System.Web.Script.Serialization;
using Payroll.portal.SessionLogout;

namespace Payroll.portal.Controllers
{
    [SessionExpire]
    public class InvoiceController : Controller
    {

        commonFunctions comfun = new commonFunctions();
        payrollFunctions payfun = new payrollFunctions();
        MasterIndiaFunction masterIndia = new MasterIndiaFunction();

        // GET: Deployment
        #region new

        public ActionResult NewInvoice()
        {
            return View();
        }
        //public ActionResult NewInvoice1()
        //{
        //    return View();
        //}


        public ActionResult _NewInvoiceSubmit()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
            list.Add("@BillDate", Request.Form["BillDate"].ToString());
            list.Add("@WorkOrderNo", Request.Form["WorkOrderNo"].ToString());
            list.Add("@NoOfResource", Request.Form["NoOfResource"].ToString());
            list.Add("@Billno", Request.Form["Billno"].ToString());
            list.Add("@AgencyId", Request.Form["ddlAgency"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@DepartmentAddress", Request.Form["DepartmentAddress"].ToString());
            list.Add("@Description", Request.Form["Description"].ToString());
            list.Add("@Narration", Request.Form["Narration"].ToString());
            list.Add("@AgencyBillAmt", Request.Form["AgencyBill"].ToString());
            list.Add("@AdminAmt", Request.Form["Admin"].ToString());
            list.Add("@LibaryAmt", Request.Form["Libary"].ToString());
            list.Add("@cgstAmt", Request.Form["cgst"].ToString());
            list.Add("@SGSTAtm", Request.Form["SGST"].ToString());
            list.Add("@TotalAmt", Request.Form["Total"].ToString());
            list.Add("@IGSTAmt", Request.Form["IGSTAmt"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@BillforMonth", Request.Form["BillForMonth"].ToString());
            list.Add("@BillingId", Request.Form["BillingId1"].ToString());
            list.Add("@AttendanceId", Request.Form["AttendaceId"].ToString());

            string mes = "";
            try
            {
                mes = comfun.executeNonQueryWMessage("TallyAgencyBill_AcceptUpdate", "", list).ToString();

            }
            catch (Exception ex) { mes = ex.Message; }

            return Json(mes);
        }
        public JsonResult _NewInvoiceList()

        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
            list.Add("@EmpId", Session["Empid"].ToString());
            list.Add("@UserRole", Session["RoleId"].ToString());
            list.Add("@MonthId", Request.Form["MonthId"].ToString());
            dt = comfun.fillDataTable("TallyAgencyBill1_List", "", list);
            //dt = comfun.fillDataTable("TallyAgencyBill_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        #region Cancel Sale Bill
        public ActionResult CancelSaleBill()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
            list.Add("@IsCancel", Request.Form["chkCancelSaleBill"].ToString());
            list.Add("@CancelRemarks", Request.Form["txtCancelSaleRemarks"].ToString());
            list.Add("@CancelBy", Session["EmpId"].ToString());

            string mes = "";
            try
            {
                mes = comfun.executeNonQueryWMessage("TallySaleBillCancel_AcceptUpdate", "", list).ToString();

            }
            catch (Exception ex) { mes = ex.Message; }

            return Json(mes);
        }

        #endregion Cancel Sale Bill
        #region Cancel Purchase Bill
        public ActionResult CancelPurchaseBill()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AttendanceId", Request.Form["AttendaceId"].ToString());
            list.Add("@AgencyBIllId", Request.Form["AgencyBillId"].ToString());
            //list.Add("@IsCancelPurchaseBill", Request.Form["IsCancelPurchaseBill"].ToString());
            //list.Add("@txtCancelRemarks", Request.Form["txtCancelRemarks"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString()); 

            string mes = "";
            try
            {
                mes = comfun.executeNonQueryWMessage("TallyDeletePurchaseBIll_AcceptUpdate", "", list).ToString();

            }
            catch (Exception ex) { mes = ex.Message; }

            return Json(mes);
        }

        #endregion Cancel Purcahse Bill
        #region Delete Attandance
        public ActionResult DeleteAttandance()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AttendaceId", Request.Form["AttendanceId"].ToString());
            list.Add("@IsCancel", Request.Form["chkDeleteAttendance"].ToString());
            //list.Add("@CancelRemarks", Request.Form["txtDeleteRemarks"].ToString());
            list.Add("@CancelBy", Session["Empid"].ToString()); 
            string mes = "";
            try
            {
                mes = comfun.executeNonQueryWMessage("TallyAttendanceCancel_AcceptUpdate", "", list).ToString();

            }
            catch (Exception ex) { mes = ex.Message; }

            return Json(mes);
        }

        #endregion Delete Attandance

        public JsonResult _NewInvoiceListget()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
            dt = comfun.fillDataTable("TallyDeptBill_ListGet1", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        #region Enter Purchase Bill(Existing)
        public JsonResult _InvoicePurchaseListget()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@WorkorderNo", Request.Form["WorkorderNo"].ToString());
            list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());

            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());

            dt = comfun.fillDataTable("TallyAgencyBillDetails_get", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _TallyAgencyDepartmentAddress_list()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            dt = comfun.fillDataTable("TallyAgencyDepartmentAddress_list", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _TallyAgencyWorkOrder_list()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            dt = comfun.fillDataTable("TallyAgencyWorkOrder_list", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }


        public JsonResult TallyAgencyDepartment_list()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            dt = comfun.fillDataTable("TallyAgencyDepartment_list", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        #endregion
        public JsonResult _NewInvoiceListFill()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
            dt = comfun.fillDataTable("TallyDeptBill_ListGet_Edit", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _DepartmentPaymentFill()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
            dt = comfun.fillDataTable("TallyAgencyBill_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _AvoDepartmentddlNew1()
        {
            SortedList list = new SortedList();
            list.Add("@DeptId", Session["DeptId"]?.ToString() ?? "0");
            list.Add("@createdBy", Session["EmpId"].ToString());
            DataTable dt = comfun.fillDataTable("stpHPSEDCSSODepartmentDDL", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult _Agencyddl1()
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

        public ActionResult _NewSaleInvoiceSubmit()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@DeptBillId", Request.Form["DeptBillId"].ToString());
            list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
            list.Add("@BillDate", Request.Form["BillDate1"].ToString());
            list.Add("@Billno", Request.Form["Billno1"].ToString());
            list.Add("@SaleBillNo", Request.Form["SaleBillNo"].ToString());
            list.Add("@WorkOderNo", Request.Form["WorkNo"].ToString());
            list.Add("@AgencyId", Request.Form["ddlAgency1"].ToString());
            list.Add("@DeptId", Request.Form["DeptId1"].ToString());
            list.Add("@DepartmentAddress", Request.Form["DepartmentAddress1"].ToString());
            list.Add("@Description", Request.Form["Description1"].ToString());
            list.Add("@Narration", Request.Form["Narration1"].ToString());
            list.Add("@AgencyBillAmt", Request.Form["AgencyBill1"].ToString());
            list.Add("@AdminAmt", Request.Form["Admin1"].ToString());
            list.Add("@LibaryAmt", Request.Form["Libary1"].ToString());
            list.Add("@cgstAmt", Request.Form["cgst1"].ToString());
            list.Add("@SGSTAtm", Request.Form["SGST1"].ToString());
            list.Add("@HSN", Request.Form["HSN"].ToString());
            list.Add("@gstNo", Request.Form["gst"].ToString());
            list.Add("@Pin", Request.Form["Pin"].ToString());
            list.Add("@TotalAmt", Request.Form["Total1"].ToString());
            list.Add("@PaymentAmt", "0");
            list.Add("@BalanceAmt", "0");
            list.Add("@IsActive", "Y");
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@BillforMonth", Request.Form["BillForMonth"].ToString());
            list.Add("@BillingId", Request.Form["BillingId1"].ToString());

            string mes = "";
            try
            {
                mes = comfun.executeNonQueryWMessage("TallyDepartmentBill_AcceptUpdate1", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }

            return Json(mes);
        }
        public ActionResult HpsedcInvoice1()
        {
            return View();
        }
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

       
        public JsonResult _NewInvoiceDeptList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
            dt = comfun.fillDataTable("TallyDepartmentBill_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
       
        public ActionResult BankList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            dt = comfun.fillDataTable("tblTallyBank_list", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
       
        public ActionResult _NewDeptPaySubmit()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();

            string msg = "";
            try
            {
                list.Add("@DepartmentBillId", "0");
                list.Add("ReceiptId", "0");
                list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
                list.Add("@BankNameId", Request.Form["BankNameId"].ToString());
                list.Add("@TransactionId", Request.Form["Transaction"].ToString());
                list.Add("@ModeOfPayment", Request.Form["ModePayment"].ToString());
                list.Add("@Narration", Request.Form["Remarks"].ToString());
                list.Add("@ReceivedDate", Request.Form["PaymentDate"].ToString());
                list.Add("@ReceivedAmt", Request.Form["PaymentAmount"].ToString());

                list.Add("@GSTTds2", Request.Form["GSTTDS"].ToString());
                list.Add("@Tds2", Request.Form["TDS"].ToString());

                list.Add("@CreatedBy", Session["EmpId"].ToString());
                msg = comfun.executeNonQueryWMessage("TallyReceivedPayment_AcceptUpdate", "", list).ToString();

            }
            catch (Exception ex) { msg = ex.Message; }

            return Json(msg);
        }


        public JsonResult _NewDeptPaymentList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@ReceiptId", Request.Form["ReceiptId"].ToString());
            list.Add("@DepartmentBillId", "0");
            list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
            dt = comfun.fillDataTable("TallyReceivedPaymentTransaction_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public JsonResult _NewInvoiceAgencytList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
            dt = comfun.fillDataTable("TallyDepartmentBill_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _NewAgencytPaySubmit()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            string Mes = "";
            try
            {

                list.Add("@PaymentId", Request.Form["PaymentId"].ToString());
                list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
                list.Add("@TransactionId", Request.Form["Transaction1"].ToString());
                list.Add("@ModeOfPayment", Request.Form["ModePayment2"].ToString());
                list.Add("@Narration", Request.Form["Remarks1"].ToString());
                list.Add("@PaymentDate", Request.Form["PaymentDate1"].ToString());

                list.Add("@GstTds2", Request.Form["GSTTDS1"].ToString());
                list.Add("@Tds1", Request.Form["TDS1"].ToString());
                list.Add("@Tds2", Request.Form["TDS2"].ToString());

                list.Add("@PaymentAmt", Request.Form["PaymentAmount1"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                Mes = comfun.executeNonQueryWMessage("TallyAgencyPayment_AcceptUpdate", "", list).ToString();

            }
            catch (Exception ex) { Mes = ex.Message; }

            return Json(Mes);
        }
        public JsonResult _NewAgencyPaymentList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@PaymentId", Request.Form["PaymentId"].ToString());
            list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
            dt = comfun.fillDataTable("TallyAgencyPaymentTransaction_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _PaymentModeddl1()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId, Session["Customer"]);
            list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("HpsedcPaymentmodeDdl_List", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public ActionResult _PaymentModeddl2()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId, Session["Customer"]);
            list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("HpsedcPaymentmodeDdl_List", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }

        
        #endregion

        #region Report

        public ActionResult NewInvoiceReport()
        {
            return View();
        }

        public ActionResult InvoiceReport()
        {
            return View();
        }

        public JsonResult _invoiceReportList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@MonthId", Request.Form["MonthId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@PReceived", Request.Form["PReceived"].ToString());
            list.Add("@PReleased", Request.Form["PReleased"].ToString());
            list.Add("@SaleBillNo", Request.Form["SaleBillNo"].ToString());
            list.Add("@Balance", Request.Form["Balance"].ToString());
            dt = comfun.fillDataTable("TallyDeptBill_Report1", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }


        public JsonResult _PaymentReceivedReportList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@MonthId", Request.Form["MonthId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@PReceived", Request.Form["PReceived"].ToString());
            list.Add("@SaleBillNo", Request.Form["SaleBillNo"].ToString());

            dt = comfun.fillDataTable("TallyPaymentReceived_Report", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        #endregion
        #region Master Department Address

        public ActionResult DepartmentAddress()
        {
            return View();
        }

        public ActionResult _District_selectddl1()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@StateId", Request.Form["StateId"].ToString());

            dt = comfun.fillDataTable("stpDistrictddl", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }


        public ActionResult _DepartmentAddressSubmit()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();

            string msg = "";
            try
            {
                
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@DeptId", Request.Form["departmentID"].ToString());
                list.Add("@NodalName", Request.Form["Nodel"].ToString());
                list.Add("@ContactNo", Request.Form["Mobileno"].ToString());
                list.Add("@EmailId", Request.Form["EmailId"].ToString());
                list.Add("@BillingAddress", Request.Form["Address"].ToString());
                msg = comfun.executeNonQueryWMessage("TallyAddressJune_AcceptUpdate", "", list).ToString();

            }
            catch (Exception ex) { msg = ex.Message; }

            return Json(msg);
        }

        public JsonResult _DepartmentAddressStatus()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("TallyAddressJune_AcceptUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }

        public JsonResult _DepartmentAddress_List()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("TallyAddressJune_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }


        public JsonResult _DepartmentAddress_Edit()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("stpDepartmentAddress_Edit", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        #endregion

        #region Agency Bill Status
        public ActionResult AgencyBillStatus()
        {
            return View();
        }

        public JsonResult _AgencyBillStatus_List()

        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@MonthYearId", Request.Form["MonthYearId"].ToString()); 
            dt = comfun.fillDataTable("TallyAgencyBillStatus_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        #endregion

        #region Payment Recived

        public ActionResult PaymentRecived()
        {
            return View();
        }
        #endregion

         #region Agency Payment
        public ActionResult AgencyPayment()
        {
            return View();
        }

        #endregion


        #region
        public ActionResult PaymentReceivedReport()
        {
            return View();
        }
        #endregion
        
        public ActionResult PaymentRelease()
        {
            return View();
        }

        public JsonResult _PaymentRelease_List()

        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@MonthId", Request.Form["MonthId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
           // list.Add("@PReceived", Request.Form["PReceived"].ToString());
            list.Add("@SaleBillNo", Request.Form["SaleBillNo"].ToString());
            dt = comfun.fillDataTable("TallyAgencyPaymentReleased_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult NewPaymentReportTally()
        {
            return View();
        }
        public ActionResult _NewPaymentReportTallyList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthId"].ToString());
            dt = comfun.fillDataTable("HardwarePaymentAmtReport_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult NewRecieptReportTally()
        {
            return View();
        }

        #region EInvoice
        public ActionResult EInvoice()
        {
            return View();
        }

        public JsonResult _EInvoiceList()
        {
            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("TallyDeptBillForEInvoice_List", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }


        public async Task<ActionResult> Generate_IRN(int id)
        {
            int invoiceId = id;

            //int.TryParse(Request.QueryString["id"], out invoiceId);
            //invoiceId = 5517;
            string Message = "";

            try
            {
                masterIndia.generate_App_Key(invoiceId);

                var client = new HttpClient();
                // get invoice authtoken details
                InvoiceInfo invInfo = new InvoiceInfo();
                invInfo = masterIndia.InvoiceInformation(invoiceId);
                //InvoiceDatajSON(invoiceId);
                var json1 = new JavaScriptSerializer().Serialize(masterIndia.InvoiceDatajSON(invoiceId));

                client.DefaultRequestHeaders.Add("Authorization", invInfo.AuthenticationToken);

                var content = new StringContent(json1, Encoding.UTF8, "application/json");
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                //var response = client.PostAsync("https://sandb-api.mastersindia.co/api/v1/einvoice/", content).GetAwaiter().GetResult();
                var response = client.PostAsync("https://prod-api.mastersindia.co/api/v1/einvoice/", content).GetAwaiter().GetResult();

                if (response.IsSuccessStatusCode)
                {
                    var contents = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    Message = contents;
                    var response1 = JsonConvert.DeserializeObject<Response>(contents);

                    var Status = response1.results.message.Status;
                    var Ack = response1.results.message.Ackno;
                    var AckDt = response1.results.message.AckDt;
                    var Irn = response1.results.message.Irn;
                    var SignedInvoice = response1.results.message.SignedInvoice;
                    var SignedQRCode = response1.results.message.SignedQRCode;

                    //save data in database
                    using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))
                    {
                        using (SqlCommand cmd = new SqlCommand("InsertInvoiceData", con))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add(new SqlParameter("@InvoiceId", invoiceId));
                            cmd.Parameters.Add(new SqlParameter("@AckNo", Ack));
                            cmd.Parameters.Add(new SqlParameter("@AckDt", AckDt));
                            cmd.Parameters.Add(new SqlParameter("@Irn", Irn));
                            cmd.Parameters.Add(new SqlParameter("@SignedInvoice", SignedInvoice));
                            cmd.Parameters.Add(new SqlParameter("@SignedQRCode", SignedQRCode));
                            cmd.Parameters.Add(new SqlParameter("@Status", Status));


                            cmd.Parameters.Add(new SqlParameter("@IsActive", true));
                            con.Open();
                            cmd.ExecuteNonQuery();
                            masterIndia.SignedINVQRCode(invoiceId);
                            Message = "Invoice data Submitted Successfully";
                        }
                    }

                }
                else
                {
                    Message = "Error in IRN API";

                }
            }
            catch (Exception ex)
            {
                //Message = "Error in Saving Invoice data";

            }
            return Json(Message.ToString(), JsonRequestBehavior.AllowGet);
        }
        #endregion

        #region Raushan
        public ActionResult TdsReport()
        {
            return View();
        }

        public ActionResult GstTaxReport()
        {
            return View();
        }

        public ActionResult GstReport()
        {
            return View();
        }
        public ActionResult _TdsReportList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthId"].ToString());
            dt = comfun.fillDataTable("TallyAgencyPayment_Report", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _GstReportList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthId"].ToString());
            dt = comfun.fillDataTable("TallyAgencyDeptBill_Report", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult TotalResourceReport()
        {
            return View();
        }

        public ActionResult SubMonthlyReport()
        {
            return View();
        }


        public ActionResult _ResourceReportList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            dt = comfun.fillDataTable("TotalResReport", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _SubMonthlyReportList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@MonthId", Request.Form["MonthId"].ToString());
            list.Add("@DepartmentId", Request.Form["DepId"].ToString());
            dt = comfun.fillDataTable("MonthlySubAmount_Report", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        #endregion

        #region Employee Upload Excel
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
            // string date = DateTime.UtcNow.AddMinutes(330).ToString("yyyyMMdd_HHmm");
            // excelName = excelName + "_" + date;
            DataTable dt = new DataTable();
            dt.Columns.Add("Sr.", typeof(string));
            dt.Columns.Add("Empname", typeof(string));
            dt.Columns.Add("Father's Name", typeof(string));
            dt.Columns.Add("IsFullTime", typeof(string));
            dt.Columns.Add("DesigationId", typeof(string));
            dt.Columns.Add("AADHAR", typeof(string));
            dt.Columns.Add("Basic", typeof(string));
            dt.Columns.Add("Others", typeof(string));
            dt.Columns.Add("IsPf", typeof(string));
            dt.Columns.Add("IsEsi", typeof(string));


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
            string mes = "";
            DateTime today = DateTime.UtcNow;
            //var todays = today.ToString("dd/MMM/yyyy");
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
            dtExcel.Columns.Add("Sr.", typeof(string));
            dtExcel.Columns.Add("Empname", typeof(string));
            dtExcel.Columns.Add("Father_Name", typeof(string));
            dtExcel.Columns.Add("IsFullTime", typeof(string));
            dtExcel.Columns.Add("DesigationId", typeof(string));
            dtExcel.Columns.Add("AADHAR", typeof(string));
            dtExcel.Columns.Add("Basic", typeof(string));
            dtExcel.Columns.Add("Others", typeof(string));
            dtExcel.Columns.Add("IsPf", typeof(string));
            dtExcel.Columns.Add("IsEsi", typeof(string));
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

                    if (dt.Rows[i]["Sr."] != System.DBNull.Value)
                    {
                        dr["Sr."] = dt.Rows[i]["Sr."].ToString();

                    }
                    else
                    {
                        dr["sr."] = "0";
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

                    if (dt.Rows[i]["IsFullTime"] != System.DBNull.Value)
                    {
                        dr["IsFullTime"] = dt.Rows[i]["IsFullTime"].ToString();


                    }
                    else
                    {

                        dr["IsFullTime"] = "0";
                    }

                    if (dt.Rows[i]["DesigationId"] != System.DBNull.Value)
                    {
                        dr["DesigationId"] = dt.Rows[i]["DesigationId"].ToString();
                    }
                    else
                    {
                        dr["DesigationId"] = "0";
                    }

                    if (dt.Rows[i]["AADHAR"] != System.DBNull.Value)
                    {
                        dr["AADHAR"] = dt.Rows[i]["AADHAR"].ToString();
                    }
                    else
                    {
                        dr["AADHAR"] = "0";
                    }

                    if (dt.Rows[i]["Basic"] != System.DBNull.Value)
                    {
                        dr["Basic"] = dt.Rows[i]["Basic"].ToString();
                    }
                    else
                    {
                        dr["Basic"] = "0";
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

                    dtExcel.Rows.Add(dr);

                }
                //i++ ;
            }
            return dtExcel;
        }


        public JsonResult _InsertDetailVerification()
        {

            string mes = string.Empty;
            DataTable dataTable = new DataTable();
            try
            {
                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(Request.Form["Table"].ToString());
                DataTable table = new DataTable();
                //table.Columns.Add("Sr.", typeof(string));
                table.Columns.Add("DeptId", typeof(string));
                table.Columns.Add("AgencyId", typeof(string));
                table.Columns.Add("WorkOrderNo", typeof(string));
                table.Columns.Add("TotalManpower", typeof(string));
                //table.Columns.Add("NoOfEmp", typeof(string));
                table.Columns.Add("EmpName", typeof(string));
                table.Columns.Add("Father_Name", typeof(string));
                table.Columns.Add("IsFullTime", typeof(string));
                table.Columns.Add("DesigationId", typeof(string));
                table.Columns.Add("AADHAR", typeof(string));
                table.Columns.Add("fnBasic", typeof(string));
                table.Columns.Add("Others", typeof(string));
                table.Columns.Add("IsPf", typeof(string));
                table.Columns.Add("IsEsi", typeof(string));

                // DeptId ,AgencyId ,WorkOrderNo ,TotalManpower

                foreach (var Item1 in jsonData)
                {
                    DataRow dr1 = table.NewRow();

                    //dr1["Sr."] = Item1.Sr;
                    dr1["DeptId"] = Request.Form["DeptId"].ToString();
                    dr1["AgencyId"] = Request.Form["ddlAgencynew"].ToString();
                    dr1["WorkOrderNo"] = Request.Form["WorkOrderNo"].ToString();
                    dr1["TotalManpower"] = Request.Form["Resources"].ToString();
                    //dr1["NoOfEmp"] = Request.Form["NoOfEmp"].ToString();
                    dr1["EmpName"] = Item1.EmpName;
                    dr1["Father_Name"] = Item1.Father_Name;
                    dr1["IsFullTime"] = Item1.IsFullTime;
                    dr1["DesigationId"] = Item1.DesigationId;
                    dr1["AADHAR"] = Item1.AADHAR;
                    dr1["fnBasic"] = Item1.Basic;
                    dr1["Others"] = Item1.Others;
                    dr1["IsPf"] = Item1.IsPf;
                    dr1["IsEsi"] = Item1.IsEsi;

                    table.Rows.Add(dr1);
                }


                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "stpTallyEmployeesImportMasterVerified";
                // com.CommandText = "stpTallyEmployeesImportMaster";
                SqlParameter parameter = new SqlParameter();
                com.Parameters.AddWithValue("@EmpId", Session["EmpId"].ToString());
                com.Parameters.AddWithValue("@NoOfEmp", Request.Form["Resources"].ToString());
                //com.Parameters.AddWithValue("@EmpDetails", Session["SessionId"].ToString());
                parameter.ParameterName = "@EmpDetails";
                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table;
                com.Parameters.Add(parameter);
                //com.Parameters.Add("@Mes", SqlDbType.VarChar, 500);
                //com.Parameters["@Mes"].Direction = ParameterDirection.Output;
                // com.CommandType = System.Data.CommandType.StoredProcedure;
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

        public JsonResult _InsertDetail()
        {

            string mes = string.Empty;
            DataTable dataTable = new DataTable();
            try
            {
                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(Request.Form["Table"].ToString());
                DataTable table = new DataTable();
                //table.Columns.Add("Sr.", typeof(string));
                table.Columns.Add("DeptId", typeof(string));
                table.Columns.Add("AgencyId", typeof(string));
                table.Columns.Add("WorkOrderNo", typeof(string));
                table.Columns.Add("TotalManpower", typeof(string));
                //table.Columns.Add("NoOfEmp", typeof(string));
                table.Columns.Add("EmpName", typeof(string));
                table.Columns.Add("Father_Name", typeof(string));
                table.Columns.Add("IsFullTime", typeof(string));
                table.Columns.Add("DesigationId", typeof(string));
                table.Columns.Add("AADHAR", typeof(string));
                table.Columns.Add("fnBasic", typeof(string));
                table.Columns.Add("Others", typeof(string));
                table.Columns.Add("IsPf", typeof(string));
                table.Columns.Add("IsEsi", typeof(string));

                // DeptId ,AgencyId ,WorkOrderNo ,TotalManpower

                foreach (var Item1 in jsonData)
                {
                    DataRow dr1 = table.NewRow();

                    //dr1["Sr."] = Item1.Sr;
                    dr1["DeptId"] = Request.Form["DeptId"].ToString();
                    dr1["AgencyId"] = Request.Form["ddlAgencynew"].ToString();
                    dr1["WorkOrderNo"] = Request.Form["WorkOrderNo"].ToString();
                    dr1["TotalManpower"] = Request.Form["Resources"].ToString();
                 //   dr1["NoOfEmp"] = Request.Form["NoOfEmp"].ToString();
                    dr1["EmpName"] = Item1.EmpName;
                    dr1["Father_Name"] = Item1.Father_Name;
                    dr1["IsFullTime"] = Item1.IsFullTime;
                    dr1["DesigationId"] = Item1.DesigationId;
                    dr1["AADHAR"] = Item1.AADHAR;
                    dr1["fnBasic"] = Item1.Basic;
                    dr1["Others"] = Item1.Others;
                    dr1["IsPf"] = Item1.IsPf;
                    dr1["IsEsi"] = Item1.IsEsi;

                    table.Rows.Add(dr1);
                }


                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                //com.CommandText = "stpTallyEmployeesImportMasterVerified";
                 com.CommandText = "stpTallyEmployeesImportMaster";
                SqlParameter parameter = new SqlParameter();
                com.Parameters.AddWithValue("@EmpId", Session["EmpId"].ToString());
                //com.Parameters.AddWithValue("@NoOfEmp", Request.Form["Resources"].ToString());
                parameter.ParameterName = "@EmpDetails";
                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table;
                com.Parameters.Add(parameter);
                com.Parameters.Add("@Mes", SqlDbType.VarChar, 500);
                com.Parameters["@Mes"].Direction = ParameterDirection.Output;
                com.CommandTimeout = 0;
                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();
                com.ExecuteNonQuery();
                mes = (string)com.Parameters["@Mes"].Value;
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
             var json = JsonConvert.SerializeObject(mes);
            return Json(json);
        }




        public JsonResult _EmpAgencyList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@WorkOrderNo", Request.Form["WorkOrderNo"].ToString());
            list.Add("@EmpId", Session["EmpId"].ToString());
            dt = comfun.fillDataTable("stpTallyEmployees_list", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        
        #endregion
        #region MAster
        public ActionResult DEPTMaster()
        {
            return View();
        }
        public ActionResult _DEPTMasterSubmit()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@WorkOrderAgencyId", Request.Form["WorkOrderAgencyId"].ToString());
            list.Add("@AgencyId", Request.Form["ddlAgency"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@WorkOrderId", Request.Form["Work"].ToString());
            list.Add("@BillingAddress", Request.Form["BillingAddress"].ToString());
            list.Add("@BillingId", Request.Form["BillingId"].ToString());
            list.Add("@NoDeployedRes", Request.Form["NoofResources"].ToString());
            list.Add("@BillAddressEmail", Request.Form["Email"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());

            string mes = "";
            try
            {
                mes = comfun.executeNonQueryWMessage("TallyAgencyWorkOrder_AcceptUpdate", "", list).ToString();

            }
            catch (Exception ex) { mes = ex.Message; }

            return Json(mes);
        }
        public ActionResult _BillingAddressMerge()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId, Session["Customer"]);
            list.Add("@AgencyId", Session["EmpId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@BillingAddress", Request.Form["WorkOrderAgencyId"].ToString());
            dt = comfun.fillDataTable("TallyAgencyWorkOrder2_list", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public JsonResult _DEPTMasterList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@WorkOrderId", Request.Form["WorkOrderId"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@RoleId", Session["RoleId"].ToString());
            //dt = comfun.fillDataTable("TallyAgencyDeptWorkOrder_List", "", list);
            dt = comfun.fillDataTable("TallyAgencyDeptWorkOrder_List1", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }


        public ActionResult _Agencyddl2()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId, Session["Customer"]);
            list.Add("@AgencyId", Session["EmpId"].ToString());
            list.Add("@RoleId", Session["RoleId"].ToString());
            dt = comfun.fillDataTable("TallyAgencyDdl_List", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public ActionResult _AvoDepartmentddlNew2()
        {
            SortedList list = new SortedList();
            list.Add("@DeptId", "0");
            DataTable dt = comfun.fillDataTable("TallyAgencyDepartment1_list", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult _BillingAddressddl()
        {
            SortedList list = new SortedList();
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            //list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            DataTable dt = comfun.fillDataTable("TallyWorkOrderBillingAddressJune_dll", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }

        public ActionResult _BillingAddressddl1()
        {
            SortedList list = new SortedList();
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            //list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            DataTable dt = comfun.fillDataTable("TallyWorkOrderBillingAddress_dll", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public JsonResult _DEPTMasterListEdit()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@WorkOrderAgencyId", Request.Form["WorkOrderAgencyId"].ToString());
            dt = comfun.fillDataTable("TallyAgencyDeptWorkOrder_Edit", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        #endregion

        #region Attendance
        public ActionResult DeptAttendance()
        {
            return View();
        }
        public ActionResult _NewDeptAttendanceSubmit()
        {
            string mes = "";
            try
            {
                string AttendanceCertificate = "";
                string AnnexureFile = "";
                string AgencyBillFile = "";

                string attendanceFolderPath = Server.MapPath("/TallyAttendanceFiles/");
                string annexureFolderPath = Server.MapPath("/TallyAnnexureFiles/");
                string agencyBillFolderPath = Server.MapPath("/TallyAgencyBillFiles/");

                if (!Directory.Exists(attendanceFolderPath))
                {
                    Directory.CreateDirectory(attendanceFolderPath);
                }
                if (!Directory.Exists(annexureFolderPath))
                {
                    Directory.CreateDirectory(annexureFolderPath);
                }
                if (!Directory.Exists(agencyBillFolderPath))
                {
                    Directory.CreateDirectory(agencyBillFolderPath);
                }

                HttpFileCollectionBase files = Request.Files;
                if (System.Web.HttpContext.Current.Request.Files.AllKeys.Any())
                {
                    for (int i = 0; i < files.Count; i++)
                    {
                        HttpPostedFileBase file = files[i];
                        string fileName = Path.GetFileName(file.FileName);
                        string extension = Path.GetExtension(fileName);
                        string fileWithoutExt = Path.GetFileNameWithoutExtension(fileName);

                        if (i == 0)
                        {
                            string attendanceFilePath = Path.Combine(attendanceFolderPath, fileWithoutExt + extension);
                            AttendanceCertificate = fileWithoutExt + extension;
                            file.SaveAs(attendanceFilePath);
                        }
                        else if (i == 1)
                        {
                            string annexureFilePath = Path.Combine(annexureFolderPath, fileWithoutExt + extension);
                            AnnexureFile = fileWithoutExt + extension;
                            file.SaveAs(annexureFilePath);
                        }
                        else if (i == 2)
                        {
                            string agencyBillFilePath = Path.Combine(agencyBillFolderPath, fileWithoutExt + extension);
                            AgencyBillFile = fileWithoutExt + extension;
                            file.SaveAs(agencyBillFilePath);
                        }
                    }
                }

                DataTable dt = new DataTable();
                SortedList list = new SortedList();
                list.Add("@Id", "0");
                list.Add("@UpladNoOfResource", Request.Form["UpladNoOfResource"].ToString());
                list.Add("@MonthYear", Request.Form["MonthYear"].ToString());
                list.Add("@WorkOrderId", Request.Form["WorkOrderNo"].ToString());
                list.Add("@AttendanceCertificate", AttendanceCertificate);
                list.Add("@AnnexureFile", AnnexureFile);
                list.Add("@AgencyBillFile", AgencyBillFile); // Naya field
                list.Add("@CreatedBy", Session["EmpId"].ToString());

                mes = comfun.executeNonQueryWMessage("tblTallyAttendance_AcceptUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }

            return Json(mes);
        }
        public ActionResult _AgencyUploadAttendanceSubmit()
        {
            string mes = "";
            try
            {
                //string AttendanceCertificate = "";
                string AnnexureFile = "";
                string AgencyBillFile = "";

                //string attendanceFolderPath = Server.MapPath("/TallyAttendanceFiles/");
                string annexureFolderPath = Server.MapPath("/TallyAnnexureFiles/");
                string agencyBillFolderPath = Server.MapPath("/TallyAgencyBillFiles/");

                //if (!Directory.Exists(attendanceFolderPath))
                //{
                //    Directory.CreateDirectory(attendanceFolderPath);
                //}
                if (!Directory.Exists(annexureFolderPath))
                {
                    Directory.CreateDirectory(annexureFolderPath);
                }
                if (!Directory.Exists(agencyBillFolderPath))
                {
                    Directory.CreateDirectory(agencyBillFolderPath);
                }

                HttpFileCollectionBase files = Request.Files;
                if (System.Web.HttpContext.Current.Request.Files.AllKeys.Any())
                {
                    for (int i = 0; i < files.Count; i++)
                    {
                        HttpPostedFileBase file = files[i];
                        string fileName = Path.GetFileName(file.FileName);
                        string extension = Path.GetExtension(fileName);
                        string fileWithoutExt = Path.GetFileNameWithoutExtension(fileName);

                        //if (i == 0)
                        //{
                        //    string attendanceFilePath = Path.Combine(attendanceFolderPath, fileWithoutExt + extension);
                        //    AttendanceCertificate = fileWithoutExt + extension;
                        //    file.SaveAs(attendanceFilePath);
                        //}
                         if (i == 0)
                        {
                            string annexureFilePath = Path.Combine(annexureFolderPath, fileWithoutExt + extension);
                            AnnexureFile = fileWithoutExt + extension;
                            file.SaveAs(annexureFilePath);
                        }
                        else if (i == 1)
                        {
                            string agencyBillFilePath = Path.Combine(agencyBillFolderPath, fileWithoutExt + extension);
                            AgencyBillFile = fileWithoutExt + extension;
                            file.SaveAs(agencyBillFilePath);
                        }
                    }
                }

                DataTable dt = new DataTable();
                SortedList list = new SortedList();
                
                list.Add("@Id", Request.Form["Id"].ToString());
                //list.Add("@MonthYear", Request.Form["MonthYear"].ToString());
                //list.Add("@WorkOrderId", Request.Form["WorkOrderNo"].ToString());
                //list.Add("@AttendanceCertificate", AttendanceCertificate);
                list.Add("@AnnexureFile", AnnexureFile);
                list.Add("@AgencyBillFile", AgencyBillFile); // Naya field
                list.Add("@CreatedBy", Session["EmpId"].ToString());

                mes = comfun.executeNonQueryWMessage("tblTallyAnnexture_AcceptUpdate  ", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }

            return Json(mes);
        }

        public ActionResult _AgencyNewAttendanceDdl()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@AgencyId", Session["EmpId"]?.ToString() ?? "0");
            list.Add("@UserId", Session["EmpId"].ToString());
            list.Add("@RoleId", Session["RoleId"].ToString());
            dt = comfun.fillDataTable("TallyAttendanceAgencyDdl_List", "", list);
            
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _AgencyAttendanceList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthId"].ToString());
            list.Add("@CreateBy", Session["EmpId"].ToString());
            list.Add("@RoleId", Session["RoleId"].ToString());
            dt = comfun.fillDataTable("tblTallyAttendance_list", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _AgencyAttenInvget()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AttendaceId", Request.Form["AttendaceId"].ToString());
            dt = comfun.fillDataTable("tblTallyAttendanceDetails_Get", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }


        public JsonResult _AgencyAttenUploadGet()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("tblTallyAttendance_Get", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _WorkOrderDdl()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId, Session["Customer"]);
            list.Add("@AgencyId", Session["EmpId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@WorkOrderAgencyId", Request.Form["WorkOrderAgencyId"].ToString());
            list.Add("@UserId", Session["EmpId"].ToString());
            list.Add("@RoleId", Session["RoleId"].ToString());
            dt = comfun.fillDataTable("TallyAgencyWorkOrder1_get1", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        #endregion

        public ActionResult NewPaymentRecivedReportTally()
        {
            return View();
        }
        public ActionResult _NewPaymenRecivedtReportTallyList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthId"].ToString());
            dt = comfun.fillDataTable("HardwareRceivedAmtReport_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult RecieptRelesaeReportTally()
        {
            return View();
        }
        public ActionResult _RecieptRelesaeReportTallyList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthId"].ToString());
            dt = comfun.fillDataTable("HardwarePaymentAmtReport_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        #region Department Partial Payment
        public JsonResult _deptPartialPayListGet()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
            dt = comfun.fillDataTable("TallyDepartmentBill_Partial", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public JsonResult _deptPartialPayList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
            dt = comfun.fillDataTable("TallyDeptParymentReceived_get", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public JsonResult _DepartmentPaymentVerification()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@ReceiptId", Request.Form["ReceiptId"].ToString());
            dt = comfun.fillDataTable("TallysDeptPaymentTransaction_Get", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult DepartmentPaymentVerificationSubmit()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();

            string msg = "";
            try
            {
                list.Add("@ReceiptId", Request.Form["ReceiptId"].ToString());
                list.Add("@TransactionId", Request.Form["TransactionId"].ToString());
                list.Add("@IsTransactionVerified", Request.Form["IsPurchaseBillVerified"].ToString());
                list.Add("@TransactionVerifiedRemarks", Request.Form["VerificationRemarks"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                msg = comfun.executeNonQueryWMessage("TallysDeptPaymentTransaction_Verified", "", list).ToString();

            }
            catch (Exception ex) { msg = ex.Message; }

            return Json(msg);
        }

        public ActionResult _DeptPartialPaySubmit()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();

            string msg = "";
            try
            {
                list.Add("@DepartmentBillId", "0");
                list.Add("ReceiptId", "0");
                list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
                list.Add("@BankNameId", Request.Form["BankNameId"].ToString());
                list.Add("@TransactionId", Request.Form["Transaction"].ToString());
                list.Add("@ModeOfPayment", Request.Form["ModePayment"].ToString());
                list.Add("@Narration", Request.Form["Remarks"].ToString());
                list.Add("@ReceivedDate", Request.Form["PaymentDate"].ToString());
                list.Add("@ReceivedAmt", Request.Form["PaymentAmount"].ToString());

                list.Add("@GSTTds2", Request.Form["GSTTDS"].ToString());
                list.Add("@Tds2", Request.Form["TDS"].ToString());

                list.Add("@CreatedBy", Session["EmpId"].ToString());
                msg = comfun.executeNonQueryWMessage("TallyReceivedPaymentPartial_AcceptUpdate", "", list).ToString();

            }
            catch (Exception ex) { msg = ex.Message; }

            return Json(msg);
        }
        public ActionResult BankListpp()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            dt = comfun.fillDataTable("tblTallyBank_list", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _PaymentModeddlpp()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId, Session["Customer"]);
            list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("HpsedcPaymentmodeDdl_List", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        #endregion

        #region AgencyPartialPayment

        public ActionResult _AgencyPartialPaymentBankDDL()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            dt = comfun.fillDataTable("tblTallyBank_list", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _AgencyPartialModePaymentDDL()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("HpsedcPaymentmodeDdl_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public JsonResult _AgencyPartialPayListGet()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
            dt = comfun.fillDataTable("TallyAgencyBillPaymentPartial_Get", "", list);
            //dt = comfun.fillDataTable("TallyAgencyBill_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public JsonResult _AgencyPartialPayList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
            dt = comfun.fillDataTable("TallyAgencyParymentTransaction_get", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _AgencyPartialPaySubmit()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();

            string msg = "";
            try
            {
               // list.Add("@DepartmentBillId", "0");
                //list.Add("@ReceiptId", "0");
                list.Add("@PaymentId", "0");
                list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
                list.Add("@TransactionId", Request.Form["TransactionId"].ToString());
               // list.Add("@BankNameId", Request.Form["BankNameId"].ToString());
                list.Add("@ModeOfPayment", Request.Form["ModeOfPayment"].ToString());
                list.Add("@Narration", Request.Form["Narration"].ToString());
                list.Add("@PaymentDate", Request.Form["PaymentDate"].ToString());
                list.Add("@PaymentAmt", Request.Form["PaymentAmt"].ToString());

                list.Add("@GstTds2", Request.Form["GstTds2"].ToString());
                list.Add("@Tds1", Request.Form["Tds1"].ToString());
                list.Add("@Tds2", Request.Form["Tds2"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                msg = comfun.executeNonQueryWMessage("TallyAgencyPaymentPatrial_AcceptUpdate", "", list).ToString();

            }
            catch (Exception ex) { msg = ex.Message; }

            return Json(msg);
        }

        #endregion

        #region Merge
        public ActionResult MergeBill()
        {
            return View();
        }
        public ActionResult _MergeAddressDdl()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId, Session["Customer"]);
            list.Add("@AgencyId", Session["EmpId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@WorkOrderAgencyId", Request.Form["WorkOrderAgencyId"].ToString());
            dt = comfun.fillDataTable("TallyAgencyWorkOrder1_list", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public ActionResult _MergeWorkOrderBiling()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@WorkOrderAgencyId", Request.Form["WorkOrderAgencyId"].ToString());
            list.Add("@MergeIntoWorkOrderId", Request.Form["MegeWorkOrderTwo"].ToString());
            list.Add("@UserRole", Session["EmpId"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());

            string mes = "";
            try
            {
                mes = comfun.executeNonQueryWMessage("TallyAgencyWorkOrder_Merged", "", list).ToString();

            }
            catch (Exception ex) { mes = ex.Message; }

            return Json(mes);
        }


        public ActionResult NEWTDSREPORTS()
        {
            return View();
        }

        public ActionResult _NEWTDSREPORTSList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthId"].ToString());
            dt = comfun.fillDataTable("tallyAgencyTds_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _NewTDSGETList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@ChallanId", Request.Form["ChallanId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthYear"].ToString());
            dt = comfun.fillDataTable("TallyTdsChallanDeposit_Get", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

         public ActionResult _DepositedTdsSubmit()
        {
            DataTable dt = new DataTable();

            SortedList list = new SortedList();

            list.Add("@ChallanId", Request.Form["ChallanId"].ToString());

            list.Add("@TotalNoOfPurchaseInv", Request.Form["TotalPurchaseInv"].ToString());

            list.Add("@InvAmount", Request.Form["TotaltdsAmount"].ToString());

            list.Add("@MonthId", Request.Form["monthyear"].ToString());

            list.Add("@ChallanNO", Request.Form["ChallanNo"].ToString());

            list.Add("@ChallanDate", Request.Form["DepositedDate"].ToString());

            list.Add("@BSRCode", Request.Form["BSRCode"].ToString());

            list.Add("@ChallanAmount", Request.Form["ChallanAmount"].ToString());

            list.Add("@Remmarks", Request.Form["Remarks"].ToString());
            list.Add("@RoundOff", 0);
            list.Add("@AmountInRound", 0);
            list.Add("@Balance", 0);
            list.Add("@CreatedBy", Session["EmpId"].ToString());

            string mes = "";

            try

            {

                mes = comfun.executeNonQueryWMessage("TallyTdsChallanDetails_AcceptUpdate  ", "", list).ToString();

            }

            catch (Exception ex) { mes = ex.Message; }
            return Json(mes);

        }

        #endregion

        public ActionResult Dispatch()
        {
            return View();
        }
        public ActionResult _DispatchList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@DeptBillId", Request.Form["DeptBillId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthId"].ToString());
            dt = comfun.fillDataTable("TallyDispatchInv_List ", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _DispatchListGet()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@DeptBillId", Request.Form["DeptBillId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthId"].ToString());
            dt = comfun.fillDataTable("TallyDispatchInv_List ", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _DispatchSubmit()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();

            string msg = "";
            try
            {
                list.Add("@DispatchId", Request.Form["DispatchId"].ToString());
                list.Add("@DeptBIllId", Request.Form["DeptBIllId"].ToString());
                list.Add("@DispatchNo", Request.Form["DispatchNo"].ToString());
                list.Add("@OfficeAddressId", Request.Form["OfficeAddressId"].ToString());
                list.Add("@OfficeAdddress", Request.Form["OfficeAdddress"].ToString());
                list.Add("@DisptchBy", Session["EmpId"].ToString());
                msg = comfun.executeNonQueryWMessage("TallyDispatchInv_AcceptUpdate", "", list).ToString();

            }
            catch (Exception ex) { msg = ex.Message; }

            return Json(msg);
        }
        #region NewDept
        public ActionResult DeptatmentInvoice()
        {
            return View();
        }
        public ActionResult _DeptatmentInvoiceList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthId"].ToString());
            list.Add("@CreateBy", Session["EmpId"].ToString()); 
             dt = comfun.fillDataTable("TallyDeptInv_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        #endregion NewDept
        public ActionResult PurchaseBillVerification()
        {
            return View();
        }
        public JsonResult _PurchaseBillVerificationList()

        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
            list.Add("@MonthId", Request.Form["MonthId"].ToString());
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@DeptId", Request.Form["DeptFilter"].ToString());
            dt = comfun.fillDataTable("TallyPurchaseVerification_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult PurchaseBillVerificationSubmit()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();

            string msg = "";
            try
            {
                list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
                list.Add("@IsPurchaseBillVerified", Request.Form["IsPurchaseBillVerified"].ToString());
                list.Add("@VerificationRemarks", Request.Form["VerificationRemarks"].ToString());
                list.Add("@Description", Request.Form["Description"].ToString());
                list.Add("@Narration", Request.Form["Narration"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                msg = comfun.executeNonQueryWMessage("TallyPurchaseVarification_AcceptUpdate", "", list).ToString();

            }
            catch (Exception ex) { msg = ex.Message; }

            return Json(msg);
        }
        #region Epf ESI
        public ActionResult MapChalaninvoice()
        {
            return View();
        }

        public ActionResult ChalanForddl()
        {
            SortedList list = new SortedList();
            list.Add("@TypeId", "0");
            DataTable dt = comfun.fillDataTable("TallyChallanType_List", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }

        public ActionResult Chalanddl()
        {
            SortedList list = new SortedList();
            list.Add("@ChallanTypeId", Request.Form["ChallanTypeId"].ToString());
            list.Add("@ChallanId", Request.Form["ChallanId"].ToString());
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthYear"].ToString());
            DataTable dt = comfun.fillDataTable("TallyEsiEpfChallan_dll", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }

        public ActionResult EpfesiList()
        {
            SortedList list = new SortedList();
            list.Add("@ChallanId", Request.Form["ChallanId"].ToString());
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthYear"].ToString());
            list.Add("@ChallanType", Request.Form["ChallanType"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@RoleId", Session["RoleId"].ToString());
            DataTable dt = comfun.fillDataTable("TallyEsiEpfChallanVsBill_List", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }

        public ActionResult submitEpfEsiMap(string ChallanId, string MapEpfEsiData)
        {
            string mes = string.Empty;
            try
            {


                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(MapEpfEsiData);

                DataTable table = new DataTable();
                table.Columns.Add("AgencyBillId", typeof(int));
                table.Columns.Add("ChallanFor", typeof(int));
                table.Columns.Add("NoOfResource", typeof(int));
                table.Columns.Add("ChallanId", typeof(int));
                table.Columns.Add("MonthYearId", typeof(int));

                foreach (var Item in jsonData)
                {
                    DataRow dr = table.NewRow();
                    dr["AgencyBillId"] = Item.AgencyBillId;
                    dr["ChallanFor"] = Item.ChallanFor;
                    dr["NoOfResource"] = Item.NoOfResource;
                    dr["ChallanId"] = Item.ChallanId;
                    dr["MonthYearId"] = Item.MonthYearId;
                    table.Rows.Add(dr);
                }

                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "tallyESiEPInvoiceMap_AcceptUpdate";

                com.Parameters.AddWithValue("@ChallanId", ChallanId);

                SqlParameter structuredParam = new SqlParameter("@TallyEsiEpfMap", SqlDbType.Structured);
                structuredParam.Value = table;
                com.Parameters.Add(structuredParam);

                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);

                com.CommandTimeout = 0;

                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();
                com.ExecuteNonQuery();
                conObj.con.Close();

                mes = sp.Value.ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("SaleList", "Error: " + ex.Message);
            }

            return Json(mes);
        }
        #region Deposite Chalan ESIEPF
        public ActionResult DepositeChalanESIEPF()
        {
            return View();
        }
        public ActionResult _ChallanBankddl()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            //list.Add("@Id", Request.Form["Id"].ToString());

            dt = comfun.fillDataTable("tblTallyBank_list", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _AgencyNameDDL()
        {
            DataTable dt = new DataTable();

            SortedList list = new SortedList();

            list.Add("@AgencyId", Session["EmpId"].ToString());

            dt = comfun.fillDataTable("TallyAgencyDdl_List", "", list);

            var json = JsonConvert.SerializeObject(dt);

            return Json(json);

        }
        public JsonResult _DepositeChallanEsiEpfSubmit()
        {
            string mes = string.Empty;
            try
            {
                var Challan_Doc = "";
                var _comPath = "";
                var filePathA = "";
                var ChallanDetails_Doc = "";
                var filePathA2 = "";
                HttpFileCollectionBase files = Request.Files;

                if (files != null && files.Count > 0 && Request.Files.AllKeys.Any())
                {

                    for (int i = 0; i < 1; i++)
                    {
                        HttpPostedFileBase file = files[i];


                        if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                        {
                            string[] testfiles = file.FileName.Split(new char[] { '\\' });
                            Challan_Doc = testfiles[testfiles.Length - 1];
                        }
                        else
                        {
                            Challan_Doc = file.FileName;
                        }
                        var _ext = Path.GetExtension(Challan_Doc);
                        Challan_Doc = Path.GetFileNameWithoutExtension(Challan_Doc);
                        string filePath = Path.Combine(Server.MapPath("/ChallanDoc/") + Challan_Doc + _ext);
                        Challan_Doc = Challan_Doc + _ext;
                        //SaleOrder_Doc = filePath;
                        filePathA = filePath;
                        file.SaveAs(filePathA);
                    }


                    if (files.Count > 1)
                    {
                        HttpPostedFileBase file2 = files[1];

                        if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                        {
                            string[] testfiles = file2.FileName.Split(new char[] { '\\' });
                            ChallanDetails_Doc = testfiles[testfiles.Length - 1];
                        }
                        else
                        {
                            ChallanDetails_Doc = file2.FileName;
                        }

                        var _ext2 = Path.GetExtension(ChallanDetails_Doc);
                        ChallanDetails_Doc = Path.GetFileNameWithoutExtension(ChallanDetails_Doc);
                        string filePath2 = Path.Combine(Server.MapPath("/ChallanDetailsDoc/") + ChallanDetails_Doc + _ext2);
                        ChallanDetails_Doc = ChallanDetails_Doc + _ext2;
                        filePathA2 = filePath2;
                        file2.SaveAs(filePathA2);
                    }
                }

                DataTable dt = new DataTable();
                SortedList list = new SortedList();
                list.Add("@ChallanId", Request.Form["ChallanId"].ToString());
                list.Add("@ChallanFor", Request.Form["ChallanFor"].ToString());
                list.Add("@ChallanNumber", Request.Form["ChallanNumber"].ToString());
                list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
                list.Add("@BillForMonth", Request.Form["BillForMonth"].ToString());
                list.Add("@BankName", Request.Form["BankName"].ToString());
                list.Add("@ChallanDate", Request.Form["ChallanDate"].ToString());
                list.Add("@Amount", Request.Form["Amount"].ToString());
                list.Add("@NoOfResource", Request.Form["NoOfResource"].ToString());
                list.Add("@AttacheChallan", Challan_Doc);
                list.Add("@AttacheChallanDetails", ChallanDetails_Doc);
                list.Add("@IsDeclaration", Request.Form["IsDeclaration"].ToString());

                list.Add("@UploadedBy", Session["EmpId"].ToString());
                mes = comfun.executeNonQueryWMessage("TallyEsiEpfChallan_AcceptUpdate ", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        public ActionResult _ChallanEsiEpfEdit()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@ChallanId", Request.Form["ChallanId"].ToString());
            list.Add("@ChallanType", 0);
            list.Add("@AgencyId", 0);
            list.Add("@MonthYear", 0);
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@Userrole", Session["RoleId"].ToString());

            dt = comfun.fillDataTable("TallyEsiEpfChallan_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _ChallanEsiEpfList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@ChallanId", Request.Form["ChallanId"].ToString());
            list.Add("@ChallanType", Request.Form["ChallanType"].ToString());
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthYear"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@Userrole", Session["RoleId"].ToString());

            dt = comfun.fillDataTable("TallyEsiEpfChallan_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        #endregion

        public ActionResult VerifyEPFESI()
        {

            return View();
        }

        public ActionResult _VerifyEPFESIlist()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@ChallanId", Request.Form["ChallanId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthYear"].ToString());
            list.Add("@ChallanType", Request.Form["ChllanType"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@Userrole", Session["RoleId"].ToString());
            dt = comfun.fillDataTable("TallyEsiEpfChallan_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _VerifyEPFESIlistModel()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@ChallanId", Request.Form["ChallanId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthYear"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@RoleId", Session["RoleId"].ToString());
            dt = comfun.fillDataTable("TallyEsiEpfChallanVsBill_List ", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _VerifyEPFESIlistGet()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@ChallanId", Request.Form["ChallanId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthYear"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@Userrole", Session["RoleId"].ToString());
            dt = comfun.fillDataTable("TallyEsiEpfChallan_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _VerifyEPFESISubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();

                list.Add("@ChallanId", Request.Form["ChallanId"].ToString());
                list.Add("@IsVarified", Request.Form["IsVarified"].ToString());
                // list.Add("@CancelRemarks", Request.Form["CancelRemarks"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                mes = comfun.executeNonQueryWMessage("TallyEsiEpfChallanVerified_AcceptUPdate  ", "", list).ToString();
            }
            catch (Exception ex)

            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        #endregion Epf/ESI
        public ActionResult DepartmentNewPayment()
        {

            return View();
        }
        public ActionResult AgencyNewPayment()
        {

            return View();
        }
        public ActionResult TdsReportNew()
        {
            return View();
        }
        public ActionResult TdsReportNewList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthId"].ToString());
            list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
            dt = comfun.fillDataTable("tallyAgencyTds_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult DispatchReport()
        {
            return View();
        }
        public ActionResult _DispatchReportList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencId", Request.Form["AgencId"].ToString());
            list.Add("@DeptBIllId", Request.Form["DeptBIllId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@FromDate", Request.Form["FromDate"].ToString()); 
            list.Add("@ToDate", Request.Form["ToDate"].ToString());
            dt = comfun.fillDataTable("TallyDisptchAddress_list", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public string SendEmail1()
        {
            string m = "";
            try
            {
                DataTable dt = new DataTable();
                SortedList list = new SortedList();
                list.Add("@AgencyBillId", Request.Form["BillId"].ToString());
                dt = comfun.fillDataTable("TallySaleInvoiceDownload_URL", "", list);

                string email = "Wfms@horizontelecom.in";
                string password = "wfms@9002"; ;
                string displayName = "HPSEDC";

                string subject = " Invoice for services delivered | Invoice No:-" + dt.Rows[0]["SaleBillNo"].ToString();// + "| Date :-"+ dt.Rows[0]["BillDate"].ToString() +"  Month Of :"+ dt.Rows[0]["MonthId"].ToString()+'-'+ dt.Rows[0]["Year"].ToString();
                string message = "Sir/Ma'am<br>Please find enclosed the invoice for the outsource employees services provided to your office per your engagement letter or equivalent document(s).Please note,this is a digitally signed, electronic invoice.<br><br> An early settlement of the invoice will be much appreciated.<br><br> Please Click on the following link to download your invoice.<br>" + dt.Rows[0]["Url"].ToString();

                message = message + "<br><br> Regards <br>HPSEDC";

                var loginInfo = new NetworkCredential(email, password);
                var msg = new MailMessage();
                var smtpClient = new SmtpClient("smtp.rediffmailpro.com");
                smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtpClient.Port = 587;

                msg.From = new MailAddress("hrms@wfms.in", displayName);
                msg.To.Add(new MailAddress(dt.Rows[0]["EmailId"].ToString()));
                msg.CC.Add(new MailAddress("hpsedc@gmail.com"));

                msg.Subject = subject;
                msg.Body = message;
                msg.ReplyTo = new MailAddress("Hpsedc@hpsedc.in");
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
        #region User
        public ActionResult CreateUser()
        {
            return View();
        }
        public ActionResult UserRoleDdl()
        {
            SortedList list = new SortedList();
            list.Add("@RollId", Request.Form["Id"].ToString());
            DataTable dt = comfun.fillDataTable("stpTallyUserRole_List", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult _UserRoleSubmit()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@UserName", Request.Form["UserName"].ToString());
            list.Add("@Email", Request.Form["Email"].ToString());
            list.Add("@Mobile", Request.Form["Mobile"].ToString());
            list.Add("@RollId", Request.Form["RoleId"].ToString());
            list.Add("@Password", Request.Form["Password"].ToString());

            string mes = "";
            try
            {
                mes = comfun.executeNonQueryWMessage("stpTallyUserPassword_AcceptUpdate", "", list).ToString();

            }
            catch (Exception ex) { mes = ex.Message; }

            return Json(mes);
        }
        public ActionResult UserRoleList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@EmailId", Request.Form["EmailId"].ToString());
            dt = comfun.fillDataTable("stpTallyUserPassword_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
      
        public ActionResult ChangePassword()
        {
            return View();
        }
        public ActionResult _ChangePasswordSubmit()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList(); 
            list.Add("@EmpID", Session["EmpID"].ToString());
            list.Add("@OldPassword", Request.Form["OldPassword"].ToString());
            list.Add("@UserPassword", Request.Form["confirmpassword"].ToString());

            string mes = "";
            try
            {
                mes = comfun.executeNonQueryWMessage("stpTallyChangePassword", "", list).ToString();

            }
            catch (Exception ex) { mes = ex.Message; }

            return Json(mes);
        }
        #endregion User
        #region ESIEPF Report

        public ActionResult EsiEpfReport()
        {
            return View();
        }

        public ActionResult _ESIEPFReportlist()
        {
            SortedList list = new SortedList();
            list.Add("@ChallanId", Request.Form["ChallanId"].ToString());
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthYear"].ToString());
            list.Add("@ChallanType", Request.Form["ChallanType"].ToString());
            list.Add("@EsiEpfDeposited", Request.Form["EsiEpfDeposited"].ToString());
            DataTable dt = comfun.fillDataTable("TallyEsiEpfReport_list", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        #endregion
        #region
        public ActionResult EmployeeDetailsDetilsList()
        {
            return View();
        }
        public ActionResult _AvoDepartmentddlEmployee()
        {
            SortedList list = new SortedList();
            list.Add("@DeptId", Session["DeptId"]?.ToString() ?? "0");
            list.Add("@createdBy", Session["EmpId"].ToString());
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            DataTable dt = comfun.fillDataTable("stpHPSEDCSSODepartmentDDL1", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult EmployeeList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgenyId", Request.Form["AgencyId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@workorderid", Request.Form["Work_Order"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@RoleId", Session["RoleId"].ToString());
            dt = comfun.fillDataTable("stpTallyEmployees_list1", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult EmployeeListGet()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["NemEmpId"].ToString());
            dt = comfun.fillDataTable("stpTallyEmployees_Get", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult EmployeeDetsilsUpdateSubmit()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpIdnew"].ToString());
            list.Add("@Empname", Request.Form["Employee"].ToString());
            list.Add("@FatherName", Request.Form["Father"].ToString());
            list.Add("@Email", Request.Form["Email"].ToString());
            list.Add("@ContactNo", Request.Form["Mobile"].ToString());
            list.Add("@IsFullTime", Request.Form["ISFullTime"].ToString());
            list.Add("@DesigationId", Request.Form["Designation"].ToString());
            list.Add("@AADHARNO", Request.Form["Addhar"].ToString());
            list.Add("@BasicSalary", Request.Form["Basic"].ToString());
            list.Add("@OthersAllowance", Request.Form["Others"].ToString());
            list.Add("@IsPf", Request.Form["chkEpf"].ToString());
            list.Add("@IsEsi", Request.Form["chkEsic"].ToString());
            list.Add("@AcNO", Request.Form["Account"].ToString());
            list.Add("@Ifsc", Request.Form["Ifse"].ToString());
            list.Add("@UANNo", Request.Form["Uan"].ToString());
            list.Add("@ESICNo", Request.Form["Esic"].ToString());
            list.Add("@EducationId", Request.Form["Education"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            string mes = "";
            try
            {
                mes = comfun.executeNonQueryWMessage("stpTallyEmployees_AcceptUPdate", "", list).ToString();

            }
            catch (Exception ex) { mes = ex.Message; }

            return Json(mes);
        }
        public ActionResult DesignationDll()
        {
            SortedList list = new SortedList();
            list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
            DataTable dt = comfun.fillDataTable("TallyDesignation_List", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult _WorkOrderDdlEmployee()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyId", Session["EmpId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@WorkOrderAgencyId", Request.Form["WorkOrderAgencyId"].ToString());
            list.Add("@UserId", Session["EmpId"].ToString());
            list.Add("@RoleId", Session["RoleId"].ToString());
            dt = comfun.fillDataTable("TallyAgencyWorkOrder1_get2", "", list);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public ActionResult EducationDll()
        {
            SortedList list = new SortedList();
            list.Add("@EducationID", Request.Form["EducationID"].ToString());
            DataTable dt = comfun.fillDataTable("TallyEducation_List", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult DesignationListDownload()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@fiDesignationID", Request.Form["fiDesignationID"].ToString());
            dt = comfun.fillDataTable("stp_GetActiveDesignations_IsActive", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        #region Add Designation
        public ActionResult DesignationView()
        {
            return View();
        }


        public ActionResult DesignationList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@fiDesignationID", Request.Form["fiDesignationID"].ToString());
            dt = comfun.fillDataTable("stp_GetActiveDesignations", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult DesignationListUpdate()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@fiDesignationID", Request.Form["fiDesignationID"].ToString());
                list.Add("@fvIsActive", Request.Form["fvIsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("DesignationStatusUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }
        public ActionResult DesignationSubmit()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();

            list.Add("@fvDesignationName", Request.Form["DesignationId"].ToString()); 
            list.Add("@fvIsActive", Request.Form["IsActive"].ToString());            
            list.Add("@createdBy", Session["EmpId"].ToString());                    

            string mes = "";
            try
            {
                mes = comfun.executeNonQueryWMessage("TallyHpsedcDesignation_Accept", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }

            return Json(mes);
        }

        //public JsonResult _DesigantionNameEdit()
        //{
        //    DataTable dt = new DataTable();
        //    SortedList list = new SortedList();
        //    list.Add("@Id", Request.Form["Id"].ToString());
        //    dt = comfun.fillDataTable("PiCompany_Edit", "", list);
        //    var json = JsonConvert.SerializeObject(dt);
        //    return Json(json);
        //}
        #endregion

        public ActionResult EmployeeDetsilsUpdateSubmitNew()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpIdnew"].ToString());
            list.Add("@DeptId", Request.Form["DeptIdEmployee"].ToString());
            list.Add("@AgencyId", Request.Form["ddlAgencynewEmployee"].ToString());
            list.Add("@WorkOrderNo", Request.Form["Work_OrderEmployee"].ToString());
            list.Add("@TotalManpower", Request.Form["TotalManpower"].ToString());
            list.Add("@Empname", Request.Form["Employee"].ToString());
            list.Add("@FatherName", Request.Form["Father"].ToString());
            list.Add("@Email", Request.Form["Email"].ToString());
            list.Add("@ContactNo", Request.Form["Mobile"].ToString());
            list.Add("@IsFullTime", Request.Form["ISFullTime"].ToString());
            list.Add("@DesigationId", Request.Form["Designation"].ToString());
            list.Add("@EducationId", Request.Form["Education"].ToString());
            list.Add("@AADHARNO", Request.Form["Addhar"].ToString());
            list.Add("@BasicSalary", Request.Form["Basic"].ToString());
            list.Add("@OthersAllowance", Request.Form["Others"].ToString());
            list.Add("@IsPf", Request.Form["chkEpf"].ToString());
            list.Add("@IsEsi", Request.Form["chkEsic"].ToString());
            list.Add("@AcNO", Request.Form["Account"].ToString());
            list.Add("@Ifsc", Request.Form["Ifse"].ToString());
            list.Add("@UANNo", Request.Form["Uan"].ToString());
            list.Add("@ESICNo", Request.Form["Esic"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            string mes = "";
            try
            {
                mes = comfun.executeNonQueryWMessage("stpTallyEmployees_AcceptUPdateNew", "", list).ToString();

            }
            catch (Exception ex) { mes = ex.Message; }

            return Json(mes);
        }
        #endregion
        public ActionResult TotalResourceReportCompanyWise()
        {
            return View();
        }
        public ActionResult _TotalResourceReportCompanyWiseList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@BillforMonthList", Request.Form["MonthId"].ToString());
            dt = comfun.fillDataTable("SumOfResources_ByMonth", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult HpsedcOutStandingLetter()
        {
            return View();
        }
        public JsonResult HpsedcOutStandingLetterList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@BillingAddress", Request.Form["BillingId"].ToString()); 
            list.Add("@MonthYear", Request.Form["MonthYear"].ToString());
            dt = comfun.fillDataTable("TallyDepartmentBalanceBill_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json, JsonRequestBehavior.AllowGet);
        }
        public ActionResult DepartmentOutstandingView()
        {
            return View();
        }
        public JsonResult DepartmentOutstandingViewList()

        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@BillId", Request.Form["BillId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthId"].ToString());
            list.Add("@BillingAddress", Request.Form["BillingAddress"].ToString());
            dt = comfun.fillDataTable("TallyDepartmentBalance_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult AddressDDlnew()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyId", "0");
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@WorkOrderAgencyId", Request.Form["WorkOrderAgencyId"].ToString());
            dt = comfun.fillDataTable("TallyAgencyWorkOrder3_list", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public ActionResult ManpowerDepartmentAdvancePayment()
        {
            return View();
        }
        public JsonResult ManpowerDepartmentAdvanceList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@BillId", Request.Form["BillId"].ToString());
            list.Add("@MonthYear", Request.Form["MonthId"].ToString());
            list.Add("@BillingAddress", Request.Form["BillingAddress"].ToString());
            dt = comfun.fillDataTable("TallyDepartmentBalance_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult ManpowerPartialPaymentReport()
        {
            return View();
        }
    }
}