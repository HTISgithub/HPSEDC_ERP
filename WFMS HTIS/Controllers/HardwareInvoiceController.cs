//using iTextSharp.text.pdf;
//using iTextSharp.text.pdf;
using Newtonsoft.Json;

using Payroll.portal.SessionLogout;
using Payroll.portal.Models;
using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using MasterIndia;
using MasterIndia.Model;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Web.Script.Serialization;

namespace Payroll.portal.Controllers
{
    [SessionExpire]
    public class HardwareInvoiceController : Controller
    {
        commonFunctions comfun = new commonFunctions();
        payrollFunctions payfun = new payrollFunctions();
        MasterIndiaFunction masterIndia = new MasterIndiaFunction();

        // GET: Deployment
        #region  Hardware
        public ActionResult SaleOrder()
        {
            return View();
        }

        public ActionResult SaleOrder1()
        {
            return View();
        }

        public ActionResult SaleInovice()
        {
            return View();
        }

        public ActionResult SaleOrderDetail()
        {
            return View();
        }

        public ActionResult SaleInovice1(int? id)
        {
            ViewBag.InvoiceId = id;
            return View();
        }

        public ActionResult SaleInoviceDublicate(int? id)
        {
            ViewBag.InvoiceId = id;
            return View();
        }

        public ActionResult SaleInoviceTriplate(int? id)
        {
            ViewBag.InvoiceId = id;
            return View();
        }

        public ActionResult ViewDeptDetail()
        {
            return View();
        }

        public ActionResult _SaleOrderList()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@RoleId", Session["RoleId"].ToString());
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@PurchaseIssue", Request.Form["PurchaseIssue"].ToString());
            dt = comfun.fillDataTable("HardwareSaleOrder_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _HarwareDepartmentddl()
        {
            SortedList list = new SortedList();
            list.Add("@DeptId", "0");
            DataTable dt = comfun.fillDataTable("TallyAgencyDepartment1_list", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult _ItemDetailList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
            list.Add("@PurchaseOrderId", Request.Form["PurchaseOrderId"].ToString());
            list.Add("@SaleOrderType", Request.Form["SaleOrderType"].ToString());
            list.Add("@UserRole", Session["RoleId"].ToString());
            dt = comfun.fillDataTable("HardwareProductSaleOrderExportPIssue_list", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _HardwareDetail()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
            //list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@UserRole", Session["RoleId"].ToString());
            dt = comfun.fillDataTable("HardwareProductSaleOrderExportPIssue_list", "", list);
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
            dt = comfun.fillDataTable("HardwareAgencyDdl_List", "", list);
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
        //public ActionResult _RequisitionDetailsEdit()
        //{
        //    //comfun.saveformname("_purchaseBillList", "/pm/_purchaseBillList", "Voucher/PurchaseBill", "Purchase Bill List", "N", "PurchaseBill", "PurchaseBill", "List", 4);
        //    DataTable dt = new DataTable();
        //    SortedList list = new SortedList();
        //    list.Add("@PostId", Request.Form["Id"].ToString());
        //    dt = comfun.fillDataTable("HpsedcRequisitionDetails_Edit", "", list);
        //    var json = JsonConvert.SerializeObject(dt);
        //    return Json(json);
        //}
        public ActionResult _ProductDetailsDelete()
        {
            //comfun.saveformname("_purchaseBillList", "/pm/_purchaseBillList", "Voucher/PurchaseBill", "Purchase Bill List", "N", "PurchaseBill", "PurchaseBill", "List", 4);
            string mes = string.Empty;
            SortedList list = new SortedList();

            try
            {
                list.Add("@ProductId", Request.Form["ProductId"].ToString());

                mes = comfun.executeNonQueryWMessage("HardwareDeleveryAddress_Delete", "", list).ToString();
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


        public JsonResult _SaleOrderSubmit(string RequestDetails, string SaleOrderId, string SaleOrderNo, string SaleOrderNoText, string OrderDate, string DeptId, string BillingAddressId, string BillingAddressText, string ReferenceNo, string DeliveryDate, string Total, string Cgst, string Sgst, string Gst, string AdminCharge,
         string Gtotal, string PaymentAmt, string Balance, string IsPaymentRequired)
        {
            string mes = string.Empty;
            try
            {
                var SaleOrder_Doc = "";
                var _comPath = "";
                var filePathA = "";
                var HardwareDelivery_Doc = "";
                var filePathA2 = "";
                HttpFileCollectionBase files = Request.Files;

                if (System.Web.HttpContext.Current.Request.Files.AllKeys.Any())
                {

                    for (int i = 0; i < 1; i++)
                    {
                        HttpPostedFileBase file = files[i];


                        if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                        {
                            string[] testfiles = file.FileName.Split(new char[] { '\\' });
                            SaleOrder_Doc = testfiles[testfiles.Length - 1];
                        }
                        else
                        {
                            SaleOrder_Doc = file.FileName;
                        }
                        var _ext = Path.GetExtension(SaleOrder_Doc);
                        SaleOrder_Doc = Path.GetFileNameWithoutExtension(SaleOrder_Doc);
                        string filePath = Path.Combine(Server.MapPath("/HWSaleOrderDoc/") + SaleOrder_Doc + _ext);
                        SaleOrder_Doc = SaleOrder_Doc + _ext;
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
                            HardwareDelivery_Doc = testfiles[testfiles.Length - 1];
                        }
                        else
                        {
                            HardwareDelivery_Doc = file2.FileName;
                        }

                        var _ext2 = Path.GetExtension(HardwareDelivery_Doc);
                        HardwareDelivery_Doc = Path.GetFileNameWithoutExtension(HardwareDelivery_Doc);
                        string filePath2 = Path.Combine(Server.MapPath("/HWDeliveryDoc/") + HardwareDelivery_Doc + _ext2);
                        HardwareDelivery_Doc = HardwareDelivery_Doc + _ext2;
                        //HardwareDelivery_Doc = filePath2;
                        filePathA2 = filePath2;
                        file2.SaveAs(filePathA2);
                    }
                }


                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(RequestDetails);

                DataTable table = new DataTable();
                table.Columns.Add("ProductId", typeof(int));
                table.Columns.Add("OrderQty", typeof(double));
                table.Columns.Add("Price", typeof(double));
                table.Columns.Add("Gst", typeof(double));
                table.Columns.Add("AdminCharge", typeof(double));
                table.Columns.Add("Gtotal", typeof(double));
                table.Columns.Add("Narration", typeof(string));

                foreach (var Item in jsonData)
                {
                    DataRow dr = table.NewRow();
                    dr["ProductId"] = Item.ProductCategoryId;
                    dr["OrderQty"] = Item.Quantity;
                    dr["Price"] = Item.Rate;
                    dr["Gst"] = Item.GST;
                    dr["AdminCharge"] = Item.AdminCharge;
                    dr["Gtotal"] = Item.Total;
                    dr["Narration"] = Item.Narration;

                    table.Rows.Add(dr);
                }

                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "HardwareSaleOrder_AcceptUpdate";
                SqlParameter parameter = new SqlParameter();

                com.Parameters.AddWithValue("@SaleOrderId", SaleOrderId);
                com.Parameters.AddWithValue("@SaleOrderNo", 0);
                com.Parameters.AddWithValue("@SaleOrderNoText", 0);
                com.Parameters.AddWithValue("OrderDate", OrderDate);
                com.Parameters.AddWithValue("DeptId", DeptId);
                com.Parameters.AddWithValue("BillingAddressId", BillingAddressId);
                com.Parameters.AddWithValue("BillingAddressText", BillingAddressText);
                com.Parameters.AddWithValue("LetterReferenceNo", ReferenceNo);
                com.Parameters.AddWithValue("DeliveryDate", DeliveryDate);
                com.Parameters.AddWithValue("Total", Total);
                com.Parameters.AddWithValue("Cgst", 0);
                com.Parameters.AddWithValue("Sgst", 0);
                com.Parameters.AddWithValue("Gst", Gst);
                com.Parameters.AddWithValue("AdminCharge", AdminCharge);
                com.Parameters.AddWithValue("Gtotal", Gtotal);
                com.Parameters.AddWithValue("PaymentAmt", 0);
                com.Parameters.AddWithValue("Balance", 0);
                com.Parameters.AddWithValue("@IsPaymentRequired", IsPaymentRequired);
                com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());
                com.Parameters.AddWithValue("@Attachement", SaleOrder_Doc);
                com.Parameters.AddWithValue("@DeliveryAttachement", HardwareDelivery_Doc); // Add the second file

                SqlParameter parameter1 = new SqlParameter();
                parameter1.ParameterName = "@tblTempHardwarSaleOrder2";
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

        public JsonResult _SaleOrderDetailsSubmit(string RequestDetails)
        {

            string mes = string.Empty;
            try
            {
                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(RequestDetails);

                DataTable table = new DataTable();

                table.Columns.Add("ProductId", typeof(int));
                table.Columns.Add("DeliveryQty", typeof(int));
                table.Columns.Add("consigneeAddress", typeof(string));
                foreach (var Item in jsonData)
                {
                    DataRow dr = table.NewRow();
                    dr["ProductId"] = Item.ProductNameId;
                    dr["DeliveryQty"] = Item.DeliveryQuantity;
                    dr["consigneeAddress"] = Item.consigneeAddress;

                    table.Rows.Add(dr);
                }
                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "HardwareSaleOrderDetails_AcceptUpdate";
                SqlParameter parameter1 = new SqlParameter();
                //com.Parameters.AddWithValue("@SaleOrderId", SaleOrderId);

                parameter1.ParameterName = "@tblTempHardwarSaleOrder3";
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
        public ActionResult PiMainCategory_ddl()
        {
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            DataTable dt = comfun.fillDataTable("PiMainCategory_list", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }

        public ActionResult PIProductCategory_ddl()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@CategoryId", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("PiCategory_ddl", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }

        public ActionResult HardwareProduct_list()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId", Session["Customer"]);
            list.Add("@MainCategoryId", Request.Form["MainCategoryId"].ToString());
            list.Add("@PCategoryId", Request.Form["PCategoryId"].ToString());
            list.Add("@CompanyId", Request.Form["CompanyId"].ToString());

            dt = comfun.fillDataTable("HardwareProduct_list", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult HardwareProductDetails_list()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@ProductId", Request.Form["ProductId"].ToString());
            dt = comfun.fillDataTable("HardwareProductDetails_list", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public ActionResult _PaymentModeddl()
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
        public ActionResult calculateAdminCharge()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@ProductId", Request.Form["ProductId"].ToString());
            list.Add("@Qty", Request.Form["Qty"].ToString());

            dt = comfun.fillDataTable("HardwareProductDetailsQty_list", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }


        public ActionResult PIProductCategory_ddl1()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@PurchaseOrdrId", Request.Form["Id"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            dt = comfun.fillDataTable("HardwareProductDetailsProduct_list", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public JsonResult _NewAgencyPaymentList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
            dt = comfun.fillDataTable("HardwareSaleOrderDetails_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult PurchaseOrderIssue()
        {
            String msg = "";
            SortedList list = new SortedList();
            try
            {
                list.Add("@OrderDetailsId", Request.Form["OrderId"].ToString());
                list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                msg = comfun.executeNonQueryWMessage("HardwarePurchaseOrder_Issue1", "", list).ToString();
            }
            catch (Exception e)
            {
                msg = e.Message;
            }
            return Json(msg);
        }

        public ActionResult ProductDeliveryList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
            list.Add("@ProductId", Request.Form["ProductId"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            dt = comfun.fillDataTable("HardwarePurchaseOrderDelivery_list", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult ProductDeliverySubmit()
        {
            String mes = "";
            try
            {

                var ProductDelivery_ActionDoc = "";

                var ProductDeliveryActionImage1 = "";
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
                            ProductDeliveryActionImage1 = testfiles[testfiles.Length - 1];

                        }
                        else
                        {
                            ProductDeliveryActionImage1 = file.FileName;
                        }
                        var _ext = Path.GetExtension(ProductDeliveryActionImage1);
                        ProductDelivery_ActionDoc = Path.GetFileNameWithoutExtension(ProductDeliveryActionImage1);
                        string filePath = Path.Combine(Server.MapPath("/ProductDeliveryDoc/") + ProductDelivery_ActionDoc + _ext);
                        ProductDelivery_ActionDoc = ProductDelivery_ActionDoc + _ext;
                        ProductDeliveryActionImage1 = filePath;
                        filePathA = filePath;
                        file.SaveAs(filePathA);

                    }

                }

                SortedList list = new SortedList();
                list.Add("@OrderDetailsId", Request.Form["OrderDetailsid"].ToString());
                list.Add("@DeliveredQty", Request.Form["DeliveryQty"].ToString());
                list.Add("@DeliveredTo", Request.Form["DeliveryTo"].ToString());
                list.Add("@DeliveryDoc", ProductDelivery_ActionDoc);
                list.Add("@Createdby", Session["EmpId"].ToString());

                mes = comfun.executeNonQueryWMessage("HardwarePurchaseOrderDeliv_AcceptUpdate", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }

            return Json(mes);
        }



        public ActionResult BillingAddress()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@DeptId", Request.Form["DepId"].ToString());
            dt = comfun.fillDataTable("HardwareBillingAddress_ddl", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult GetAdvancePayDetail()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
            dt = comfun.fillDataTable("HardwareSaleOrderAdvance_get", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult submitAdvancePayment()
        {
            SortedList list = new SortedList();
            string Mes = "";
            try
            {

                list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
                list.Add("@TransactionId", Request.Form["transactionId"].ToString());
                list.Add("@ReceivedDate", Request.Form["transactionDate"].ToString());
                list.Add("@ModeOfPayment", Request.Form["modeofpayment"].ToString());
                list.Add("@Narration", Request.Form["narration"].ToString());
                list.Add("@ReceivedAmount", Request.Form["advanceAmount"].ToString());
                list.Add("@GSTTds", Request.Form["GstTds"].ToString());
                list.Add("@Tds", Request.Form["tds"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                list.Add("@BankNameId", Request.Form["BankNameId"].ToString());
                list.Add("@BalanceAmt", Request.Form["BalanceAmt"].ToString());
                Mes = comfun.executeNonQueryWMessage("HardwareSaleOrderAdvance_AcceptUpdate", "", list).ToString();
            }
            catch (Exception ex) { Mes = ex.Message; }
            return Json(Mes);
        }



        #endregion


        #region new

        public ActionResult NewInvoice()
        {
            return View();
        }

        public ActionResult _NewInvoiceSubmit()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
            list.Add("@BillDate", Request.Form["BillDate"].ToString());
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
            list.Add("@CreatedBy", Session["EmpId"].ToString());

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
            dt = comfun.fillDataTable("TallyAgencyBill_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _NewInvoiceListget()
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
            dt = comfun.fillDataTable("HardwareAgencyDdl_List", "", list);
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


            string mes = "";
            try
            {
                mes = comfun.executeNonQueryWMessage("TallyDepartmentBill_AcceptUpdate", "", list).ToString();
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
            list.Add("@DeptBillId", "1");
            list.Add("@AgencyBillId", "0");
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
                list.Add("@TransactionId", Request.Form["Transaction"].ToString());
                list.Add("@ModeOfPayment", Request.Form["ModePayment"].ToString());
                list.Add("@Narration", Request.Form["Remarks"].ToString());
                list.Add("@ReceivedDate", Request.Form["PaymentDate"].ToString());
                list.Add("@ReceivedAmt", Request.Form["PaymentAmount"].ToString());
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
                list.Add("@ModeOfPayment", Request.Form["ModePayment1"].ToString());
                list.Add("@Narration", Request.Form["Remarks1"].ToString());
                list.Add("@PaymentDate", Request.Form["PaymentDate1"].ToString());
                list.Add("@PaymentAmt", Request.Form["PaymentAmount1"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                Mes = comfun.executeNonQueryWMessage("TallyAgencyPayment_AcceptUpdate", "", list).ToString();
            }
            catch (Exception ex) { Mes = ex.Message; }
            return Json(Mes);
        }
        //public JsonResult _NewAgencyPaymentList()
        //{
        //    DataTable dt = new DataTable();
        //    SortedList list = new SortedList();
        //    list.Add("@PaymentId", Request.Form["PaymentId"].ToString());
        //    list.Add("@AgencyBillId", Request.Form["AgencyBillId"].ToString());
        //    dt = comfun.fillDataTable("TallyAgencyPaymentTransaction_List", "", list);
        //    var json = JsonConvert.SerializeObject(dt);
        //    return Json(json);
        //}

        public ActionResult NewInvoice1()
        {
            return View();
        }
        #endregion

        #region HardwarePurchase
        public ActionResult OrderDetail()
        {
            return View();
        }

        public ActionResult PurchaseOrder()
        {
            return View();
        }
        public ActionResult PurchaseOrderDetailList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@RoleId", Session["RoleId"].ToString());
            list.Add("@PurchaseOrderNoId", 0);
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            dt = comfun.fillDataTable("HardwarePurchaseOrder_List1", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _PurchaseOrderDetailAcceptReject_List()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@RoleId", Session["RoleId"].ToString());
            list.Add("@PurchaseOrderNoId", 0);
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            dt = comfun.fillDataTable("HardwarePurchaseOrderAcceptReject_List ", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult GetPurchaseOrderDetail()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("HardwarePurchaseOrder_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult PurchaseOrderDetail()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@PurchaseOrderId", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("HardwarePurchaseOrder_Details", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult GetOrderDetail()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@PurchaseOrderNoId", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("HardwarePurchaseOrderDetails_List1", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult AgencyPaymentviewList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@PruchaseOrderId", Request.Form["PruchaseOrderId"].ToString());
            list.Add("@PruchaseInvId", 0);
            dt = comfun.fillDataTable("HardwareAgencyParymentTransaction_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        #endregion

        #region Hardware Master

        #region Vishu
        public ActionResult HardwareProductCategory()
        {
            return View();
        }
        public ActionResult _HardwareProductCategorySubmit()
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
                        string filePath = Path.Combine(Server.MapPath("/PCategory/") + Requisition_ActionDoc + _ext);
                        Requisition_ActionDoc = Requisition_ActionDoc + _ext;
                        RequisitionActionImage1 = filePath;
                        filePathA = filePath;
                        file.SaveAs(filePathA);

                    }

                }

                DataTable dt = new DataTable();
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@Title", Request.Form["Tittle"].ToString());
                list.Add("@MainCatgId", Request.Form["MainCatgId"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                list.Add("@Attachement", Requisition_ActionDoc);

                mes = comfun.executeNonQueryWMessage("HardwareCategory_AcceptUpdate", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }

            return Json(mes);
        }
        public ActionResult _HardwareProductCategoryList()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            // list.Add("@CustomerId", Session["Customer"]);
            list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("HardwareProductategory_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _HardwareProductCategoryEdit()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();

            list.Add("@Id", Request.Form["Id"].ToString());

            dt = comfun.fillDataTable("HardwareProductategory_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult PMainCategory_ddl()
        {
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            DataTable dt = comfun.fillDataTable("PiMainCategory_list", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public JsonResult _HardwareProductStatusUpdate()
        {

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpECategoryStatusUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }



        #endregion


        #region Vishu

        public ActionResult Product()
        {
            SortedList list = new SortedList();
            DataTable dt = comfun.fillDataTable("PiPCompnay_List", "", list);
            return View(dt);
        }
        public ActionResult ProductSubmit()
        {
            string mes = "";
            try
            {

                var Requisition_ActionDoc = "";

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
                        string filePath = Path.Combine(Server.MapPath("/ProductImages/") + Requisition_ActionDoc + _ext);
                        Requisition_ActionDoc = Requisition_ActionDoc + _ext;
                        RequisitionActionImage1 = filePath;
                        filePathA = filePath;
                        file.SaveAs(filePathA);

                    }

                }

                DataTable dt = new DataTable();
                SortedList list = new SortedList();
                list.Add("@ProductId", Request.Form["ProductId"].ToString());
                list.Add("@MainCategoryId", Request.Form["MainCategoryId"].ToString());
                list.Add("@PCategoryId", Request.Form["Category"].ToString());
                list.Add("@CompanyId", Request.Form["CompanyID"].ToString());
                list.Add("@Sepcification", Request.Form["Specification"].ToString());
                list.Add("@CurrentStorck", Request.Form["CurrentStock"].ToString());
                list.Add("@HPSEDCCharges", Request.Form["HpsedcCharge"].ToString());
                list.Add("@ProductPrice", Request.Form["ProductPrice"].ToString());
                list.Add("@GrandTotal", Request.Form["GrandTotal"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                list.Add("@ModelNo", Request.Form["modelno"].ToString());
                list.Add("@HSNCode", Request.Form["HSNCode"].ToString());
                list.Add("@Gst", Request.Form["Gst"].ToString());
                list.Add("@ImageFile", Requisition_ActionDoc);

                list.Add("@Createdby", Session["EmpId"].ToString());

                mes = comfun.executeNonQueryWMessage("HardwareProduct_AcceptUpdate", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }

            return Json(mes);
        }


        public ActionResult ProductCompany_ddl()
        {
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            DataTable dt = comfun.fillDataTable("PiCompany_ddl", "", null);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult Product_List()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());

            dt = comfun.fillDataTable("HardwareMasterProduct_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult Product_Edit()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());

            dt = comfun.fillDataTable("HardwareMasterProduct_Edit", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }
        public JsonResult ProductStatusUpdate()
        {

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpPiProductStatusUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }

        public ActionResult PIPCategory_ddl()
        {
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            list.Add("@MainCatgId", Request.Form["MainCatgId"].ToString());

            DataTable dt = comfun.fillDataTable("HardwareProductategory_List", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult PiMainCategory_ddl3()
        {
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            DataTable dt = comfun.fillDataTable("PiMainCategory_list", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }

        #endregion



        #region Vishu
        public ActionResult AddCompany()
        {
            return View();
        }
        public ActionResult _AddCompanySubmit()
        {
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            list.Add("@CompanyName", Request.Form["CompanyName"].ToString());
            list.Add("@Description", Request.Form["Description"].ToString());
            list.Add("@IsActive", Request.Form["IsActive"].ToString());

            string mes = comfun.executeNonQueryWMessage("PiCompany_AcceptUpdate", "", list).ToString();

            return Json(mes);
        }
        public ActionResult _AddCompanyList()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("PiCompany_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public JsonResult _AddCompanyEdit()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("PiCompany_Edit", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _CompanyStatusUpdate()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("CompanyStatusUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }

        public ActionResult SubmitProductDelivery()
        {
            string mes = string.Empty;
            string pod = "";
            string ir = "";

            SortedList list = new SortedList();
            try
            {
                HttpFileCollectionBase files = Request.Files;
                HttpPostedFileBase firstFile = files[0];
                string firstFileName = Path.GetFileName(firstFile.FileName);
                if (files != null && (files.Count == 1) && Request.Files.AllKeys.Any() && firstFileName.Contains("IR"))
                {
                    if (firstFile != null && firstFile.ContentLength > 0)
                    {
                        string fileName = Path.GetFileName(firstFile.FileName);
                        string ext = Path.GetExtension(fileName);
                        string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);

                        string uploadPath = Server.MapPath("~/InstallationReport/");
                        string finalPath = Path.Combine(uploadPath, nameWithoutExt + ext);

                        firstFile.SaveAs(finalPath);

                        ir = nameWithoutExt + ext;

                    }
                }

                else if (files != null && files.Count > 0 && Request.Files.AllKeys.Any())
                {
                    string[] folders = { "~/ProofOfDelivery/", "~/InstallationReport/" };
                    string[] fileNames = new string[3];


                    for (int i = 0; i < (files.Count); i++)
                    {
                        HttpPostedFileBase file = files[i];

                        if (file != null && file.ContentLength > 0)
                        {
                            string fileName = Path.GetFileName(file.FileName);
                            string ext = Path.GetExtension(fileName);
                            string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);

                            string uploadPath = Server.MapPath(folders[i]);
                            string finalPath = Path.Combine(uploadPath, nameWithoutExt + ext);

                            file.SaveAs(finalPath);

                            fileNames[i] = nameWithoutExt + ext;

                        }
                    }

                    if (fileNames[1] != null)
                    {
                        pod = fileNames[0];
                        ir = fileNames[1];

                    }
                    else
                    {
                        pod = fileNames[0];
                    }
                }

                list.Add("@ItemDetailsId", Request.Form["OrderDetailId"].ToString());
                list.Add("@DeliveredQty", Request.Form["DeliveredQty"].ToString());
                list.Add("@DeliveryId", Request.Form["DeliveryId"].ToString());
                list.Add("@DeliveredDate", Request.Form["DeliveredDate"].ToString());
                list.Add("@DeliveredTo", " ");
                list.Add("@Document1", pod);
                list.Add("@Document2", ir);
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                mes = comfun.executeNonQueryWMessage("HardwareDeliveryAddress_AcceptUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }

        public ActionResult HardwareProductDeliveryList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@DeliveryId", Request.Form["Id"].ToString());
            list.Add("@PurchaseOrderId", Request.Form["PurchaseOrderId"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            dt = comfun.fillDataTable("HardwareProductDelivery_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        //public ActionResult HardWareDeliveryItem_List()
        //{
        //    DataTable dt = new DataTable();
        //    SortedList list = new SortedList();
        //    list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
        //    dt = comfun.fillDataTable("HardwareEnterDeliveryItem_List", "", list);
        //    var json = JsonConvert.SerializeObject(dt);
        //    return Json(json);



        //}



        public ActionResult HardWareDeliveryItem_List()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
            list.Add("@PurchaseOrderId", Request.Form["PurchaseOrderId"].ToString());
            //dt = comfun.fillDataTable("HardwareEnterDeliveryItem_List", "", list);
            dt = comfun.fillDataTable("HardwareEnterDeliveryItem_List1", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);



        }

        public ActionResult HardwareDeliveryLocationList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
            list.Add("@PurchaseOrderId", Request.Form["PurchaseOrderId"].ToString());
            // list.Add("@CreatedBy", Session["EmpId"].ToString());
            //dt = comfun.fillDataTable("HardwareEnterDeliveryAddress_List", "", list);
            dt = comfun.fillDataTable("HardwareEnterDeliveryAddress_List1", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        //public ActionResult HardwareDeliveryLocationList()
        //{
        //    DataTable dt = new DataTable();
        //    SortedList list = new SortedList();
        //    list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
        //    list.Add("@PurchaseOrderId", Request.Form["@PurchaseOrderId"].ToString());
        //    // list.Add("@CreatedBy", Session["EmpId"].ToString());
        //    dt = comfun.fillDataTable("HardwareEnterDeliveryAddress_List", "", list);
        //    var json = JsonConvert.SerializeObject(dt);
        //    return Json(json);
        //}

        public ActionResult HardWareDeliveryItem_get()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@OrderDetailsId", Request.Form["OrderDetailsId"].ToString());
            dt = comfun.fillDataTable("HardwareEnterDeliveryItemDetails_Get", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);

        }



        public ActionResult ProductDeliveryLocationSubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@OrderDeliveryId", Request.Form["OrderDeliveryId"].ToString());
                list.Add("@ItemDetailsId", Request.Form["ItemDetailsId"].ToString());
                list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
                list.Add("@ProductId", Request.Form["ProductId"].ToString());
                list.Add("@DeliveryQty", Request.Form["DeliveryQty"].ToString());
                list.Add("@consigneeAddress", Request.Form["consigneeAddress"].ToString());
                list.Add("@ConsigneeName", Request.Form["ConsigneeName"].ToString());
                list.Add("@ConsigneeContactNo", Request.Form["ConsigneeContactNo"].ToString());
                //list.Add("@CreatedBy", Session["EmpId"].ToString());
                mes = comfun.executeNonQueryWMessage("HardwareEnterDeliveryAddress_AcceptUpdate", "", list).ToString();
            }
            catch (Exception ex)

            {
                mes = ex.Message;
            }
            return Json(mes);
        }

        public ActionResult SubmitPurchaseBill()
        {
            string mes = string.Empty;
            string purchaseInvoiceFile = "";

            SortedList list = new SortedList();
            try
            {
                HttpFileCollectionBase files = Request.Files;

                if (files != null && files.Count > 0 && Request.Files.AllKeys.Any())
                {

                    HttpPostedFileBase file = files[0];

                    if (file != null && file.ContentLength > 0)
                    {
                        string fileName = Path.GetFileName(file.FileName);
                        string ext = Path.GetExtension(fileName);
                        string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);

                        string uploadPath = Server.MapPath("~/PurchaseInvoice/");
                        string finalPath = Path.Combine(uploadPath, nameWithoutExt + ext);

                        file.SaveAs(finalPath);

                        purchaseInvoiceFile = nameWithoutExt + ext;

                    }

                }

                list.Add("@Id", 0);
                list.Add("@SaleOrderId", Request.Form["SaleOrderNo"].ToString());
                list.Add("@PurchaseOrderNo", Request.Form["PurchaseOrderNo"].ToString());
                list.Add("@PurchaseInvNo", Request.Form["PurchaseInvoiceNo"].ToString());
                list.Add("@InvDate", Request.Form["PurchaseBillDate"].ToString());
                list.Add("@Narration", Request.Form["Narration"].ToString());
                list.Add("@BaseAmount", Request.Form["BaseAmount"].ToString());
                list.Add("@Gst", Request.Form["Gst"].ToString());
                list.Add("@Cgst", Request.Form["Cgst"].ToString());
                list.Add("@Sgst", Request.Form["Sgst"].ToString());
                list.Add("@Total", Request.Form["Total"].ToString());
                list.Add("@AgencyAddressId", Request.Form["AgencyAddressId"].ToString());
                list.Add("@InvAttachement", purchaseInvoiceFile);

                list.Add("@CreatedBy", Session["EmpId"].ToString());
                mes = comfun.executeNonQueryWMessage("spHardwarePurchaseInv_AcceptUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }

        public ActionResult GetSaleInvoiceData()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@SaleOrderId", Request.Form["SaleId"].ToString());
            dt = comfun.fillDataTable("HardwareSaleOrderAmount_Get", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        //public ActionResult GetConsigneeAddress()
        //{
        //    DataTable dt = new DataTable();
        //    SortedList list = new SortedList();
        //    list.Add("@ItemDetailsId", Request.Form["OrderDetailId"].ToString());
        //    list.Add("@PurchaseOrderNo", Request.Form["PurchaseOrderNo"].ToString());
        //    dt = comfun.fillDataTable("HardwareDeliveryAddress_List", "", list);
        //    var json = JsonConvert.SerializeObject(dt);
        //    return Json(json);
        //}

        public ActionResult GetDeliveryDetail()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@OrderDetailsId", Request.Form["OrderDetailId"].ToString());
            dt = comfun.fillDataTable("HardwareDeliveryAddress_Get", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult SubmitVerificationDetail()
        {
            String msg = "";
            SortedList list = new SortedList();
            try
            {
                list.Add("@PurchareOrderNo", Request.Form["PurchaseOrderNo"].ToString());
                list.Add("@IsPODVerified", Request.Form["pod"].ToString());
                list.Add("@IsIRVerified", Request.Form["ir"].ToString());
                list.Add("@VerifedBy", Session["EmpId"].ToString());
                msg = comfun.executeNonQueryWMessage("HardwarePODIRVerification_AcceptUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                msg = ex.Message;
            }
            return Json(msg);
        }


        public ActionResult BillForDdl()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            dt = comfun.fillDataTable("HardwareAgencyBillOffice_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        //public JsonResult SubmitPurchaseBill1(string RequestDetails, string SaleOrderNo, string InvNo, string AgencyBillingId, string PurchaseOrderNo, string InvDate, string TotalAmount,
        //   string Narration, string IgstAmount, string CgstAmount, string SgstAmount, string GrandTotal)
        //{
        //    string mes = string.Empty;
        //    try
        //    {
        //        var PurchaseInvoice_Doc = "";
        //        var _comPath = "";
        //        var filePathA = "";
        //        //var HardwareDelivery_Doc = "";
        //        //var filePathA2 = "";
        //        HttpFileCollectionBase files = Request.Files;

        //        if (System.Web.HttpContext.Current.Request.Files.AllKeys.Any())
        //        {

        //            for (int i = 0; i < 1; i++)
        //            {
        //                HttpPostedFileBase file = files[i];


        //                if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
        //                {
        //                    string[] testfiles = file.FileName.Split(new char[] { '\\' });
        //                    PurchaseInvoice_Doc = testfiles[testfiles.Length - 1];
        //                }
        //                else
        //                {
        //                    PurchaseInvoice_Doc = file.FileName;
        //                }
        //                var _ext = Path.GetExtension(PurchaseInvoice_Doc);
        //                PurchaseInvoice_Doc = Path.GetFileNameWithoutExtension(PurchaseInvoice_Doc);
        //                string filePath = Path.Combine(Server.MapPath("/PurchaseInvoice/") + PurchaseInvoice_Doc + _ext);
        //                PurchaseInvoice_Doc = PurchaseInvoice_Doc + _ext;
        //                filePathA = filePath;
        //                file.SaveAs(filePathA);

        //            }

        //        }
        //        dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(RequestDetails);

        //        DataTable table = new DataTable();
        //        table.Columns.Add("ProductId", typeof(int));
        //        table.Columns.Add("Quantity", typeof(int));
        //        table.Columns.Add("HSNCode", typeof(string));
        //        table.Columns.Add("Rate", typeof(double));
        //        table.Columns.Add("Gst", typeof(double));
        //        table.Columns.Add("AdminCharge", typeof(double));
        //        table.Columns.Add("Total", typeof(double));

        //        foreach (var Item in jsonData)
        //        {
        //            DataRow dr = table.NewRow();
        //            dr["ProductId"] = Item.ProductId;
        //            dr["Quantity"] = Item.Quantity;
        //            dr["HSNCode"] = Item.HSNCode;
        //            dr["Rate"] = Item.Rate;
        //            dr["Gst"] = Item.GST;
        //            dr["AdminCharge"] = Item.AdminCharge;
        //            dr["Total"] = Item.Total;

        //            table.Rows.Add(dr);
        //        }

        //        SqlCommand com = new SqlCommand();
        //        connection conObj = new connection();
        //        com.Connection = conObj.con;
        //        com.CommandType = CommandType.StoredProcedure;
        //        com.CommandText = "HardwarePurchaseInv_AcceptUpdate";
        //        SqlParameter parameter = new SqlParameter();

        //        com.Parameters.AddWithValue("@PInvId", 0);
        //        com.Parameters.AddWithValue("@InvNo", InvNo);

        //        com.Parameters.AddWithValue("@SaleOrderNo", SaleOrderNo);
        //        com.Parameters.AddWithValue("@AgencyBillingId", AgencyBillingId);
        //        com.Parameters.AddWithValue("PurchaseOrderNo", PurchaseOrderNo);
        //        com.Parameters.AddWithValue("InvDate", InvDate);
        //        com.Parameters.AddWithValue("Narration", Narration);

        //        com.Parameters.AddWithValue("TotalAmount", TotalAmount);
        //        com.Parameters.AddWithValue("IgstAmount", IgstAmount);
        //        com.Parameters.AddWithValue("CgstAmount", CgstAmount);
        //        com.Parameters.AddWithValue("SgstAmount", SgstAmount);
        //        com.Parameters.AddWithValue("GrandTotal", GrandTotal);

        //        com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());
        //        com.Parameters.AddWithValue("@InvAttachement", PurchaseInvoice_Doc);

        //        SqlParameter parameter1 = new SqlParameter();
        //        parameter1.ParameterName = "@TempHardwarePurchaseInv2";
        //        parameter1.SqlDbType = System.Data.SqlDbType.Structured;
        //        parameter1.Value = table;
        //        com.Parameters.Add(parameter1);

        //        SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
        //        sp.Direction = ParameterDirection.Output;
        //        com.Parameters.Add(sp);

        //        com.CommandTimeout = 0;

        //        if (conObj.con.State == ConnectionState.Closed)
        //            conObj.con.Open();
        //        com.ExecuteNonQuery();
        //        conObj.con.Close();

        //        mes = "Record saved successfully";
        //        mes = sp.Value.ToString();
        //        string retunvalue = (string)com.Parameters["@Mes"].Value;
        //    }
        //    catch (Exception ex)
        //    {
        //        mes = comfun.errorMessage("POList", "Error: " + ex.Message);
        //    }
        //    return Json(mes);
        //}

        //    public JsonResult SubmitSaleBill(string SaleData, string SaleOrderId, string SaleInvoiceNo, string PurchaseInvoiceNo, string InvoiceNo,
        //string AgencyBillingId, string PurchaseOrderNo, string InvDate, string TotalAmount,
        //string Igst, string cgst, string sgst, string AdminCharge, string Gtotal)
        //    {
        //        string mes = string.Empty;
        //        try
        //        {


        //            dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(SaleData);

        //            DataTable table = new DataTable();
        //            table.Columns.Add("ProductId", typeof(int));
        //            table.Columns.Add("Quantity", typeof(int));
        //            table.Columns.Add("HSNCode", typeof(string));
        //            table.Columns.Add("Rate", typeof(double));
        //            table.Columns.Add("Gst", typeof(double));
        //            table.Columns.Add("AdminCharge", typeof(double));
        //            table.Columns.Add("Total", typeof(double));

        //            foreach (var Item in jsonData)
        //            {
        //                DataRow dr = table.NewRow();
        //                dr["ProductId"] = Item.ProductId;
        //                dr["Quantity"] = Item.Quantity;
        //                dr["HSNCode"] = Item.HSNCode;
        //                dr["Rate"] = Item.Rate;
        //                dr["Gst"] = Item.GST;
        //                dr["AdminCharge"] = Item.AdminCharge;
        //                dr["Total"] = Item.Total;

        //                table.Rows.Add(dr);
        //            }

        //            SqlCommand com = new SqlCommand();
        //            connection conObj = new connection();
        //            com.Connection = conObj.con;
        //            com.CommandType = CommandType.StoredProcedure;
        //            com.CommandText = "HardwareSasleInv_AcceptUpdate";

        //            com.Parameters.AddWithValue("@PurchaseOrderNo", PurchaseOrderNo);
        //            com.Parameters.AddWithValue("@SInvId", SaleInvoiceNo);
        //            com.Parameters.AddWithValue("@PInvId", PurchaseInvoiceNo);
        //            com.Parameters.AddWithValue("@InvNo", InvoiceNo);
        //            com.Parameters.AddWithValue("@AgencyBillingId", AgencyBillingId);
        //            com.Parameters.AddWithValue("@SaleOrderId", SaleOrderId);
        //            com.Parameters.AddWithValue("@InvDate", InvDate);
        //            com.Parameters.AddWithValue("@TotalAmount", TotalAmount);
        //            com.Parameters.AddWithValue("@IgstAmount", Igst);
        //            com.Parameters.AddWithValue("@CgstAmount", cgst);
        //            com.Parameters.AddWithValue("@SgstAmount", sgst);
        //            com.Parameters.AddWithValue("@AdminCharge", AdminCharge);
        //            com.Parameters.AddWithValue("@GrandTotal", Gtotal);
        //            com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());


        //            SqlParameter structuredParam = new SqlParameter("@TempHardwarePurchaseInv2", SqlDbType.Structured);
        //            structuredParam.Value = table;
        //            com.Parameters.Add(structuredParam);

        //            SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
        //            sp.Direction = ParameterDirection.Output;
        //            com.Parameters.Add(sp);

        //            com.CommandTimeout = 0;

        //            if (conObj.con.State == ConnectionState.Closed)
        //                conObj.con.Open();
        //            com.ExecuteNonQuery();
        //            conObj.con.Close();

        //            mes = sp.Value.ToString();
        //        }
        //        catch (Exception ex)
        //        {
        //            mes = comfun.errorMessage("SaleList", "Error: " + ex.Message);
        //        }

        //        return Json(mes);
        //    }

        public ActionResult SaleInvoiceList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
            list.Add("@PurchaseId", Request.Form["PurchaseId"].ToString());
            list.Add("@SinvId", Request.Form["SinvId"].ToString());
            dt = comfun.fillDataTable("spHardwareSaleInvoice_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        //public ActionResult SubmitVerificationDetailPODIR()
        //{
        //    string mes = string.Empty;
        //    try
        //    {
        //        SortedList list = new SortedList();
        //        list.Add("@OrderDetailsId", Request.Form["OrderDetailsId"].ToString());
        //        list.Add("@IsPODVerified", Request.Form["IsPODVerified"].ToString());
        //        list.Add("@IsIRVerified", Request.Form["IsIRVerified"].ToString());
        //        list.Add("@VerifedBy", Session["EmpId"].ToString());

        //        mes = comfun.executeNonQueryWMessage("HardwarePODIRVerificationSingle_AcceptUpdate", "", list).ToString();
        //    }
        //    catch (Exception ex)

        //    {
        //        mes = ex.Message;
        //    }
        //    return Json(mes);
        //}

        public ActionResult PurchaseInvoiceView()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@SaleOrderNo", 0);
            list.Add("@PurchaseOrderNoId", Request.Form["PurchaseId"].ToString());
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@RoleId", Session["RoleId"].ToString());
            dt = comfun.fillDataTable("HardwarePurchaseOrder_InvList", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult SaleInvoiceView()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@PurchaseOrderNo", Request.Form["PurchaseId"].ToString());
            list.Add("@SaleOrderId", Request.Form["Saleorderid"].ToString());
            dt = comfun.fillDataTable("stpHardwareSaleInv_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        #endregion

        #endregion

        #region agencypayment

        public ActionResult GetAgencyPaymentData()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@PruchaseInvId", Request.Form["PruchaseInvId"].ToString());
            dt = comfun.fillDataTable("HardwareAgencyPayment_Get", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult ModeOfPaymentDdl()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@Id", 0);
            dt = comfun.fillDataTable("HpsedcPaymentmodeDdl_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult AgencypaymentList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@PruchaseOrderId", Request.Form["PurchaseOrderNoId"].ToString());
            dt = comfun.fillDataTable("HardwareAgencyParymentTransaction_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult SubmitAgencyPayment()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@PruchaseInvId", Request.Form["PruchaseInvId"].ToString());
                list.Add("@PurchaseOrderNoId", Request.Form["PurchaseOrderNoId"].ToString());
                list.Add("@TransactionId", Request.Form["TransactionId"].ToString());
                list.Add("@ModeOfPayment", Request.Form["ModeOfPayment"].ToString());
                list.Add("@Narration", Request.Form["Narration"].ToString());
                list.Add("@PaymentDate", Request.Form["PaymentDate"].ToString());
                list.Add("@GstTds2", Request.Form["GstTds2"].ToString());
                list.Add("@Tds1", Request.Form["Tds1"].ToString());
                list.Add("@PaymentId", Request.Form["PaymentId"].ToString());
                list.Add("@PaymentAmt", Request.Form["PaymentAmt"].ToString());
                list.Add("@Penalty", Request.Form["Penalty"].ToString());
                list.Add("@Pbg", Request.Form["PBG"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());

                mes = comfun.executeNonQueryWMessage("HardwareAgencyPayment_AcceptUpdate", "", list).ToString();
            }
            catch (Exception ex)

            {
                mes = ex.Message;
            }
            return Json(mes);
        }

        public ActionResult _Trackingstatusddl()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@StatusId", Request.Form["StatusId"].ToString());
            dt = comfun.fillDataTable("HardwareWorkOrderStatus_list", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _TrackingStatusSubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@PurchaseOrderId", "0");
                list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
                list.Add("@ActionId", Request.Form["ActionId"].ToString());
                list.Add("@Narration", Request.Form["Narration"].ToString());
                list.Add("@ActionTakenBy", Session["EmpId"].ToString());
                mes = comfun.executeNonQueryWMessage("TrackingStatus_AcceptUpdate", "", list).ToString();
            }
            catch (Exception ex)

            {
                mes = ex.Message;
            }
            return Json(mes);
        }

        public ActionResult _TrackingStatusList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
            dt = comfun.fillDataTable("HardwareTrackingStatus_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        #endregion

        #region Cancellation

        public ActionResult _SaleorderCancellationGetData()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
            dt = comfun.fillDataTable("HardwareSaleOrderCancel_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }


        public ActionResult _CancellationSubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();

                list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
                list.Add("@IsCancel", Request.Form["IsCancel"].ToString());
                list.Add("@CancelRemarks", Request.Form["CancelRemarks"].ToString());
                list.Add("@CancelBy", Session["EmpId"].ToString());
                mes = comfun.executeNonQueryWMessage("HardwareSaleOrderCancel_AcceptUpdate", "", list).ToString();
            }
            catch (Exception ex)

            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        #endregion


        #region Location Address Delete
        public ActionResult _LocationAddressDelete()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();

                list.Add("@OrderDeliveryId", Request.Form["OrderDeliveryId"].ToString());
                mes = comfun.executeNonQueryWMessage("HardwareEnterDeliveryAddress_Removed", "", list).ToString();
            }
            catch (Exception ex)

            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        #endregion

        #region SaleOrderUpdate
        public ActionResult _SaleOrderEdit()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@SaleOrderId", Request.Form["Id"].ToString());
            //list.Add("@CreatedBy", Session["EmpId"].ToString());
            //list.Add("@UserRole", Session["RoleId"].ToString());
            DataSet ds = comfun.fillDataSet("HardwareSaleOrder_Get", "", list);

            var json = JsonConvert.SerializeObject(ds);
            return Json(json);
        }
        public ActionResult _SaleOrderUpdate(int? id)
        {
            ViewBag.SaleOrderId = id;
            return View();
        }

        public JsonResult _SaleOrderSubmitUpdate(string RequestDetails, string SaleOrderId, string SaleOrderNo, string SaleOrderNoText, string OrderDate, string DeptId, string BillingAddressId, string BillingAddressText, string ReferenceNo, string DeliveryDate, string Total, string Cgst, string Sgst, string Gst, string AdminCharge,
         string Gtotal, string PaymentAmt, string Balance, string IsPaymentRequired, string Attachement, string DeliveryDocument)
        {
            string mes = string.Empty;
            try
            {
                var SaleOrder_Doc = "";
                var _comPath = "";
                var filePathA = "";
                var HardwareDelivery_Doc = "";
                var filePathA2 = "";
                HttpFileCollectionBase files = Request.Files;

                if (files != null && System.Web.HttpContext.Current.Request.Files.AllKeys.Any())
                {

                    for (int i = 0; i < 1; i++)
                    {
                        HttpPostedFileBase file = files[i];


                        if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                        {
                            string[] testfiles = file.FileName.Split(new char[] { '\\' });
                            HardwareDelivery_Doc = testfiles[testfiles.Length - 1];
                        }
                        else
                        {
                            HardwareDelivery_Doc = file.FileName;
                        }
                        var _ext = Path.GetExtension(HardwareDelivery_Doc);
                        HardwareDelivery_Doc = Path.GetFileNameWithoutExtension(HardwareDelivery_Doc);
                        string filePath = Path.Combine(Server.MapPath("/HWDeliveryDoc/") + HardwareDelivery_Doc + _ext);
                        HardwareDelivery_Doc = HardwareDelivery_Doc + _ext;
                        //SaleOrder_Doc = filePath;
                        filePathA2 = filePath;
                        file.SaveAs(filePathA2);
                        SaleOrder_Doc = Attachement;
                    }

                }
                else
                {
                    SaleOrder_Doc = Attachement;
                    HardwareDelivery_Doc = DeliveryDocument;
                }


                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(RequestDetails);

                DataTable table = new DataTable();
                table.Columns.Add("ProductId", typeof(int));
                table.Columns.Add("OrderQty", typeof(double));
                table.Columns.Add("Price", typeof(double));
                table.Columns.Add("Gst", typeof(double));
                table.Columns.Add("AdminCharge", typeof(double));
                table.Columns.Add("Gtotal", typeof(double));
                table.Columns.Add("Narration", typeof(string));

                foreach (var Item in jsonData)
                {
                    DataRow dr = table.NewRow();
                    dr["ProductId"] = Item.ProductCategoryId;
                    dr["OrderQty"] = Item.Quantity;
                    dr["Price"] = Item.Rate;
                    dr["Gst"] = Item.GST;
                    dr["AdminCharge"] = Item.AdminCharge;
                    dr["Gtotal"] = Item.Total;
                    dr["Narration"] = Item.Narration;

                    table.Rows.Add(dr);
                }

                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "HardwareSaleOrderUpdate_AcceptUpdate";
                SqlParameter parameter = new SqlParameter();

                com.Parameters.AddWithValue("@SaleOrderId", SaleOrderId);
                com.Parameters.AddWithValue("@SaleOrderNo", SaleOrderNo);
                com.Parameters.AddWithValue("@SaleOrderNoText", SaleOrderNoText);
                com.Parameters.AddWithValue("OrderDate", OrderDate);
                com.Parameters.AddWithValue("DeptId", DeptId);
                com.Parameters.AddWithValue("BillingAddressId", BillingAddressId);
                com.Parameters.AddWithValue("BillingAddressText", BillingAddressText);
                com.Parameters.AddWithValue("LetterReferenceNo", ReferenceNo);
                com.Parameters.AddWithValue("DeliveryDate", DeliveryDate);
                com.Parameters.AddWithValue("Total", Total);
                com.Parameters.AddWithValue("Cgst", 0);
                com.Parameters.AddWithValue("Sgst", 0);
                com.Parameters.AddWithValue("Gst", Gst);
                com.Parameters.AddWithValue("AdminCharge", AdminCharge);
                com.Parameters.AddWithValue("Gtotal", Gtotal);
                com.Parameters.AddWithValue("PaymentAmt", 0);
                com.Parameters.AddWithValue("Balance", 0);
                com.Parameters.AddWithValue("@IsPaymentRequired", IsPaymentRequired);
                com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());
                com.Parameters.AddWithValue("@Attachement", SaleOrder_Doc);
                com.Parameters.AddWithValue("@DeliveryAttachement", HardwareDelivery_Doc); // Add the second file

                SqlParameter parameter1 = new SqlParameter();
                parameter1.ParameterName = "@tblTempHardwarSaleOrder2";
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

        #endregion

        #region Billing Address
        public ActionResult DepartmentBillingAddress()
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
        public ActionResult DepartmentBillingAddressSubmit()

        {
            DataTable dt = new DataTable();

            SortedList list = new SortedList();
            string msg = "";

            try

            {
                list.Add("@BillingAddressId", Request.Form["BillingAddressId"].ToString());

                list.Add("@DeptId", Request.Form["DeptId"].ToString());

                list.Add("@DistrictId", Request.Form["DistrictId"].ToString());

                list.Add("@BillingAddress", Request.Form["BillingAddress"].ToString());

                // list.Add("@CreatedBy", Session["EmpId"].ToString());

                msg = comfun.executeNonQueryWMessage("HardwareBillingAddress_AcceptUpdate", "", list).ToString();

            }

            catch (Exception ex) { msg = ex.Message; }

            return Json(msg);

        }


        public JsonResult DepartmentBillingAddressList()

        {

            DataTable dt = new DataTable();

            SortedList list = new SortedList();

            list.Add("@BillingAddressId", Request.Form["BillingAddressId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@DistrictId", Request.Form["DistrictId"].ToString());



            dt = comfun.fillDataTable("HardwareBillingAddress_List", "", list);

            var json = JsonConvert.SerializeObject(dt);

            return Json(json);

        }

        public JsonResult DepartmentBillingAddressEdit()

        {

            DataTable dt = new DataTable();

            SortedList list = new SortedList();

            list.Add("@BillingAddressId", Request.Form["BillingAddressId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            list.Add("@DistrictId", Request.Form["DistrictId"].ToString());


            dt = comfun.fillDataTable("HardwareBillingAddress_List", "", list);

            var json = JsonConvert.SerializeObject(dt);

            return Json(json);

        }

        #endregion

        #region PurchasePaymentReport

        public ActionResult HardwareAgencyReport()
        {
            return View();
        }
        public JsonResult _HardwareAgencyReportList()
        {

            DataTable dt = new DataTable();

            SortedList list = new SortedList();

            list.Add("@AgencyId", Request.Form["AgencyId"].ToString());
            list.Add("@isSalegenerated", Request.Form["IsSaleGenerated"].ToString());
            list.Add("@min", Request.Form["min"].ToString());
            list.Add("@max", Request.Form["max"].ToString());
            //dt = comfun.fillDataTable("HardwarePurchaseOrderReport_List", "", list);
            dt = comfun.fillDataTable("HardwareAgencyPaymentReport_list", "", list);

            var json = JsonConvert.SerializeObject(dt);

            return Json(json);

        }

        public ActionResult _HardwareAgencyReportDetailList()
        {
            DataTable dt = new DataTable();

            SortedList list = new SortedList();

            list.Add("@InvoiceId", Request.Form["Invoiceid"].ToString());

            dt = comfun.fillDataTable("HardwareAgencyPaymentReportDetail_list", "", list);

            var json = JsonConvert.SerializeObject(dt);

            return Json(json);
        }
        #endregion

        #region EInvoice1

        public ActionResult EInvoice()

        {

            return View();

        }

        public JsonResult _EInvoiceList()

        {

            DataTable dt = new DataTable();

            dt = comfun.fillDataTable("stpHardwareSaleInvEInvlice_List", "", null);

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

                    var response1 = JsonConvert.DeserializeObject<MasterIndia.Model.Response>(contents);


                    var Status = response1.results.message.Status;

                    var Ack = response1.results.message.Ackno;

                    var AckDt = response1.results.message.AckDt;

                    var Irn = response1.results.message.Irn;

                    var SignedInvoice = response1.results.message.SignedInvoice;

                    var SignedQRCode = response1.results.message.SignedQRCode;



                    //save data in database

                    using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["cn"].ConnectionString))

                    {

                        using (SqlCommand cmd = new SqlCommand("HardwareInsertInvoiceData", con))

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

        #region HPSEDC Annerxture
        public ActionResult PrintHPSEDCPO(int? id, int? pid)
        {
            ViewBag.SaleOrder = id;
            ViewBag.PurchaseOrderId = pid;

            return View();
        }
        public ActionResult DeliveryAddress(int? id, int? pid)
        {
            ViewBag.SaleOrder = id;
            ViewBag.PurchaseOrderId = pid;
            return View();
        }
        public ActionResult DeliveryLocationList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
            list.Add("@PurchaseOrderId", Request.Form["PurchaseOrderId"].ToString());
            dt = comfun.fillDataTable("HardwareDeliveryList", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Content(json, "application/json");
        }

        #endregion

        #region SaleOrderVerfy
        public ActionResult SaleOrderVerfiy(int? id)
        {
            ViewBag.SaleOrderId = id;
            return View();
        }

        public ActionResult SubmitSOVerification()
        {
            SortedList list = new SortedList();
            string Mes = "";
            try
            {

                list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
                list.Add("@Verify", Request.Form["RadioDetail"].ToString());
                list.Add("@CreateBy", Session["EmpId"].ToString());
                list.Add("@Remarks", Request.Form["Rematks"].ToString());
                Mes = comfun.executeNonQueryWMessage("TallySaleOrderVerify_AcceptUpdate", "", list).ToString();
            }
            catch (Exception ex) { Mes = ex.Message; }
            return Json(Mes);
        }

        public ActionResult VerificationDetails()
        {
            DataTable dt = new DataTable();

            SortedList list = new SortedList();

            list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());

            dt = comfun.fillDataTable("HardwareVerfication_detail", "", list);

            var json = JsonConvert.SerializeObject(dt);

            return Json(json);
        }
        #endregion

        #region POVerification
        public ActionResult POVerification()
        {
            return View();

        }

        public ActionResult _AcceptRejectSubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("SaleOrderId", Request.Form["SaleOrderId"].ToString());
                list.Add("@OrderDetailsId", Request.Form["OrderDetailsId"].ToString());
                list.Add("@PurchaseOrderId", Request.Form["PurchaseOrderId"].ToString());
                list.Add("@IsOrderAccept", Request.Form["IsOrderAccept"].ToString());
                list.Add("@Remarks", Request.Form["Remarks"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                mes = comfun.executeNonQueryWMessage("HardwarePurchaseOrder_AcceptReject", "", list).ToString();
            }
            catch (Exception ex)

            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        #endregion

        #region GST Report
        public ActionResult GSTHardwareReport()
        {
            return View();
        }

        public ActionResult _GstHardwareReportList()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@MonthId", 0);
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            //list.Add("@RoleId", Session["RoleId"].ToString());
            //list.Add("@PurchaseIssue", Request.Form["PurchaseIssue"].ToString());
            dt = comfun.fillDataTable("HardwarePurchaseSaleInvoice_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        #endregion

        #region PurchaseOrderIsseueNew
        public ActionResult _PurchaseOrderIsseueSubmit(string PurchaseOrderData, string OrderDetailsId, string AgencyId, string IsUniquePruchaseOrder)
        {
            string mes = string.Empty;
            try
            {


                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(PurchaseOrderData);

                DataTable table = new DataTable();
                table.Columns.Add("OrderDetailsId", typeof(int));


                foreach (var Item in jsonData)
                {
                    DataRow dr = table.NewRow();
                    dr["OrderDetailsId"] = Item.OrderDetailsId;

                    table.Rows.Add(dr);
                }

                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "HardwarePurchaseOrder_Issue2";


                com.Parameters.AddWithValue("@OrderDetailsId", OrderDetailsId);
                com.Parameters.AddWithValue("@AgencyId", AgencyId);
                com.Parameters.AddWithValue("@IsUniquePruchaseOrder", IsUniquePruchaseOrder);

                com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());

                SqlParameter structuredParam = new SqlParameter("@OrderDetailsIdTable", SqlDbType.Structured);
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
        #endregion

        #region DeliverylocationNew

        public ActionResult HardwareDeliveryAddressItemWise_List()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
            list.Add("@PurchaseOrderNo", Request.Form["PurchaseOrderNo"].ToString());
            //dt = comfun.fillDataTable("HardwareEnterDeliveryItem_List", "", list);
            dt = comfun.fillDataTable("HardwareDeliveryAddressItemWise_List ", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);



        }

        public ActionResult _HardwareAddressLocationSubmit(string AddressDetailData, string OrderDeliveryId)
        {
            string mes = string.Empty;
            try
            {

                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(AddressDetailData);

                DataTable table = new DataTable();
                table.Columns.Add("ItemDetailsId", typeof(int));
                table.Columns.Add("SaleOrderId", typeof(int));
                table.Columns.Add("ProductId", typeof(int));
                table.Columns.Add("DeliveryQuantity", typeof(double));
                table.Columns.Add("ConsigneeName", typeof(String));
                table.Columns.Add("ConsigneeContactNo", typeof(String));
                table.Columns.Add("consigneeAddress", typeof(String));
                table.Columns.Add("DistrictId", typeof(int));
                table.Columns.Add("DeliveredQty", typeof(double));

                foreach (var Item in jsonData)
                {
                    DataRow dr = table.NewRow();
                    dr["ItemDetailsId"] = Item.ItemDetailsId;
                    dr["SaleOrderId"] = Item.SaleOrderId;
                    dr["ProductId"] = Item.ProductId;
                    dr["DeliveryQuantity"] = Item.DeliveryQuantity;
                    dr["ConsigneeName"] = Item.ConsigneeName;
                    dr["ConsigneeContactNo"] = Item.ConsigneeContactNo;
                    dr["consigneeAddress"] = Item.consigneeAddress;
                    dr["DistrictId"] = Item.DistrictId;
                    dr["DeliveredQty"] = Item.DeliveredQty;
                    table.Rows.Add(dr);
                }

                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "HardwareEnterDeliveryAddress_AcceptUpdate1";

                com.Parameters.AddWithValue("@OrderDeliveryId", OrderDeliveryId);
                // com.Parameters.AddWithValue("@AgencyId", AgencyId);

                //com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());

                SqlParameter structuredParam = new SqlParameter("@TemptblHardwarSaleOrder3", SqlDbType.Structured);
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

        #endregion

        #region TermCondition
        public ActionResult TermCondition()
        {
            return View();
        }


        public ActionResult _TermConditionSubmit()
        {

            string mes = "";
            try
            {
                var TermConditon_Doc = "";
                var _comPath = "";
                var TermConditon_Doc1 = "";
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
                            TermConditon_Doc1 = testfiles[testfiles.Length - 1];

                        }
                        else
                        {
                            TermConditon_Doc1 = file.FileName;
                        }
                        var _ext = Path.GetExtension(TermConditon_Doc1);
                        TermConditon_Doc = Path.GetFileNameWithoutExtension(TermConditon_Doc1);
                        string filePath = Path.Combine(Server.MapPath("/TermConditionDoc/") + TermConditon_Doc + _ext);
                        TermConditon_Doc = TermConditon_Doc + _ext;
                        TermConditon_Doc1 = filePath;
                        filePathA = filePath;
                        file.SaveAs(filePathA);

                    }

                }

                DataTable dt = new DataTable();
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@Title", Request.Form["Title"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                list.Add("@DocAttachment", TermConditon_Doc);

                mes = comfun.executeNonQueryWMessage("HardwarePenaltyDoc_AcceptUpdate", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }

            return Json(mes);
        }
        public ActionResult _TermConditionList()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("HardwarePenaltyDoc_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        #endregion

        #region New Method  Rausahn
        public ActionResult ConsigneeAddressDdl()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            //list.Add("@ItemDetailsId", Request.Form["OrderDetailId"].ToString());
            list.Add("@PurchaseOrderId", Request.Form["PurchaseOrderNo"].ToString());
            //dt = comfun.fillDataTable("HardwareDeliveryAddress_List", "", list);
            //dt = comfun.fillDataTable("HardwareDeliveryAddress_List1", "", list);
            dt = comfun.fillDataTable("HardwareDeliveryAddress_List2", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult PodProductList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@consigneeAddress", Request.Form["consigneeAddress"].ToString());
            list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
            list.Add("@PurchaseOrderNo", Request.Form["PurchaseOrderId"].ToString());
            dt = comfun.fillDataTable("HardwareDeliveryItemWise_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult SubmitProductDelivery1(String data, String PurchaseOrderid, String ItemDetailsIdTxt)
        {
            string mes = string.Empty;
            string pod = "";
            string ir = "";

            SortedList list = new SortedList();
            try
            {
                HttpFileCollectionBase files = Request.Files;
                HttpPostedFileBase firstFile = files[0];
                string firstFileName = Path.GetFileName(firstFile.FileName);
                if (files != null && (files.Count == 1) && Request.Files.AllKeys.Any() && firstFileName.Contains("IR"))
                {
                    if (firstFile != null && firstFile.ContentLength > 0)
                    {
                        string fileName = Path.GetFileName(firstFile.FileName);
                        string ext = Path.GetExtension(fileName);
                        string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);

                        string uploadPath = Server.MapPath("~/InstallationReport/");
                        string finalPath = Path.Combine(uploadPath, nameWithoutExt + ext);

                        firstFile.SaveAs(finalPath);

                        ir = nameWithoutExt + ext;

                    }
                }

                else if (files != null && files.Count > 0 && Request.Files.AllKeys.Any())
                {
                    string[] folders = { "~/ProofOfDelivery/", "~/InstallationReport/" };
                    string[] fileNames = new string[3];


                    for (int i = 0; i < (files.Count); i++)
                    {
                        HttpPostedFileBase file = files[i];

                        if (file != null && file.ContentLength > 0)
                        {
                            string fileName = Path.GetFileName(file.FileName);
                            string ext = Path.GetExtension(fileName);
                            string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);

                            string uploadPath = Server.MapPath(folders[i]);
                            string finalPath = Path.Combine(uploadPath, nameWithoutExt + ext);

                            file.SaveAs(finalPath);

                            fileNames[i] = nameWithoutExt + ext;

                        }
                    }

                    if (fileNames[1] != null)
                    {
                        pod = fileNames[0];
                        ir = fileNames[1];

                    }
                    else
                    {
                        pod = fileNames[0];
                    }
                }

                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(data);

                DataTable table = new DataTable();
                table.Columns.Add("OrderDetailsId", typeof(int));
                table.Columns.Add("OrderDeliveryId", typeof(int));
                table.Columns.Add("ProductId", typeof(int));
                table.Columns.Add("ConsigneeName", typeof(string));
                table.Columns.Add("ConsigneeContactNo", typeof(string));
                table.Columns.Add("consigneeAddress", typeof(string));
                table.Columns.Add("DeliveredQty", typeof(double));
                table.Columns.Add("DeliveredTo", typeof(string));
                table.Columns.Add("DeliveredDate", typeof(string));
                table.Columns.Add("Document", typeof(string));
                table.Columns.Add("Document2", typeof(string));

                foreach (var Item in jsonData)
                {
                    DataRow dr = table.NewRow();
                    dr["OrderDetailsId"] = Item.OrderDetailsId;
                    dr["OrderDeliveryId"] = Item.OrderDeliveryId;
                    dr["ProductId"] = Item.ProductId;
                    dr["ConsigneeName"] = Item.ConsigneeName;
                    dr["ConsigneeContactNo"] = Item.ConsigneeContactNo;
                    dr["consigneeAddress"] = Item.consigneeAddress;
                    dr["DeliveredQty"] = Item.DeliveredQty;
                    dr["DeliveredTo"] = Item.DeliveredTo;
                    dr["DeliveredDate"] = Item.DeliveredDate;
                    dr["Document"] = Item.Document;
                    dr["Document2"] = Item.Document2;
                    table.Rows.Add(dr);
                }

                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "HardwareDeliveryMultiItems_AcceptUpdate";

                com.Parameters.AddWithValue("@PurchaseOrderNo", PurchaseOrderid);
                com.Parameters.AddWithValue("@ItemDetailsIdTxt", ItemDetailsIdTxt);
                com.Parameters.AddWithValue("@Document1", pod);
                com.Parameters.AddWithValue("@Document2", ir);
                com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());

                SqlParameter structuredParam = new SqlParameter("@TemptblHardwarSaleOrderDelivered", SqlDbType.Structured);
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
                mes = ex.Message;
            }
            return Json(mes);
        }


        public ActionResult GetConsigneeAddress()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@ItemDetailsId", Request.Form["OrderDetailId"].ToString());
            list.Add("@PurchaseOrderNo", Request.Form["PurchaseOrderNo"].ToString());
            list.Add("@consigneeAddress", Request.Form["consigneeAddress"].ToString());
            //dt = comfun.fillDataTable("HardwareDeliveryAddress_List", "", list);
            dt = comfun.fillDataTable("HardwareDeliveryAddress_List1", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }


        public ActionResult SubmitProductIRDelivery(String PurchaseOrderid, String data)
        {
            string mes = string.Empty;
            string ir = "";

            SortedList list = new SortedList();
            try
            {
                HttpFileCollectionBase files = Request.Files;
                HttpPostedFileBase firstFile = files[0];
                string firstFileName = Path.GetFileName(firstFile.FileName);
                if (files != null && (files.Count == 1) && Request.Files.AllKeys.Any() && firstFileName.Contains("IR"))
                {
                    if (firstFile != null && firstFile.ContentLength > 0)
                    {
                        string fileName = Path.GetFileName(firstFile.FileName);
                        string ext = Path.GetExtension(fileName);
                        string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);

                        string uploadPath = Server.MapPath("~/InstallationReport/");
                        string finalPath = Path.Combine(uploadPath, nameWithoutExt + ext);

                        firstFile.SaveAs(finalPath);

                        ir = nameWithoutExt + ext;

                    }
                }

                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(data);

                DataTable table = new DataTable();
                table.Columns.Add("OrderDeliveryId", typeof(int));
                table.Columns.Add("OrderDetailsId", typeof(int));


                foreach (var Item in jsonData)
                {
                    DataRow dr = table.NewRow();
                    dr["OrderDeliveryId"] = Item.OrderDeliveryId;
                    dr["OrderDetailsId"] = Item.OrderDetailsId;

                    table.Rows.Add(dr);
                }

                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "HardwareDeliveryMultiItemsIrUploaded_AcceptUpdate";

                com.Parameters.AddWithValue("@PurchaseOrderNo", PurchaseOrderid);
                com.Parameters.AddWithValue("@Document2", ir);
                com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());

                SqlParameter structuredParam = new SqlParameter("@TempTblIRUpload", SqlDbType.Structured);
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
                mes = ex.Message;
            }
            return Json(mes);
        }



        //public ActionResult DeleteConsigneeAddress()
        //{
        //    string Mes = "";
        //    try
        //    {
        //        DataTable dt = new DataTable();

        //        SortedList list = new SortedList();

        //        list.Add("@OrderDeliveryId", Request.Form["OrderDeliveryId"].ToString());
        //        list.Add("@ItemDetailsId", Request.Form["ItemDetailsId"].ToString());
        //        Mes = comfun.executeNonQueryWMessage("HardwareDeliveryAddress_Delete", "", list).ToString();
        //    }
        //    catch (Exception ex) { Mes = ex.Message; }
        //    return Json(Mes);


        //}
        public ActionResult DeleteConsigneeAddress()
        {
            string Mes = "";
            try
            {
                DataTable dt = new DataTable();

                SortedList list = new SortedList();

                list.Add("@OrderDeliveryId", Request.Form["OrderDeliveryId"].ToString());
                list.Add("@ItemDetailsId", Request.Form["ItemDetailsId"].ToString());
                //list.Add("@OrderDetailsId", Request.Form["ItemDetailsId"].ToString());
                //Mes = comfun.executeNonQueryWMessage("HardwareDeliveryAddress_Delete", "", list).ToString();
                Mes = comfun.executeNonQueryWMessage("HardwareDeliveryAddress_Delete1", "", list).ToString();
            }
            catch (Exception ex) { Mes = ex.Message; }
            return Json(Mes);


        }

        public ActionResult GetOrderDetail1()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@PurchaseOrderNoId", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("HardwarePurchaseOrderDetails_List2", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public JsonResult SubmitPurchaseBill1(string RequestDetails, string SaleOrderNo, string InvNo, string AgencyBillingId, string PurchaseOrderNo, string InvDate, string TotalAmount,
           string Narration, string IgstAmount, string CgstAmount, string SgstAmount, string GrandTotal)
        {
            string mes = string.Empty;
            try
            {
                var PurchaseInvoice_Doc = "";
                var _comPath = "";
                var filePathA = "";
                //var HardwareDelivery_Doc = "";
                //var filePathA2 = "";
                HttpFileCollectionBase files = Request.Files;

                if (System.Web.HttpContext.Current.Request.Files.AllKeys.Any())
                {

                    for (int i = 0; i < 1; i++)
                    {
                        HttpPostedFileBase file = files[i];


                        if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                        {
                            string[] testfiles = file.FileName.Split(new char[] { '\\' });
                            PurchaseInvoice_Doc = testfiles[testfiles.Length - 1];
                        }
                        else
                        {
                            PurchaseInvoice_Doc = file.FileName;
                        }
                        var _ext = Path.GetExtension(PurchaseInvoice_Doc);
                        PurchaseInvoice_Doc = Path.GetFileNameWithoutExtension(PurchaseInvoice_Doc);
                        string filePath = Path.Combine(Server.MapPath("/PurchaseInvoice/") + PurchaseInvoice_Doc + _ext);
                        PurchaseInvoice_Doc = PurchaseInvoice_Doc + _ext;
                        filePathA = filePath;
                        file.SaveAs(filePathA);

                    }

                }
                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(RequestDetails);

                DataTable table = new DataTable();
                table.Columns.Add("ProductId", typeof(int));
                table.Columns.Add("Quantity", typeof(double));
                table.Columns.Add("HSNCode", typeof(string));
                table.Columns.Add("Rate", typeof(double));
                table.Columns.Add("Gst", typeof(double));
                table.Columns.Add("AdminCharge", typeof(double));
                table.Columns.Add("Total", typeof(double));

                foreach (var Item in jsonData)
                {
                    DataRow dr = table.NewRow();
                    dr["ProductId"] = Item.ProductId;
                    dr["Quantity"] = Item.Quantity;
                    dr["HSNCode"] = Item.HSNCode;
                    dr["Rate"] = Item.Rate;
                    dr["Gst"] = Item.GST;
                    dr["AdminCharge"] = Item.AdminCharge;
                    dr["Total"] = Item.Total;

                    table.Rows.Add(dr);
                }

                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "HardwarePurchaseInv_AcceptUpdate";
                SqlParameter parameter = new SqlParameter();

                com.Parameters.AddWithValue("@PInvId", 0);
                com.Parameters.AddWithValue("@InvNo", InvNo);

                com.Parameters.AddWithValue("@SaleOrderNo", SaleOrderNo);
                com.Parameters.AddWithValue("@AgencyBillingId", AgencyBillingId);
                com.Parameters.AddWithValue("PurchaseOrderNo", PurchaseOrderNo);
                com.Parameters.AddWithValue("InvDate", InvDate);
                com.Parameters.AddWithValue("Narration", Narration);

                com.Parameters.AddWithValue("TotalAmount", TotalAmount);
                com.Parameters.AddWithValue("IgstAmount", IgstAmount);
                com.Parameters.AddWithValue("CgstAmount", CgstAmount);
                com.Parameters.AddWithValue("SgstAmount", SgstAmount);
                com.Parameters.AddWithValue("GrandTotal", GrandTotal);

                com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());
                com.Parameters.AddWithValue("@InvAttachement", PurchaseInvoice_Doc);

                SqlParameter parameter1 = new SqlParameter();
                parameter1.ParameterName = "@TempHardwarePurchaseInv2";
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

        public ActionResult SubmitVerificationDetailPODIR()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@OrderDeliveryId", Request.Form["orderdeliveryid"].ToString());
                list.Add("@OrderDetailsId", Request.Form["OrderDetailsId"].ToString());
                list.Add("@IsPODVerified", Request.Form["IsPODVerified"].ToString());
                list.Add("@IsIRVerified", Request.Form["IsIRVerified"].ToString());
                list.Add("@VerifedBy", Session["EmpId"].ToString());

                // mes = comfun.executeNonQueryWMessage("HardwarePODIRVerificationSingle_AcceptUpdate", "", list).ToString();
                mes = comfun.executeNonQueryWMessage("HardwarePODIRVerificationSingle_AcceptUpdate1", "", list).ToString();
            }
            catch (Exception ex)

            {
                mes = ex.Message;
            }
            return Json(mes);
        }

        public ActionResult GetSaleOrderDetail()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            list.Add("@PurchaseOrderNoId", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("HardwareSaleOrderDetails_List2", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public JsonResult SubmitSaleBill(string SaleData, string SaleOrderId, string SaleInvoiceNo, string PurchaseInvoiceNo, string InvoiceNo,
                 string AgencyBillingId, string PurchaseOrderNo, string InvDate, string TotalAmount,
                      string Igst, string cgst, string sgst, string AdminCharge, string Gtotal, string AdditionalRoundOff)
        {
            string mes = string.Empty;
            try
            {

                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(SaleData);

                DataTable table = new DataTable();
                table.Columns.Add("ProductId", typeof(int));
                table.Columns.Add("Quantity", typeof(double));
                table.Columns.Add("HSNCode", typeof(string));
                table.Columns.Add("Rate", typeof(double));
                table.Columns.Add("Gst", typeof(double));
                table.Columns.Add("AdminCharge", typeof(double));
                table.Columns.Add("Total", typeof(double));

                foreach (var Item in jsonData)
                {
                    DataRow dr = table.NewRow();
                    dr["ProductId"] = Item.ProductId;
                    dr["Quantity"] = Item.Quantity;
                    dr["HSNCode"] = Item.HSNCode;
                    dr["Rate"] = Item.Rate;
                    dr["Gst"] = Item.GST;
                    dr["AdminCharge"] = Item.AdminCharge;
                    dr["Total"] = Item.Total;

                    table.Rows.Add(dr);
                }

                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                //com.CommandText = "HardwareSasleInv_AcceptUpdate";
                com.CommandText = "HardwareSasleInv_AcceptUpdate1";

                com.Parameters.AddWithValue("@PurchaseOrderNo", PurchaseOrderNo);
                com.Parameters.AddWithValue("@SInvId", SaleInvoiceNo);
                com.Parameters.AddWithValue("@PInvId", PurchaseInvoiceNo);
                com.Parameters.AddWithValue("@InvNo", InvoiceNo);
                com.Parameters.AddWithValue("@AgencyBillingId", AgencyBillingId);
                com.Parameters.AddWithValue("@SaleOrderId", SaleOrderId);
                com.Parameters.AddWithValue("@InvDate", InvDate);
                com.Parameters.AddWithValue("@TotalAmount", TotalAmount);
                com.Parameters.AddWithValue("@IgstAmount", Igst);
                com.Parameters.AddWithValue("@CgstAmount", cgst);
                com.Parameters.AddWithValue("@SgstAmount", sgst);
                com.Parameters.AddWithValue("@AdminCharge", AdminCharge);
                com.Parameters.AddWithValue("@GrandTotal", Gtotal);
                com.Parameters.AddWithValue("@AdditionalRoundOff", AdditionalRoundOff);
                com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());


                SqlParameter structuredParam = new SqlParameter("@TempHardwarePurchaseInv2", SqlDbType.Structured);
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



        #endregion

        #region AdvancePay Report

        public ActionResult HardwarePaymentReport()
        {
            return View();
        }
        public ActionResult _HardwarePaymentReportList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            //list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            dt = comfun.fillDataTable("HardwarePaymentReportList ", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);



        }
        #endregion

        #region Deptt AdvancePay Report

        public ActionResult DepttAdvancePaymentReport()
        {
            return View();
        }
        public ActionResult _DepttAdvancePaymentReportList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            //list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
            list.Add("@DeptId", Request.Form["DeptId"].ToString());
            dt = comfun.fillDataTable("DepttAdvancePaymentReport ", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);



        }
        #endregion


        #region sale Invoice Cancel
        public ActionResult _CancellSaleInvSubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();

                list.Add("@SInvId", Request.Form["SInvId"].ToString());
                //list.Add("@IsCancel", Request.Form["IsCancel"].ToString());
                list.Add("@Remarks", Request.Form["Remarks"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                mes = comfun.executeNonQueryWMessage("stpHardwareSaleInv_Cancle ", "", list).ToString();
            }
            catch (Exception ex)

            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        #endregion
       

        #region Purchase Invoice Delete
        public ActionResult _PurchaseInvoiceDelete()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();

                list.Add("@PInvId", Request.Form["PInvId"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                list.Add("@RoleId", Session["RoleId"].ToString());
                mes = comfun.executeNonQueryWMessage("HardwarePurchaseInv_Delete", "", list).ToString();
            }
            catch (Exception ex)

            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        #endregion


        #region Advance Partial Payment
        public ActionResult _PartialPayList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@ReceiptId", Request.Form["ReceiptId"].ToString());
            list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
            dt = comfun.fillDataTable("HardwareDeptPaymentTransaction_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _GetInvoiceDetail()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
            list.Add("@PurchaseOrderId", Request.Form["PurchaseOrderId1"].ToString());
            dt = comfun.fillDataTable("HardwareDeptBalance_Get", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult PaymentMode()
        {
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            var json = "";
            DataTable dt1 = new DataTable();
            try
            {
                dt1 = comfun.fillDataTable("HpsedcPaymentmodeDdl_List", "", list);
                json = JsonConvert.SerializeObject(dt1);
            }
            catch (Exception ex) { }

            return Json(json);

        }

        public ActionResult BankList()
        {
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            var json = "";
            DataTable dt1 = new DataTable();
            try
            {
                dt1 = comfun.fillDataTable("HpsedcBankDdl_List", "", list);
                json = JsonConvert.SerializeObject(dt1);
            }
            catch (Exception ex) { }

            return Json(json);

        }


        public ActionResult _PaymentAccept()
        {
            //DataTable dt = new DataTable();
            SortedList list = new SortedList();
            string msg = "";
            try
            {
                list.Add("@ReceiptId", Request.Form["ReceiptId"].ToString());
                list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
                list.Add("@PurchaseOrderId", '0');
                list.Add("@TransactionId", Request.Form["TransactionId"].ToString());
                list.Add("@ModeOfPayment", Request.Form["ModeOfPayment"].ToString());
                list.Add("@BankNameId", Request.Form["BankNameId"].ToString());
                list.Add("@Narration", Request.Form["Narration"].ToString());
                list.Add("@ReceivedDate", Request.Form["ReceivedDate"].ToString());
                list.Add("@ReceivedAmt", Request.Form["ReceivedAmt"].ToString());
                list.Add("@GSTTds2", Request.Form["GSTTds2"].ToString());
                list.Add("@Tds2", Request.Form["Tds2"].ToString());
                list.Add("@BalanceAmt", Request.Form["BalanceAmt"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());

                msg = comfun.executeNonQueryWMessage("HardwareDeptPayment_AcceptUpdate", "", list).ToString();

            }

            catch (Exception ex) { msg = ex.Message; }

            return Json(msg);

        }
        public ActionResult GetConsigneeAddressOptimized()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@ItemDetailsId", Request.Form["OrderDetailId"].ToString());
            list.Add("@PurchaseOrderNo", Request.Form["PurchaseOrderNo"].ToString());
            list.Add("@consigneeAddress", Request.Form["consigneeAddress"].ToString());
            list.Add("@PageNumber", Request.Form["PageNumber"].ToString());
            list.Add("@PageSize", Request.Form["PageSize"].ToString());
            list.Add("@SearchTerm", Request.Form["SearchTerm"].ToString());
            //dt = comfun.fillDataTable("HardwareDeliveryAddress_List", "", list);
            dt = comfun.fillDataTable("HardwareDeliveryAddress_List1_optimize", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        #endregion
        #region Purchase Invoice Cancel
        public ActionResult _PurchaseInvCancelSubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();

                list.Add("@PInvId", Request.Form["PInvId"].ToString());
                list.Add("@IsRejected", Request.Form["IsRejected"].ToString());
                list.Add("@RejectRemarks", Request.Form["RejectRemarks"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                list.Add("@RoleId", Session["RoleId"].ToString());
                mes = comfun.executeNonQueryWMessage("HardwarePurchaseInv_Reject ", "", list).ToString();
            }
            catch (Exception ex)

            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        #endregion
        #region Remarks

        public ActionResult SubmitRemarks()
        {
            string msg = "";
            string remarksfile = "";
            string remarksfilePath = "";
            SortedList list = new SortedList();

            try
            {
                HttpFileCollectionBase files = Request.Files;

                if (System.Web.HttpContext.Current.Request.Files.AllKeys.Any())
                {

                    HttpPostedFileBase file = files[0];
                    if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                    {
                        string[] testfiles = file.FileName.Split(new char[] { '\\' });
                        remarksfile = testfiles[testfiles.Length - 1];
                    }
                    else
                    {
                        remarksfile = file.FileName;
                    }
                    var _ext = Path.GetExtension(remarksfile);
                    remarksfile = Path.GetFileNameWithoutExtension(remarksfile);
                    string filePath = Path.Combine(Server.MapPath("/SaleOrderRemarksDoc/") + remarksfile + _ext);
                    remarksfile = remarksfile + _ext;
                    remarksfilePath = filePath;
                    file.SaveAs(remarksfilePath);
                }
                list.Add("RemarkId", Request.Form["RemarkId"]);
                list.Add("SaleOrderId", Request.Form["SaleOrderId"]);
                list.Add("Remarks", Request.Form["Remarks"]);
                list.Add("RemarksBy", Session["EmpId"]);
                list.Add("Attachment", remarksfile);
                msg = comfun.executeNonQueryWMessage("HardwareRemarks_AcceptUpdate", "", list).ToString();
            }
            catch (Exception e)
            {
                msg = e.Message;
            }
            return Json(msg);
        }


        public ActionResult RemarksList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@RemarkId", Request.Form["RemarkId"].ToString());
            list.Add("@SaleOrderId", Request.Form["SaleOrderId"].ToString());
            dt = comfun.fillDataTable("HardwareRemarks_List", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }


        #endregion
    }
}