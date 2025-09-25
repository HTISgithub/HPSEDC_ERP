using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Payroll.portal.Models;
using Newtonsoft.Json;
using System.Collections;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Globalization;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using System.Web.Helpers;
using System.Web.UI.WebControls.WebParts;
using Microsoft.Web.Administration;
using System.Configuration;
using System.Web.Script.Serialization;
using System.Text;
using System.Net;
using DataTable = System.Data.DataTable;
using static Payroll.portal.Controllers.AdminController;

namespace Payroll.portal.Controllers
{
    [Authorize]
    public class PayrollController : Controller
    {
        public const string _roles = payrollFunctions._rolesGlobal;
        
        
        commonFunctions comfun = new commonFunctions();
        payrollFunctions payfun = new payrollFunctions();

        // GET: Payroll_EmployeeReimbursementJsonList
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult EmployeeReimbursement()
        {
            comfun.saveformname("EmployeeReimbursement", "/hr/EmployeeReimbursement", "Employee Reimbursement", "Leave Approval", "Y", "Employee Reimbursement", "Employee", "EmployeeReimbursement");

            if (payfun.sessionRecreate() == "expires")
            {
                return RedirectToAction("Login", "account");
            }
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            SortedList list = new SortedList();
            list.Add("@EmpId", Session["EmpId"].ToString());
            list.Add("@SessionId", Session["SessionId"].ToString());
            DataTable dt = comfun.fillDataTable("stpViksatReimbursementMonthByEmp", "", list);
            return View(dt);

        }

        public ActionResult _Employeelist()
        {
            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("stpEmployeeddl", "", null);
            //return PartialView("_Employeelist", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _ReimbursementList()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return PartialView("_sessionExpired");
            }
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            string EmployeeId = "0";
            if (Request.Form["EmpId"].ToString() == "0")
            {
                EmployeeId = Session["EmpID"].ToString();
            }
            else
            {
                EmployeeId = Request.Form["EmpId"].ToString();
            }
            list.Add("@EmpId", EmployeeId);
            list.Add("@fdDate", Convert.ToDateTime(Request.Form["Month"]).ToString("dd-MMM-yyyy"));
            list.Add("@SessionId", Session["SessionId"].ToString());
            dt = comfun.fillDataTable("stpViksatReimbursementByEmployee", "", list);
            ViewData["MonthYear"] = Request.Form["Month"].ToString();
            if (dt.Rows.Count == 0)
            {


                ViewData["Employee"] = Session["EmpName"].ToString();
            }
            else
            {
                ViewData["Employee"] = dt.Rows[0]["EmployeeName"].ToString();
            }
            ViewData["Department"] = dt.Rows[0]["DepttName"].ToString();
            ViewData["EmpCode"] = dt.Rows[0]["EmpCode"].ToString();
            ViewData["Designation"] = dt.Rows[0]["DesignationName"].ToString();

            return PartialView("_ReimbursementList", dt);
        }
        public ActionResult _ReimbursementBillSubmitView()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return PartialView("_sessionExpired");
            }
            string EmployeeId = "0";
            if (Request.Form["EmpId"].ToString() == "0")
            {
                EmployeeId = Session["EmpID"].ToString();
            }
            else
            {
                EmployeeId = Request.Form["EmpId"].ToString();
            }
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@EmpId", EmployeeId);
            list.Add("@fdDate", Convert.ToDateTime(Request.Form["Month"]).ToString("dd-MMM-yyyy"));
            list.Add("@SessionId", Session["SessionId"].ToString());
            dt = comfun.fillDataTable("stpViksatReimbursementByEmployee", "", list);
            ViewData["MonthYear"] = Request.Form["Month"].ToString();
            return PartialView("_ReimbursementBillSubmit", dt);
        }
        public JsonResult _ReimbursementAddSubmit(string ReimbursementData, string Month)
        {
            string mes = string.Empty;
            string MonthId = Convert.ToDateTime(Month).ToString("MM");
            string MonthyearId = Convert.ToDateTime(Month).ToString("Myyyy");

            try
            {
                DataTable table = (DataTable)JsonConvert.DeserializeObject(ReimbursementData, (typeof(DataTable)));
                //dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject(ReimbursementData);
                //DataTable table = new DataTable();
                //table.Columns.Add("Employee", typeof(int));
                //table.Columns.Add("ChargeID", typeof(int));
                //table.Columns.Add("Amount", typeof(decimal));
                //table.Columns.Add("Remarks", typeof(string));
                //foreach (var item in jsonData)
                //{
                //    DataRow dr = table.NewRow();
                //    dr["Employee"] = item.Employee;
                //    dr["ChargeID"] = item.ChargeID;
                //    dr["Amount"] = item.Amount;
                //    dr["Remarks"] = item.Remarks;
                //    table.Rows.Add(dr);
                //}

                SqlCommand com = new SqlCommand();
                connection conObj = new connection();

                com.Connection = conObj.con;

                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "stpViksatReimbursement_Save";
                com.Parameters.AddWithValue("@CreatedDate", comfun.dateISTstr());
                //com.Parameters.AddWithValue("@CreatedBy", 1);
                // com.Parameters.AddWithValue("@fiEmployeeId", Request.Form["EmployeeId"].ToString());

                com.Parameters.AddWithValue("@fiMonthId", MonthId);
                com.Parameters.AddWithValue("@fiMonthYearId", MonthyearId);
                com.Parameters.AddWithValue("@fiSessionId", Session["SessionId"].ToString());
                com.Parameters.AddWithValue("@fiBranchId", "0");
                com.Parameters.AddWithValue("@fdDate", Convert.ToDateTime(Month).ToString("dd/MMM/yyyy"));
                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@ReimbursementTable";

                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table;
                com.Parameters.Add(parameter);
                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);

                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();
                //mes = comfun.executeNonQueryWMessage("stpViksatReimbursement_Save","", ).ToString();
                com.ExecuteNonQuery();
                conObj.con.Close();
                mes = sp.Value.ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Payroll  _ReimbursementAddSubmit", "Error: " + ex.Message);
            }
            return Json(mes);



            //SortedList list = new SortedList();
            //list.Add("@fiCharges", Request.Form["ChargeID"].ToString());
            //list.Add("@fiSessionId", 1);
            //list.Add("@fiBranchId", 10);
            //list.Add("@ReimbursementAmount", Request.Form["Amount"].ToString());
            //list.Add("@fiEmployeeId", Request.Form["EmpId"].ToString());
            //list.Add("@fiMonthid",MonthId );
            //list.Add("@fiMonthYearid",MonthyearId);
            //list.Add("@CreatedBy", Request.Form["EmpId"].ToString());
            //list.Add("@CreatedDate", comf.dateISTstr());

        }
        [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult Loanissued()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return RedirectToAction("Login", "account");
            }

            comfun.saveformname("Loanissued", "/Payroll/Loanissued", "Loanissued", "Loanissued", "Y", "Loan Issue", "Loan Issue", "List");
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            DataTable dt = comfun.fillDataTable("stpViksatLoanTypeSelect", "", null);
            return View(dt);
        }
        public JsonResult _loanissuedSubmit()
        {
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
                mes = comfun.errorMessage("Payroll _loanissuedSubmit", ex.Message);
            }
            return Json(mes);
        }
        [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult LoanList()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return RedirectToAction("Login", "account");
            }
            SortedList list = new SortedList();
            // list.Add("@Employeeid", Session["EmpId"].ToString());
            list.Add("@CompanyId", Session["CompanyId"].ToString());

            DataTable dt = comfun.fillDataTable("stpVIKSATLoanIssuedGridWithoutid", "", list);
            return View(dt);
        }
        public ActionResult _LoanIssueDetail()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return PartialView("_sessionExpired");
            }
            SortedList list = new SortedList();
            list.Add("@fiLoanId", Request.Form["fiLoanID"].ToString());
            DataTable dt = comfun.fillDataTable("stpViksatLoandetailByID", "", list);
            return PartialView("_LoanDetail", dt);

        }
        [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult LoanEdit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return RedirectToAction("Login", "account");
            }
            string LoanId = "0";
            if (Request.QueryString["id"] == null)
            {
                return RedirectToAction("LoanList");
            }
            LoanId = comfun.decryptString(Request.QueryString["Id"].ToString());
            SortedList list = new SortedList();
            list.Add("@LoanId", LoanId);
            DataSet ds = comfun.fillDataSet("stpLoanSelectByLoanId", "", list);
            return View(ds);

        }
        public ActionResult _LoanEditSubmit()
        {
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
                mes = comfun.errorMessage("Payroll _LoanEditSubmit", ex.Message);
            }
            return Json(mes);
        }
        // public ActionResult _loanIssueGrid()
        // {


        //   }
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
                mes = comfun.errorMessage("Payroll _LoanDetail", "Error: " + ex.Message);
            }
            return Json(mes);
        }

        public ActionResult CtcFormula()
        {
            SortedList list = new SortedList();
            list.Add("@ChargeID", "0");
            DataTable dt = comfun.fillDataTable("stpEmployeeCtcChargesddl", "", list);
            return View(dt);
        }

        public ActionResult _CtcFormulaCheckboxList()
        {
            SortedList list = new SortedList();
            list.Add("@ChargeID", Request.Form["ChargeId"].ToString());
            DataTable dt = comfun.fillDataTable("stpEmployeeCtcChargesddl", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _ctcFormnulaList()
        {
            SortedList list = new SortedList();
            DataTable dt = comfun.fillDataTable("stpEmployeeCtcformula_list", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _ctcFormnulaListByID()
        {
            SortedList list = new SortedList();
            list.Add("@ID", Request.Form["FormulaId"].ToString());
            DataTable dt = comfun.fillDataTable("stpEmployeeCtcformulaSelectViaID", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _chargeListSave(string Charges)
        {
            string mes = string.Empty;
            string formulaName = "";
            bool isSuccess = false;
            try
            {
                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject(Charges);
                DataTable table = new DataTable();
                table.Columns.Add("Chargeid", typeof(int));
                table.Columns.Add("ChargePer", typeof(decimal));
                table.Columns.Add("OnChargeID", typeof(string));
                table.Columns.Add("IsGross", typeof(string));
                table.Columns.Add("Amount", typeof(decimal));
                table.Columns.Add("isFixed", typeof(string));


                foreach (var item in jsonData)
                {
                    DataRow dr = table.NewRow();
                    dr["Chargeid"] = item.Chargeid;
                    dr["ChargePer"] = item.ChargePer;
                    dr["OnChargeID"] = (item.OnChargesID == "null" ? "" : item.OnChargesID);
                    dr["IsGross"] = item.IsGross;
                    dr["Amount"] = item.Amount;
                    dr["isFixed"] = item.IsFixed;
                    formulaName = item.FormulaName;
                    table.Rows.Add(dr);
                }

                SqlCommand com = new SqlCommand();
                connection conObj = new connection();

                com.Connection = conObj.con;

                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "stpEmployeeCtcformula_Save";
                //com.Parameters.AddWithValue("@CreatedDate", comfun.dateISTstr());
                com.Parameters.AddWithValue("@formulaName", formulaName);


                com.Parameters.AddWithValue("@ID", 0);
                com.Parameters.AddWithValue("@CreatedBy", Request.Cookies["EmpId"].Value);
                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@CtcFormula";

                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table;
                com.Parameters.Add(parameter);

                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 500);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);

                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();
                com.ExecuteNonQuery();
                conObj.con.Close();

                mes = sp.Value.ToString();

                mes = "Record saved successfully";
                isSuccess = true;
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _chargeListSave", "Error: " + ex.Message);
                mes = "Error:" + ex.Message;
                isSuccess = false;
            }
            return Json(new { success = isSuccess, responseText = mes }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult employeeCtcDisplay()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return RedirectToAction("Login", "account");
            }
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            SortedList list = new SortedList();
            list.Add("@EmpId", Session["EmpId"].ToString());
            list.Add("@FdDate", comfun.dateISTstr());
            comfun.saveformname("employeeCtcDisplay", "/Payroll/employeeCtcDisplay", "Employee Ctc", "Employee Ctc Display", "Y", "Employee Ctc Display", "Employee Ctc Display", "List");
            DataTable dt = comfun.fillDataTable("stpViksatEmployeeCTC_show", "", list);

            return View(dt);
        }
        public ActionResult _employeeCtcDisplay()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return PartialView("_sessionExpired");
            }
            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());

            DataTable dt = comfun.fillDataTable("stpViksatEmployeeCTC_show", "", list);
            return PartialView("_employeeCtcDisplay", dt);
        }
        public ActionResult PayChargesApplied()
        {
            comfun.saveformname("PayChargesApplied", "/Payroll/PayChargesApplied", "Pay Charges Applied", "Pay Charges Applied", "Y", "Pay Charges Applied", "Pay Charges Applied", "List");

            if (payfun.sessionRecreate() == "expires")
            {
                return PartialView("_sessionExpired");
            }
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("stpEmployeedropdown", "", null);

            return View(dt);
        }
        public ActionResult EmployeeCTList()
        {

            comfun.saveformname("EmployeeCTC", "/Payroll/EmployeeCTList", "Emp Management/Employee CTC", "Employee CTC Main Form", "Y", "EmployeeCTC", "EmployeeCTC", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("EmployeeCTC");
            return View();
        }
        public JsonResult _EmployeeCTCJsonList()
        {

            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("stpCtclist", "", null);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult AccountledgerReport()
        {
            return View();
        }
        public ActionResult _Accountledgerlist()
        {
            SortedList list = new SortedList();
            list.Add("@fdFromDate", Request.Form["FromDate"].ToString());
            list.Add("@fdToDate", Request.Form["ToDate"].ToString());
            list.Add("@fiAccountID", Request.Form["AccountID"].ToString());

            list.Add("@fiSessionID", "1");

            DataTable dt = comfun.fillDataTable("stpLedgerAc", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public JsonResult _EmployeeCTCJsonListJ()
        {
            comfun.saveformname("_EmployeeCTCJsonListJ", "/Payroll/_EmployeeCTCJsonListJ", "Emp Management/Employee CTC", "Employee CTC", "N", "EmployeeCTC", "EmployeeCTC", "List", 4);
            SortedList list = new SortedList();
            list.Add("@emp", Request.Form["Emp"].ToString());
            DataTable dt = comfun.fillDataTable("stpCtclistJ", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1, JsonRequestBehavior.AllowGet);
        }

        public ActionResult _EmployeeSelect()
        {

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@EmployeeId", Request.Form["EmpId"].ToString());
            dt = comfun.fillDataTable("stpEmployeedropdown", "", list);

            return PartialView("_EmployeeSelect", dt);
        }

        //DataTable dt = comfun.fillDataTable("stpEmployeeCtcformula_list", "", null);

        public ActionResult _ctcFormnulaListP2()
        {
            SortedList list = new SortedList();
            list.Add("@FormulaID", Request.Form["FormulaID"].ToString());
            DataTable dt = comfun.fillDataTable("stpEmployeeCtcformula_listP2", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _ctcByFormula()
        {
            //SortedList list = new SortedList();
            //list.Add("@FormulaID", "1006");
            //list.Add("@Basic", "300");
            //list.Add("@SpecialID", Request.Form["ChargeId"].ToString());
            //list.Add("@TotalCTC", Request.Form["TotalCTC"].ToString());
            //list.Add("@ChargeAppliedDate", "1-1-1");
            //list.Add("@EmployeeId", "0");
            //DataTable dt = comfun.fillDataTable("stpEmployeeCTC_Calculation", "", list);
            ViewData["EmpId"] = Request.Form["EmpId"].ToString();
            return PartialView("_ctcByFormula");
        }

        public ActionResult _ctcByFormulaJson(string Extra)
        {
            SortedList list = new SortedList();
            list.Add("@FormulaID", Request.Form["FormulaID"].ToString());
            list.Add("@Basic", Request.Form["Basic"].ToString());
            list.Add("@SpecialID", Request.Form["ChargeId"].ToString());
            list.Add("@TotalCTC", Request.Form["TotalCTC"].ToString());
            list.Add("@Changablelist", Request.Form["Changablelist"].ToString());
            list.Add("@ChargeAppliedDate", Request.Form["ChargeAppliedDate"].ToString());
            list.Add("@EmployeeId", Request.Form["EmpId"].ToString());
            list.Add("@ExtraEmpPart", Extra);
            //list.Add("@AdjustmentId", Request.Form["AdjustmentId"].ToString());
            DataTable dt = comfun.fillDataTable("stpEmployeeCTC_Calculation", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1, JsonRequestBehavior.AllowGet);
        }

        public ActionResult _ChargeApplied()
        {
            SortedList list = new SortedList();
            DataTable dt = new DataTable();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            dt = comfun.fillDataTable("PayChargesAppliedSearch", "", list);
            ViewData["WEF"] = Request.Form["WEF"].ToString();
            return PartialView("_Charges", dt);
        }
        public JsonResult _PayChargesSubmit()
        {
            comfun.saveformname("_PayChargesSubmit", "/Payroll/_PayChargesSubmit", "Emp Management/Employee CTC", "Employee CTC", "N", "EmployeeCTC", "EmployeeCTC", "Add", 2);
            comfun.saveformname("_PayChargesSubmit", "/Payroll/_PayChargesSubmit", "Emp Management/Employee CTC", "Employee CTC", "N", "EmployeeCTC", "EmployeeCTC", "Edit", 3);

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
        public ActionResult ReimbursementApproval()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            comfun.saveformname("ReimbursementApproval", "/Payroll/ReimbursementApproval", "Reimbursement Approval", "Reimbursement Approval", "Y", "Reimbursement Approval", "Reimbursement Approval", "List");

            return View();
        }
        public ActionResult _ReimbursementApprovalList()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return PartialView("_sessionExpired");
            }
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@fdDate", Convert.ToDateTime(Request.Form["Month"]).ToString("dd-MMM-yyyy"));
            ///list.Add("@CostCenterId", Session["CostCenterId"].ToString());
            dt = comfun.fillDataTable("stpViksatReimbursementApprovalByemployee", "", list);
            ViewData["MonthYear"] = Request.Form["Month"].ToString();
            return PartialView("_ReimbursementApprovalList", dt);
        }
        public JsonResult _ReimbursementApprovalSubmitView()
        {
            string mes = "Reimburesement Approved Succesfully!!";
            SortedList list = new SortedList();
            dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject(Request.Form["ChkApproval"].ToString());

            //string[] chkboxval = jsonData;
            foreach (var item in jsonData)
            {
                list.Clear();
                list.Add("@ReimburesementId", item);
                comfun.executeNonQuery("stpViksatReimbursementApproval", "", list);
            }


            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return PartialView("_sessionExpired");
            //}
            //DataTable dt = new DataTable();
            //SortedList list = new SortedList();
            //list.Add("@EmpId", Request.Form["EmpId"].ToString());
            //list.Add("@fdDate", Convert.chkboxval(Request.Form["Month"]).ToString("dd-MMM-yyyy"));

            ////dt = comfun.fillDataTable("stpViksatReimbursementApproval", "", list);

            //return Json(mes);
            return Json(mes);
        }
        public ActionResult freeze()
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}
            //if (Request.QueryString["key"] != null)
            //{
            //    Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            //}
            //  comfun.saveformname("freeze", "/Payroll/freeze", "Freeze", "Freeze", "Y", "Freeze", "Freeze", "List");
            return View("freezeInput");
        }
        public ActionResult _freezeFilters()
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}
            //if (Request.QueryString["key"] != null)
            //{
            //    Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            //}
            //ViewData["CostCenterName"] = Session["CostCenterName"].ToString();
            //ViewData["CostCenterId"] = Session["CostCenterId"].ToString();
            return PartialView("_freezeFilters");
        }
        public ActionResult _freezeFilterResult(FormCollection form)
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}
            string type = Request.Form["Type"].ToString();

            ViewData["ShowDays"] = "N";

            string proc = "";

            if (Request.Form["Method"].ToString() == "Attendance")
            {
                ViewData["ShowDays"] = "Y";
                if (type == "Freeze")
                {
                    proc = "stpEmployeeAttendance_MonthlyLock";
                }
                if (type == "UnFreeze")
                {
                    proc = "LockAttendance_UnFreeze";
                }
            }

            if (Request.Form["Method"].ToString() == "AttendanceForSalary")
            {
                ViewData["ShowDays"] = "Y";

                if (type == "Freeze")
                {
                    proc = "LockAttendanceForSalary_Freeze";
                }
                if (type == "UnFreeze")
                {
                    proc = "LockAttendanceForSalary_UnFreeze";
                }
            }

            if (Request.Form["Method"].ToString() == "OT")
            {

                if (type == "Freeze")
                {
                    proc = "LockOT_Freeze";
                }
                if (type == "UnFreeze")
                {
                    proc = "LockOT_UnFreeze";
                }
            }

            if (Request.Form["Method"].ToString() == "Salary")
            {
                ViewData["ShowDays"] = "Y";
                if (type == "Freeze")
                {
                    proc = "LockSalary_freeze";
                }
                if (type == "UnFreeze")
                {
                    proc = "LockSalary_UnFreeze";
                }
            }

            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 15;
            string fromDate = Request.Form["StartDate"].ToString();
            int startYear = Convert.ToDateTime(fromDate).Year;
            int startMonth = Convert.ToDateTime(fromDate).Month;

            DateTime startDate = new DateTime(startYear, startMonth, 1);
            DateTime endDate = startDate.AddMonths(1).AddDays(-1);

            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }
            if (Request.Form["PageNo"] != null)
            {
                pagesize = Convert.ToDecimal(Request.Form["PageSize"]);
            }
            SortedList list = new SortedList();
            //paging_get_downline1
            // list.Add("@MonthYearId", startDate.Month.ToString() + startDate.Year.ToString());
            //list.Add("@PageNo", pageno);
            //list.Add("@PageSize", pagesize);
            //***************Parameters*******************************************


            list.Add("@StartDate", startDate.ToString("dd/MMM/yyyy"));
            list.Add("@EndDate", endDate.ToString("dd/MMM/yyyy"));
            list.Add("@loginId", Convert.ToInt32(Session["EmpId"].ToString()));
            //list.Add("@Role", Enum.GetName(typeof(payrollFunctions.roleIds), Convert.ToInt32(Session["RoleId"].ToString())));
            //if (Request.Form["CostCenterId"].ToString() != "0" && Request.Form["CostCenterId"].ToString() != "")
            //{
            //    list.Add("@CostCenterId", Request.Form["CostCenterId"].ToString());
            //}
            //else
            //{
            //    list.Add("@CostCenterId", Session["CostCenterId"].ToString());
            //}
            list = payfun.globalParameterList(form, list);
            //DataTable dt = comfun.fillDataTable("Attendance_SalaryForAttLock", "", list);
            list.Remove("@Method");
            list.Remove("@Type");
            DataTable dt = comfun.fillDataTable(proc, "", list);

            /*
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "Attendance", "_salaryAttMonthly", "_salaryAttMonthly");

            HtmlString htm = new HtmlString(paging);
             ViewData["paging"] = htm;
            */
            ViewData["paging"] = "";
            ViewData["StartYear"] = startYear;
            ViewData["StartMonth"] = startMonth;
            ViewData["EmpId"] = Request.Form["EmpId"].ToString();

            //DateTime startDate = new DateTime(startYear, startMonth, 1);
            //DateTime endDate = startDate.AddMonths(1).AddDays(-1);

            return PartialView("_freezeFilterResult", dt);
        }
        public JsonResult _freezeInputSubmit(FormCollection form)
        {
            string mes = "";
            try
            {
                //if (payfun.sessionRecreate() == "expires")
                //{
                //    return Json("Session expires");
                //}
                SortedList list = new SortedList();
                decimal total_records = 0;
                decimal pageno = 1;
                decimal pagesize = 15;
                string fromDate = Request.Form["FromDate"].ToString();
                int startYear = Convert.ToDateTime(fromDate).Year;
                int startMonth = Convert.ToDateTime(fromDate).Month;

                DateTime startDate = new DateTime(startYear, startMonth, 1);
                DateTime endDate = startDate.AddMonths(1).AddDays(-1);

                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject(Request.Form["FreezeEmpIds"].ToString());

                DataTable dt = new DataTable();
                dt.Columns.Add("EmpId", typeof(int));
                dt.Columns.Add("AttendanceTypeId", typeof(int));
                DataRow dr;

                foreach (var item in jsonData)
                {
                    dr = dt.NewRow();
                    dr["EmpId"] = item.EmpId;
                    dr["AttendanceTypeId"] = 0;
                    dt.Rows.Add(dr);
                }

                list = payfun.globalParameterList(form, list);
                SqlCommand com = new SqlCommand();
                connection conObj = new connection();

                com.Connection = conObj.con;

                com.CommandType = CommandType.StoredProcedure;
                if (Request.Form["Method"].ToString() == "Attendance")
                {
                    com.CommandText = "LockAttendance_FreezeSubmit";
                }

                if (Request.Form["Method"].ToString() == "Salary")
                {
                    com.CommandText = "LockSalary_FreezeSubmit";
                }

                if (Request.Form["Method"].ToString() == "OT")
                {
                    com.CommandText = "Attendance_SalaryForTimeLockSubmit";
                }

                if (Request.Form["Method"].ToString() == "AttendanceForSalary")
                {
                    com.CommandText = "LockAttendanceForSalary_FreezeSubmit";
                }

                foreach (string li in list.Keys)
                {
                    if (!li.Contains("FreezeEmpIds") && !li.Contains("FromDate") && !li.Contains("Method"))
                    {
                        com.Parameters.AddWithValue(li, list[li].ToString().Trim());
                    }
                }
                com.Parameters.AddWithValue("@MonthYearId", startMonth.ToString() + startYear.ToString());
                com.Parameters.AddWithValue("@LockDate", endDate.ToString("dd-MMM-yyyy"));
                com.Parameters.AddWithValue("@CreatedDate", comfun.dateISTstr());
                com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());

                //com.Parameters.AddWithValue("@FromDate", startDate);
                //com.Parameters.AddWithValue("@ToDate", endDate);

                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@AttendanceTable";

                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = dt;
                com.Parameters.Add(parameter);


                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();
                com.ExecuteNonQuery();
                conObj.con.Close();
                mes = "Lock has been applied successfully";
            }
            catch (Exception ex)
            {
                mes = "Error:" + ex.Message;
            }
            return Json(mes);
        }


        public ActionResult Salaryfreeze()
        {
            SortedList list = new SortedList();
            list.Add("@PageNo", 1);
            list.Add("@PageSize", 1000);
            DataTable dt = comfun.fillDataTable("Department_Select", "", list);
            return View("SalaryfreezeInput", dt);
        }
        public JsonResult _FreezeSalaryList()
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
            list.Add("@BranchId", Request.Form["BranchId"].ToString());
            DataTable dt = comfun.fillDataTable("stpEmployeeSalary_MonthlyLockList", "", list);
            //return PartialView("_HelpdeskGroupMappingList", dt);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }


        public ActionResult _SalaryfreezeFilters()
        {
            //if (payfun.sessionRecreate() == "expires")
            //{
            //    return Json("Session expires");
            //}
            //if (Request.QueryString["key"] != null)
            //{
            //    Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            //}
            //ViewData["CostCenterName"] = Session["CostCenterName"].ToString();
            //ViewData["CostCenterId"] = Session["CostCenterId"].ToString();
            return PartialView("_SalaryfreezeFilters");
        }
        public ActionResult _SalryfreezeFilterResult(FormCollection form)
        {
            SortedList list = new SortedList();
            string fromDate = Request.Form["FromDate"].ToString();
            int startYear = Convert.ToDateTime(fromDate).Year;
            int startMonth = Convert.ToDateTime(fromDate).Month;
            string Type = Request.Form["Type"].ToString();
            DateTime startDate = new DateTime(startYear, startMonth, 1);
            DateTime endDate = startDate.AddMonths(1).AddDays(-1);
            list.Add("@StartDate", startDate.ToString("dd/MMM/yyyy"));
            list.Add("@EndDate", endDate.ToString("dd/MMM/yyyy"));
            list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
            list.Add("@BranchId", Request.Form["BranchId"].ToString());
            list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@CategoryId", Request.Form["CategoryId"].ToString());
            list.Add("@CircleId", Request.Form["CircleId"].ToString());
            list.Add("@MainCategoryId", Request.Form["MainCategoryId"].ToString());
            list.Add("@LoginID", Session["EmpId"].ToString());
            DataTable dt = new DataTable();
            if (Type == "Freeze")
            {
                dt = comfun.fillDataTable("LockSalary_freeze", "", list);

            }
            if (Type == "UnFreeze")
            {
                dt = comfun.fillDataTable("LockSalary_UnFreeze", "", list);

            }


            var json = JsonConvert.SerializeObject(dt);
            return Json(json);




        }
        public JsonResult _FreezeSalarySaveUpdateSubmit()
        {
            // comfun.saveformname("_HelpdeskMappingSaveUpdateSubmit", "/admin/_HelpdeskMappingSaveUpdateSubmit", "Master/Help Desk/Mapping&nbsp;&nbsp;", "HelpDesk Mapping Add", "N", "HelpDeskMapping", "HelpDeskMapping", "Add", 2);
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
                mes = comfun.executeNonQueryWMessage("EmployeeSalary_FreezeSubmit", "", list).ToString();


            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }

        public JsonResult _UnFreezesalarySaveUpdateSubmit()
        {
            // comfun.saveformname("_HelpdeskMappingSaveUpdateSubmit", "/admin/_HelpdeskMappingSaveUpdateSubmit", "Master/Help Desk/Mapping&nbsp;&nbsp;", "HelpDesk Mapping Add", "N", "HelpDeskMapping", "HelpDeskMapping", "Add", 2);
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
                mes = comfun.executeNonQueryWMessage("EmployeeSalary_FreezeSubmit", "", list).ToString();


            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }

        public ActionResult GenerateSalary()
        {
            SortedList list = new SortedList();

            return View();
        }
        public ActionResult _GenrateSalaryList()
        {
            SortedList list = new SortedList();

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
            DataTable dt = comfun.fillDataTable("EmployeeSalary_generate", "", list);




            return PartialView("_GenerateSalary", dt);
        }


        public ActionResult _EmpSalaryGenerate()
        {
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
                // list.Add("@Type", Request.Form["Type"].ToString());
                list.Add("@EmpId", Request.Form["GenerateSalaryId"].ToString());
                mes = comfun.executeNonQueryWMessage("EmployeeGenerateSalary_Submit", "", list).ToString();


            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        public ActionResult _EmpDeleteSalaryGenerate()
        {
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
                // list.Add("@Type", Request.Form["Type"].ToString());
                list.Add("@EmpId", Request.Form["GenerateSalaryId"].ToString());
                mes = comfun.executeNonQueryWMessage("EmployeeDeleteSalary_Submit", "", list).ToString();


            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }






        public ActionResult SalaryCalculation()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            comfun.saveformname("SalaryCalculation", "/Payroll/SalaryCalculation", "SalaryCalculation", "SalaryCalculation", "Y", "SalaryCalculation", "SalaryCalculation", "List");
            return View("SalaryCalculation");
        }

        public ActionResult _Salaryfilters()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            ViewData["CostCenterName"] = Session["CostCenterName"].ToString();
            ViewData["CostCenterId"] = Session["CostCenterId"].ToString();
            return PartialView("_Salaryfilters");
        }

        public ActionResult _SalaryFilterResult(FormCollection form)
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 15;
            string fromDate = Request.Form["FromDate"].ToString();
            int startYear = Convert.ToDateTime(fromDate).Year;
            int startMonth = Convert.ToDateTime(fromDate).Month;

            DateTime startDate = new DateTime(startYear, startMonth, 1);
            DateTime endDate = startDate.AddMonths(1).AddDays(-1);


            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }
            if (Request.Form["PageNo"] != null)
            {
                pagesize = Convert.ToDecimal(Request.Form["PageSize"]);
            }
            SortedList list = new SortedList();
            //paging_get_downline1
            list.Add("@MonthYearId", startDate.Month.ToString() + startDate.Year.ToString());
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            //***************Parameters*******************************************
            list.Add("@FromDate", Convert.ToDateTime(startDate).ToString("dd-MMM-yyyy"));
            list.Add("@ToDate", Convert.ToDateTime(endDate).ToString("dd-MMM-yyyy"));
            list.Add("@loginId", Convert.ToInt32(Session["EmpId"].ToString()));
            list.Add("@Role", Enum.GetName(typeof(payrollFunctions.roleIds), Convert.ToInt32(Session["RoleId"].ToString())));
            if (Request.Form["CostCenterId"].ToString() != "0" && Request.Form["CostCenterId"].ToString() != "")
            {
                list.Add("@CostCenterId", Request.Form["CostCenterId"].ToString());
            }
            else
            {
                list.Add("@CostCenterId", Session["CostCenterId"].ToString());
            }
            list = payfun.globalParameterList(form, list);
            //DataTable dt = comfun.fillDataTable("Attendance_SalaryForAttLock", "", list);

            DataTable dt = comfun.fillDataTable("LockAttendanceForSalary_UnFreeze", "", list);


            ViewData["paging"] = "";
            ViewData["StartYear"] = startYear;
            ViewData["StartMonth"] = startMonth;
            ViewData["EmpId"] = Request.Form["EmpId"].ToString();



            return PartialView("_SalaryFilterResult", dt);
        }
        public JsonResult _SalaryInputSubmit(FormCollection form)
        {
            string mes = "";
            try
            {
                if (payfun.sessionRecreate() == "expires")
                {
                    return Json("Session expires");
                }
                SortedList list = new SortedList();
                decimal total_records = 0;
                decimal pageno = 1;
                decimal pagesize = 15;
                string fromDate = Request.Form["FromDate"].ToString();
                int startYear = Convert.ToDateTime(fromDate).Year;
                int startMonth = Convert.ToDateTime(fromDate).Month;

                DateTime startDate = new DateTime(startYear, startMonth, 1);
                DateTime endDate = startDate.AddMonths(1).AddDays(-1);

                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject(Request.Form["FreezeEmpIds"].ToString());

                DataTable dt = new DataTable();
                dt.Columns.Add("EmpId", typeof(int));
                dt.Columns.Add("AttendanceTypeId", typeof(int));
                DataRow dr;

                foreach (var item in jsonData)
                {
                    dr = dt.NewRow();
                    dr["EmpId"] = item.EmpId;
                    dr["AttendanceTypeId"] = 0;
                    dt.Rows.Add(dr);
                }
                list = payfun.globalParameterList(form, list);
                SqlCommand com = new SqlCommand();
                connection conObj = new connection();

                com.Connection = conObj.con;

                com.CommandType = CommandType.StoredProcedure;
                if (Request.Form["Method"].ToString() == "SalaryCalculation")
                {
                    com.CommandText = "CMRSalaryCalculation_";
                }


                foreach (string li in list.Keys)
                {
                    if (!li.Contains("FreezeEmpIds") && !li.Contains("FromDate") && !li.Contains("Method"))
                    {
                        com.Parameters.AddWithValue(li, list[li].ToString().Trim());
                    }
                }
                // com.Parameters.AddWithValue("@MonthYearId", startMonth.ToString() + startYear.ToString());
                // com.Parameters.AddWithValue("@LockDate", endDate.ToString("dd-MMM-yyyy"));
                //com.Parameters.AddWithValue("@CreatedDate", comfun.dateISTstr());
                // com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());
                com.Parameters.AddWithValue("@FromDate", startDate.ToString("dd-MMM-yyyy"));
                com.Parameters.AddWithValue("@ToDate", endDate.ToString("dd-MMM-yyyy"));

                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@AttendanceTable";

                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = dt;
                com.Parameters.Add(parameter);
                com.CommandTimeout = 0;

                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();
                com.ExecuteNonQuery();
                conObj.con.Close();
                mes = "Salary has been Calculate successfully";
            }
            catch (Exception ex)
            {
                mes = "Error:" + ex.Message;
            }
            return Json(mes);
        }
        public ActionResult OT()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            comfun.saveformname("OT", "/Payroll/OT", "OT", "OT", "Y", "OT", "OT", "List");
            return View("OTCalculation");
        }
        public ActionResult _OTfilters()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            ViewData["CostCenterName"] = Session["CostCenterName"].ToString();
            ViewData["CostCenterId"] = Session["CostCenterId"].ToString();
            return PartialView("_OTfilters");
        }
        public ActionResult _OTFilterResult(FormCollection form)
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 15;
            string fromDate = Request.Form["FromDate"].ToString();
            int startYear = Convert.ToDateTime(fromDate).Year;
            int startMonth = Convert.ToDateTime(fromDate).Month;

            DateTime startDate = new DateTime(startYear, startMonth, 1);
            DateTime endDate = startDate.AddMonths(1).AddDays(-1);
            string type = Request.Form["Type"].ToString();
            string Method = Request.Form["Method"].ToString();
            string proc = "";
            if (Request.Form["Method"].ToString() == "OTCal")
            {
                ViewData["Showhrs"] = "N";
                proc = "Attendance_SalaryOTForAttLock";

            }
            if (Request.Form["Method"].ToString() == "OTFinal")
            {
                ViewData["Showhrs"] = "Y";
                proc = "Attendance_SalaryOTForAttLocksalary";
            }


            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }
            if (Request.Form["PageNo"] != null)
            {
                pagesize = Convert.ToDecimal(Request.Form["PageSize"]);
            }
            SortedList list = new SortedList();
            //paging_get_downline1
            list.Add("@MonthYearId", startDate.Month.ToString() + startDate.Year.ToString());
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            //***************Parameters*******************************************
            list.Add("@FromDate", startDate.ToString("dd-MMM-yyyy"));
            list.Add("@ToDate", endDate.ToString("dd-MMM-yyyy"));
            list.Add("@loginId", Convert.ToInt32(Session["EmpId"].ToString()));
            list.Add("@Role", Enum.GetName(typeof(payrollFunctions.roleIds), Convert.ToInt32(Session["RoleId"].ToString())));
            if (Request.Form["CostCenterId"].ToString() != "0" && Request.Form["CostCenterId"].ToString() != "")
            {
                list.Add("@CostCenterId", Request.Form["CostCenterId"].ToString());
            }
            else
            {
                list.Add("@CostCenterId", Session["CostCenterId"].ToString());
            }
            list = payfun.globalParameterList(form, list);
            list.Remove("@Method");
            list.Remove("@Type");

            DataTable dt = comfun.fillDataTable(proc, "", list);
            /*
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "Attendance", "_salaryAttMonthly", "_salaryAttMonthly");

            HtmlString htm = new HtmlString(paging);
             ViewData["paging"] = htm;
            */
            ViewData["paging"] = "";
            ViewData["StartYear"] = startYear;
            ViewData["StartMonth"] = startMonth;
            ViewData["EmpId"] = Request.Form["EmpId"].ToString();

            //DateTime startDate = new DateTime(startYear, startMonth, 1);
            //DateTime endDate = startDate.AddMonths(1).AddDays(-1);

            return PartialView("_OTFilterResult", dt);
        }
        public JsonResult _OTInputSubmit(FormCollection form)
        {
            string mes = "";
            try
            {
                if (payfun.sessionRecreate() == "expires")
                {
                    return Json("Session expires");
                }
                SortedList list = new SortedList();
                decimal total_records = 0;
                decimal pageno = 1;
                decimal pagesize = 15;
                string fromDate = Request.Form["FromDate"].ToString();
                fromDate = Convert.ToDateTime(fromDate).ToString("dd-MMM-yyyy");
                //int startYear = Convert.ToDateTime(fromDate).Year;
                //int startMonth = Convert.ToDateTime(fromDate).Month;
                list.Add("@Role", Enum.GetName(typeof(payrollFunctions.roleIds), Convert.ToInt32(Session["RoleId"].ToString())));
                //DateTime startDate = new DateTime(startYear, startMonth, 1);
                DateTime endDate = Convert.ToDateTime(fromDate).AddMonths(1).AddDays(-1);

                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject(Request.Form["FreezeEmpIds"].ToString());

                DataTable dt = new DataTable();
                dt.Columns.Add("EmpId", typeof(int));
                dt.Columns.Add("AttendanceTypeId", typeof(int));
                DataRow dr;

                foreach (var item in jsonData)
                {
                    dr = dt.NewRow();
                    dr["EmpId"] = item.EmpId;
                    dr["AttendanceTypeId"] = 0;
                    dt.Rows.Add(dr);
                }
                list = payfun.globalParameterList(form, list);


                SqlCommand com = new SqlCommand();
                connection conObj = new connection();

                com.Connection = conObj.con;

                com.CommandType = CommandType.StoredProcedure;
                if (Request.Form["Method"].ToString() == "OTCalculation")
                {
                    com.CommandText = "CMRAttendance_SalaryOTCalculation";
                }


                foreach (string li in list.Keys)
                {
                    if (!li.Contains("FreezeEmpIds") && !li.Contains("FromDate") && !li.Contains("Method"))
                    {
                        com.Parameters.AddWithValue(li, list[li].ToString().Trim());
                    }
                }
                // com.Parameters.AddWithValue("@MonthYearId", startMonth.ToString() + startYear.ToString());
                // com.Parameters.AddWithValue("@LockDate", endDate.ToString("dd-MMM-yyyy"));
                //com.Parameters.AddWithValue("@CreatedDate", comfun.dateISTstr());
                // com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());

                com.Parameters.AddWithValue("@FromDate", fromDate);
                com.Parameters.AddWithValue("@ToDate", endDate.ToString("dd-MMM-yyyy"));
                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@AttendanceTable";

                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = dt;
                com.Parameters.Add(parameter);
                com.CommandTimeout = 0;

                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();
                com.ExecuteNonQuery();
                conObj.con.Close();
                mes = "OT has been Calculate successfully";
            }
            catch (Exception ex)
            {
                mes = "Error:" + ex.Message;
            }
            return Json(mes);
        }
        public JsonResult _OTInputSubmitFinal(string CalcData, DateTime FromDate)
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            //if (payfun.LeaveApply(Convert.ToDateTime(Request.Form["FromDate"])) == "N")
            //{
            //    if (Session["RoleId"].ToString() == ((int)Enum.Parse(typeof(payrollFunctions.roleIds), "Manager")).ToString())
            //    {
            //        return Json("Error: Can not approve leave in back date");
            //    }
            //}
            string mes = string.Empty;

            try
            {
                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject(CalcData);
                DataTable table = new DataTable();

                table.Columns.Add("Id", typeof(int));
                table.Columns.Add("EmpId", typeof(int));
                table.Columns.Add("hrs", typeof(string));

                foreach (var item in jsonData)
                {
                    DataRow dr = table.NewRow();
                    dr["Id"] = item.Id;
                    dr["EmpId"] = item.EmpId;
                    dr["hrs"] = item.hrs;

                    table.Rows.Add(dr);
                }

                SqlCommand com = new SqlCommand();
                connection conObj = new connection();

                com.Connection = conObj.con;

                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "stpVIKSATpaySalaryCalculationOTComplete_Final";
                // com.Parameters.AddWithValue("@SalaryOtDatafinal", comfun.dateISTstr());
                com.Parameters.AddWithValue("@fdDate", FromDate.ToString("Myyyy"));

                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@SalaryOtDatafinal";

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
                mes = comfun.errorMessage("Attendance _LeaveApprovalSubmit", "Error: " + ex.Message);
            }
            return Json(mes);
        }

        /// <summary>
        /// Emp Salary PF
        /// </summary>
        /// <returns></returns>
        public ActionResult EmpSalaryPF()
        {
            return View();
        }
        public JsonResult _BranchList()
        {
            DataTable dt = comfun.fillDataTable("stpBranchddl", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult _EmpSalaryPFList()
        {
            SortedList list = new SortedList();
            list.Add("@BranchID", Request.Form["BranchId"].ToString());
            list.Add("@MonthDate", Convert.ToDateTime(Request.Form["CurrentMonth"]).ToString("dd-MMM-yyyy"));
            DataTable dt = comfun.fillDataTable("stpSalaryPF_ECR", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }


        /// <summary>
        /// Emp ESI
        /// </summary>
        /// <returns></returns>
        public ActionResult EmpESI()
        {
            return View();
        }
        public JsonResult _EmpESIList()
        {
            SortedList list = new SortedList();
            list.Add("@fiMonthYear", Convert.ToInt32(Request.Form["MonthYear"].ToString()));
            DataTable dt = comfun.fillDataTable("stpSalaryESIList", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }


        /// <summary>
        /// Emp Salary Gratuity
        /// </summary>
        /// <returns></returns>
        public ActionResult EmpSalaryGratuity()
        {
            return View();
        }
        public JsonResult _EmpSalaryGratuityList()
        {
            SortedList list = new SortedList();
            list.Add("@BranchID", Request.Form["BranchId"].ToString());
            list.Add("@MonthYear", Convert.ToInt32(Request.Form["MonthYear"].ToString()));
            DataTable dt = comfun.fillDataTable("stpSalaryGratuity_list", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }





        //*******************************************************Employee Investment Section*****************************************************//
        public ActionResult EmployeeInvestment()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            comfun.saveformname("EmployeeInvestment", "/Payroll/EmployeeInvestment", "Income Tax", "Income Tax", "Y", "EmployeeInvestment", "EmployeeInvestment", "Add");
            return View();
        }

        //*******************************************************TDS Section*********************************************************//

        //TDS Section Display (Partial View)//
        public ActionResult TDSList()
        {
            return View();
        }
        public JsonResult _TDSSectionDetails()
        {
            DataTable dt = comfun.fillDataTable("TDSSection_Display", "", null);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1, JsonRequestBehavior.AllowGet);
        }
        //TDS Section Save (Partial View)//
        public ActionResult TDSSectionAdd()
        {
            return PartialView("_TDSSectionAdd");
        }

        public JsonResult _TDSSectionAddSubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@TDSId", Request.Form["TDSId"].ToString());
                list.Add("@TDSName", Request.Form["SectionName"].ToString());
                list.Add("@TDSAmount", Request.Form["TDSAmount"].ToString());
                list.Add("@TDSExemption", Request.Form["TDSExemption"].ToString());
                list.Add("@TDSCreatedDate", comfun.dateISTstr());
                list.Add("@TDSCreatedBy", Session["EmpId"].ToString());
                mes = comfun.executeNonQueryWMessage("TDSSection_Save_Update", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Add", ex.Message);
            }
            return Json(mes);
        }

        //TDS Section Update (Partial View)//
        public ActionResult TDSSectionUpdate()
        {
            SortedList list = new SortedList();
            list.Add("@TDSId", Request.Form["TDSSectionId"].ToString());
            DataTable dt = comfun.fillDataTable("TDSSection_SelectWithId", "", list);
            return PartialView("_TDSSectionUpdate", dt);
        }



        //*******************************************************Investment Section*****************************************************//
        public ActionResult EmployeeInvestmentbyEmployee()
        {
            return View();
        }

        public ActionResult _EmployeeInvestmentbyEmployeeList()
        {
            SortedList list = new SortedList();
            list.Add("@EmpId", Session["EmpId"].ToString());

            DataTable dt = comfun.fillDataTable("Investment_DisplayEmployeeList", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1, JsonRequestBehavior.AllowGet);
        }
        //Investment Section Display (Partial View)//
        public ActionResult EmployeeInvestmentList()
        {
            // DataSet ds = comfun.fillDataSet("Investment_Display", "", null);
            return View();
        }
        public ActionResult _EmployeeInvestmentList()
        {
            DataTable dt = comfun.fillDataTable("Investment_Display", "", null);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1, JsonRequestBehavior.AllowGet);
        }

        //Investment Section Save (Partial View)//

        public ActionResult _Sectionddl()
        {
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["SectionId"].ToString());
            DataTable dt = comfun.fillDataTable("stpTDSSectionddl", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1, JsonRequestBehavior.AllowGet);
        }


        public JsonResult _InvestmentAddSubmit()
        {
            string mes = string.Empty;


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

                        ImageName = "EmployeeInvestment" + Request.Form["EmployeeId"].ToString() + fname + _ext;
                        string filePath = Path.Combine(Server.MapPath("~/EmployeeInvestment/") + ImageName);

                        fname = filePath;
                        var filePathA = filePath;
                        file.SaveAs(filePathA);
                        SortedList list = new SortedList();
                        list.Add("@InvestmentId", Request.Form["Id"].ToString());
                        list.Add("@EmployeeId", Request.Form["EmployeeId"].ToString());
                        //list.Add("@EmployeeNameWithId", Request.Form["EmployeeNameWithId"].ToString());                
                        //list.Add("@InvestmentCode", Request.Form["InvestmentCode"].ToString());
                        //list.Add("@InvestmentName", Request.Form["InvestmentName"].ToString());
                        list.Add("@InvestmentSection", Request.Form["InvestmentSection"].ToString());
                        list.Add("@InvestmentDesc", Request.Form["InvestmentDesc"].ToString());
                        list.Add("@InvestmentAmt", Request.Form["InvestmentAmt"].ToString());
                        list.Add("@InvestmentExemption", Request.Form["InvestmentExemption"].ToString());
                        list.Add("@InvestmentTaxable", Request.Form["InvestmentTaxable"].ToString());
                        list.Add("@InvestmentCreatedDate", comfun.dateISTstr());
                        list.Add("@InvestmentCreatedBy", Session["EmpId"].ToString());
                        list.Add("@fvImage", ImageName);
                        list.Add("@fvImagePath", fname);

                        mes = comfun.executeNonQueryWMessage("Investment_Save_Update", "", list).ToString();





                    }
                }
                // Returns message that successfully uploaded  

            }
            catch (Exception ex)
            {
                return Json("Error details: " + ex.Message);
            }


            return Json(mes);
        }

        //Investment Update (Partial View)//
        public ActionResult InvestmentUpdate()
        {
            SortedList list = new SortedList();
            list.Add("@InvestmentId", Request.Form["Id"].ToString());
            DataTable dt = comfun.fillDataTable("Investment_SelectWithId", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1, JsonRequestBehavior.AllowGet);
        }



        //********************************************************HRA Section********************************************************//



        public ActionResult HRAByEmployee()
        {
            comfun.saveformname("HRAByEmployee", "/Payroll/HRAByEmployee", "Employee Investment/HRA", "HRA form", "Y", "HRAByEmployee", "HRAByEmployee", "View", 1);

            return View();
        }

        public ActionResult HRAList()
        {
            comfun.saveformname("HRAList", "/Payroll/HRAList", "Employee Investment/HRA", "HRA form", "Y", "HRAList", "HRAList", "View", 1);

            return View();
        }
        //HRA Section Display (Partial View)//
        public ActionResult _HRAList()
        {
            comfun.saveformname("_HRAList", "/Payroll/_HRAList", "Employee Investment/HRA", "HRA form", "N", "HRAList", "HRAList", "List", 2);
            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmployeeId"].ToString());
            DataTable dt = comfun.fillDataTable("HRA_Display", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1, JsonRequestBehavior.AllowGet);
        }

        //HRA Save (Partial View)//
        public ActionResult HRAAdd()
        {
            return PartialView("_HRAAdd");
        }

        public JsonResult _HRAAddSubmit()
        {
            comfun.saveformname("_HRAAddSubmit", "/Payroll/_HRAAddSubmit", "Employee Investment/HRA", "HRA form", "N", "HRAList", "HRAList", "Add", 2);
            comfun.saveformname("_HRAAddSubmit", "/Payroll/_HRAAddSubmit", "Employee Investment/HRA", "HRA form", "N", "HRAList", "HRAList", "Edit", 3);
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

                            ImageName = "HRA" + Request.Form["EmployeeId"].ToString() + fname + _ext;
                            string filePath = Path.Combine(Server.MapPath("~/HRAReciept/") + ImageName);

                            fname = filePath;
                            var filePathA = filePath;
                            file.SaveAs(filePathA);
                            SortedList list = new SortedList();
                            list.Add("@HRAId", Request.Form["HRAId"].ToString());
                            list.Add("@EmployeeId", Request.Form["EmployeeId"].ToString());
                            ///list.Add("@EmployeeNameWithId", Request.Form["EmployeeNameWithId"].ToString());
                            //list.Add("@HRACode", Request.Form["HRACode"].ToString());
                            //list.Add("@HRAName", Request.Form["HRAName"].ToString());
                            list.Add("@HRAPaid", Request.Form["HRAPaid"].ToString());
                            list.Add("@HRAAmt", Request.Form["HRAAmt"].ToString());
                            list.Add("@HRACreatedDate", comfun.dateISTstr());
                            list.Add("@HRACreatedBy", Session["EmpId"].ToString());


                            list.Add("@fvImageName", ImageName);
                            list.Add("@file", fname);

                            mes = comfun.executeNonQueryWMessage("HRA_Save_Update", "", list).ToString();

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

        //HRA Update (Partial View)//
        public ActionResult HRAUpdate()
        {
            SortedList list = new SortedList();
            list.Add("@HRAId", Request.Form["HRAId"].ToString());
            DataTable dt = comfun.fillDataTable("HRA_SelectWithId", "", list);
            return PartialView("_HRAUpdate", dt);
        }



        //*******************************************************Previous Employee Income Section*****************************************************//

        //Previous Employee Income Display (Partial View)//
        public ActionResult PreviousEmployeeIncomeDetailsList()
        {
            return View();
        }
        public ActionResult _PreviousEmployeeIncomeDetailsList()
        {
            DataTable dt = comfun.fillDataTable("PreviousEmployerIncome_Display", "", null);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1, JsonRequestBehavior.AllowGet);
        }

        //Previous Employee Income Save (Partial View)//


        public JsonResult _PreviousEmployeeIncomeAddSubmit()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@PrevEmployerId", Request.Form["PrevEmployerId"].ToString());
                list.Add("@EmployeeId", Request.Form["EmployeeId"].ToString());
                //list.Add("@PrevEmployerCode", Request.Form["PrevEmployeeCode"].ToString());
                //list.Add("@PrevEmployerName", Request.Form["PrevEmployeeName"].ToString());
                list.Add("@fnPaidAmount", Request.Form["PaidAmount"].ToString());
                list.Add("@PrevEmployerIncome", Request.Form["PrevEmployeeIncome"].ToString());
                list.Add("@PrevEmployerCreatedDate", comfun.dateISTstr());
                list.Add("@PrevEmployerCreatedBy", Session["EmpId"].ToString());
                mes = comfun.executeNonQueryWMessage("PreviousEmployerIncome_Save_Update", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Add", ex.Message);
            }
            return Json(mes);
        }

        //Previous Employee Income Update (Partial View)//
        public ActionResult PreviousEmployeeIncomeUpdate()
        {
            SortedList list = new SortedList();
            list.Add("@PrevEmployerId", Request.Form["PrevEmployerId"].ToString());
            DataTable dt = comfun.fillDataTable("PreviousEmployerIncome_SelectWithId", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1, JsonRequestBehavior.AllowGet);
        }



        //**********************************************************Reimbursement Cancel Section*****************************************************************//

        public ActionResult ReimbursementCancel()
        {
            comfun.saveformname("ReimbursementCancel", "/Payroll/ReimbursementCancel", "Reimbursement Cancel", "Reimbursement Cancel", "Y", "Reimbursement Cancel", "Reimbursement Cancel", "List");
            return View();
        }

        public ActionResult _ReimbursementCancelList()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return PartialView("_sessionExpired");
            }
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@fdDate", Convert.ToDateTime(Request.Form["Month"]).ToString("dd-MMM-yyyy"));
            dt = comfun.fillDataTable("stpViksatReimbursementCancelByemployee", "", list);
            ViewData["MonthYear"] = Request.Form["Month"].ToString();
            return PartialView("_ReimbursementCancelList", dt);
        }

        public JsonResult _ReimbursementCancelSubmitView()
        {
            string mes = "Reimburesement Cancelled Succesfully!!";
            SortedList list = new SortedList();
            dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject(Request.Form["ChkCancel"].ToString());
            //string[] chkboxval = jsonData;
            foreach (var item in jsonData)
            {
                list.Clear();
                list.Add("@ReimburesementId", item);
                comfun.executeNonQuery("stpViksatReimbursementCancel", "", list);
            }
            return Json(mes);
        }

        //***********************************************************Employee CTC Display Previous***************************************************************//

        //To Display Employee Previous CTC//
        public ActionResult EmpCTCDisplayOld()
        {
            comfun.saveformname("EmployeeCTCDisplayOld", "/Payroll/EmployeeCTCDisplayOld", "Employee CTC Display Old", "Employee CTC Display Old", "Y", "Employee CTC Display Old", "Employee CTC Display Old", "List");
            if (payfun.sessionRecreate() == "expires")
            {
                return RedirectToAction("Login", "account");
            }
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            return View();
        }

        //To Display Employee Previous CTC (Partial View Part 1)//
        public ActionResult _EmpCTCDisplayOldPart1()
        {
            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@FdDate", comfun.dateISTstr());
            DataSet ds = comfun.fillDataSet("stpVIKSATEmployeeCTCOld_Display", "", list);
            return PartialView("_EmpCTCDisplayOldPart1", ds);
        }

        //To Display Employee Previous CTC Part (Partial View Part 2)//
        public ActionResult _EmpCTCDisplayOldPart2()
        {
            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@FdDate", Request.Form["FromDate"].ToString());
            DataSet ds = comfun.fillDataSet("stpVIKSATEmployeeCTCOld_Display", "", list);
            return PartialView("_EmpCTCDisplayOldPart2", ds);
        }

        //*****************************************************************************************************************************************************//
        public ActionResult freezeArear()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            comfun.saveformname("freezeArear", "/Payroll/freezeArear", "freezeArear", "freezeArear", "Y", "freezeArear", "freezeArear", "List");
            return View("freezeArearInput");
        }
        public ActionResult _freezeArearFilters()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            ViewData["CostCenterName"] = Session["CostCenterName"].ToString();
            ViewData["CostCenterId"] = Session["CostCenterId"].ToString();
            return PartialView("_freezeArearFilters");
        }
        public ActionResult _freezeArearFilterResult(FormCollection form)
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string type = Request.Form["Type"].ToString();

            ViewData["ShowDays"] = "N";

            string proc = "";



            if (Request.Form["Method"].ToString() == "ArearDays")
            {
                ViewData["ShowDays"] = "Y";
                if (type == "Freeze")
                {
                    proc = "LockArear_Freeze";
                }
                if (type == "UnFreeze")
                {
                    proc = "LockArear_UnFreeze";
                }
            }

            if (Request.Form["Method"].ToString() == "ArearPI")
            {
                ViewData["ShowDays"] = "Y";
                if (type == "Freeze")
                {
                    proc = "LockArearPI_Freeze";
                }
                if (type == "UnFreeze")
                {
                    proc = "LockArearPI_UnFreeze";
                }
            }

            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 15;
            string fromDate = Request.Form["FromDate"].ToString();
            int startYear = Convert.ToDateTime(fromDate).Year;
            int startMonth = Convert.ToDateTime(fromDate).Month;

            DateTime startDate = new DateTime(startYear, startMonth, 1);
            DateTime endDate = startDate.AddMonths(1).AddDays(-1);

            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }
            if (Request.Form["PageNo"] != null)
            {
                pagesize = Convert.ToDecimal(Request.Form["PageSize"]);
            }
            SortedList list = new SortedList();
            //paging_get_downline1
            list.Add("@MonthYearId", startDate.Month.ToString() + startDate.Year.ToString());
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            //***************Parameters*******************************************
            list.Add("@FromDate", startDate.ToString("dd/MMM/yyyy"));
            list.Add("@ToDate", endDate.ToString("dd/MMM/yyyy"));
            list.Add("@loginId", Convert.ToInt32(Session["EmpId"].ToString()));
            list.Add("@Role", Enum.GetName(typeof(payrollFunctions.roleIds), Convert.ToInt32(Session["RoleId"].ToString())));
            if (Request.Form["CostCenterId"].ToString() != "0" && Request.Form["CostCenterId"].ToString() != "")
            {
                list.Add("@CostCenterId", Request.Form["CostCenterId"].ToString());
            }
            else
            {
                list.Add("@CostCenterId", Session["CostCenterId"].ToString());
            }
            list = payfun.globalParameterList(form, list);
            //DataTable dt = comfun.fillDataTable("Attendance_SalaryForAttLock", "", list);
            list.Remove("@Method");
            list.Remove("@Type");
            DataTable dt = comfun.fillDataTable(proc, "", list);

            /*
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "Attendance", "_salaryAttMonthly", "_salaryAttMonthly");

            HtmlString htm = new HtmlString(paging);
             ViewData["paging"] = htm;
            */
            ViewData["paging"] = "";
            ViewData["StartYear"] = startYear;
            ViewData["StartMonth"] = startMonth;
            ViewData["EmpId"] = Request.Form["EmpId"].ToString();

            //DateTime startDate = new DateTime(startYear, startMonth, 1);
            //DateTime endDate = startDate.AddMonths(1).AddDays(-1);

            return PartialView("_freezeFilterResult", dt);
        }
        public JsonResult _freezeArearInputSubmit(FormCollection form)
        {
            string mes = "";
            try
            {
                if (payfun.sessionRecreate() == "expires")
                {
                    return Json("Session expires");
                }
                SortedList list = new SortedList();
                decimal total_records = 0;
                decimal pageno = 1;
                decimal pagesize = 15;
                string fromDate = Request.Form["FromDate"].ToString();
                int startYear = Convert.ToDateTime(fromDate).Year;
                int startMonth = Convert.ToDateTime(fromDate).Month;

                DateTime startDate = new DateTime(startYear, startMonth, 1);
                DateTime endDate = startDate.AddMonths(1).AddDays(-1);

                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject(Request.Form["FreezeEmpIds"].ToString());

                DataTable dt = new DataTable();
                dt.Columns.Add("EmpId", typeof(int));
                dt.Columns.Add("AttendanceTypeId", typeof(int));
                DataRow dr;

                foreach (var item in jsonData)
                {
                    dr = dt.NewRow();
                    dr["EmpId"] = item.EmpId;
                    dr["AttendanceTypeId"] = 0;
                    dt.Rows.Add(dr);
                }

                list = payfun.globalParameterList(form, list);
                SqlCommand com = new SqlCommand();
                connection conObj = new connection();

                com.Connection = conObj.con;

                com.CommandType = CommandType.StoredProcedure;
                if (Request.Form["Method"].ToString() == "ArearDays")
                {
                    if (Request.Form["Type"].ToString() == "Freeze")
                    {
                        com.CommandText = "LockArear_FreezeSubmit";
                    }
                }

                if (Request.Form["Method"].ToString() == "ArearPI")
                {
                    if (Request.Form["Type"].ToString() == "Freeze")
                    {
                        com.CommandText = "Arear_TimeLockSubmit";
                    }
                }



                foreach (string li in list.Keys)
                {
                    if (!li.Contains("FreezeEmpIds") && !li.Contains("FromDate") && !li.Contains("Method"))
                    {
                        com.Parameters.AddWithValue(li, list[li].ToString().Trim());
                    }
                }
                com.Parameters.AddWithValue("@MonthYearId", startMonth.ToString() + startYear.ToString());
                com.Parameters.AddWithValue("@LockDate", endDate.ToString("dd-MMM-yyyy"));
                com.Parameters.AddWithValue("@CreatedDate", comfun.dateISTstr());
                com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());

                //com.Parameters.AddWithValue("@FromDate", startDate);
                //com.Parameters.AddWithValue("@ToDate", endDate);

                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@AttendanceTable";

                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = dt;
                com.Parameters.Add(parameter);


                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();
                com.ExecuteNonQuery();
                conObj.con.Close();
                mes = "Lock has been applied successfully";
            }
            catch (Exception ex)
            {
                mes = "Error:" + ex.Message;
            }
            return Json(mes);
        }


        public ActionResult ArearCalculation()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            comfun.saveformname("ArearCalculation", "/Payroll/ArearCalculation", "ArearCalculation", "ArearCalculation", "Y", "ArearCalculation", "ArearCalculation", "List");
            return View("ArearCalculation");
        }

        public ActionResult _Arearfilters()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            ViewData["CostCenterName"] = Session["CostCenterName"].ToString();
            ViewData["CostCenterId"] = Session["CostCenterId"].ToString();
            return PartialView("_Arearfilters");
        }

        public ActionResult _ArearFilterResult(FormCollection form)
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 15;
            string fromDate = Request.Form["FromDate"].ToString();
            int startYear = Convert.ToDateTime(fromDate).Year;
            int startMonth = Convert.ToDateTime(fromDate).Month;

            DateTime startDate = new DateTime(startYear, startMonth, 1);
            DateTime endDate = startDate.AddMonths(1).AddDays(-1);


            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }
            if (Request.Form["PageNo"] != null)
            {
                pagesize = Convert.ToDecimal(Request.Form["PageSize"]);
            }
            SortedList list = new SortedList();
            //paging_get_downline1
            list.Add("@MonthYearId", startDate.Month.ToString() + startDate.Year.ToString());
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            //***************Parameters*******************************************
            list.Add("@FromDate", Convert.ToDateTime(startDate).ToString("dd-MMM-yyyy"));
            list.Add("@ToDate", Convert.ToDateTime(endDate).ToString("dd-MMM-yyyy"));
            list.Add("@loginId", Convert.ToInt32(Session["EmpId"].ToString()));
            list.Add("@Role", Enum.GetName(typeof(payrollFunctions.roleIds), Convert.ToInt32(Session["RoleId"].ToString())));
            if (Request.Form["CostCenterId"].ToString() != "0" && Request.Form["CostCenterId"].ToString() != "")
            {
                list.Add("@CostCenterId", Request.Form["CostCenterId"].ToString());
            }
            else
            {
                list.Add("@CostCenterId", Session["CostCenterId"].ToString());
            }
            list = payfun.globalParameterList(form, list);
            //DataTable dt = comfun.fillDataTable("Attendance_SalaryForAttLock", "", list);

            DataTable dt = comfun.fillDataTable("LockAttendanceForArear_UnFreeze", "", list);


            ViewData["paging"] = "";
            ViewData["StartYear"] = startYear;
            ViewData["StartMonth"] = startMonth;
            ViewData["EmpId"] = Request.Form["EmpId"].ToString();



            return PartialView("_ArearFilterResult", dt);
        }
        public JsonResult _ArearInputSubmit(FormCollection form)
        {
            string mes = "";
            try
            {
                if (payfun.sessionRecreate() == "expires")
                {
                    return Json("Session expires");
                }
                SortedList list = new SortedList();
                decimal total_records = 0;
                decimal pageno = 1;
                decimal pagesize = 15;
                string fromDate = Request.Form["FromDate"].ToString();
                int startYear = Convert.ToDateTime(fromDate).Year;
                int startMonth = Convert.ToDateTime(fromDate).Month;

                DateTime startDate = new DateTime(startYear, startMonth, 1);
                DateTime endDate = startDate.AddMonths(1).AddDays(-1);

                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject(Request.Form["FreezeEmpIds"].ToString());

                DataTable dt = new DataTable();
                dt.Columns.Add("EmpId", typeof(int));
                dt.Columns.Add("AttendanceTypeId", typeof(int));
                DataRow dr;

                foreach (var item in jsonData)
                {
                    dr = dt.NewRow();
                    dr["EmpId"] = item.EmpId;
                    dr["AttendanceTypeId"] = 0;
                    dt.Rows.Add(dr);
                }
                list = payfun.globalParameterList(form, list);
                SqlCommand com = new SqlCommand();
                connection conObj = new connection();

                com.Connection = conObj.con;

                com.CommandType = CommandType.StoredProcedure;
                if (Request.Form["Method"].ToString() == "ArearCalculation")
                {
                    com.CommandText = "CMRArearCalculation";
                }
                //if (Request.Form["Method"].ToString() == "PICalculation")
                //{
                //    com.CommandText = "CMRArearCalculation";
                //}


                foreach (string li in list.Keys)
                {
                    if (!li.Contains("FreezeEmpIds") && !li.Contains("FromDate") && !li.Contains("Method"))
                    {
                        com.Parameters.AddWithValue(li, list[li].ToString().Trim());
                    }
                }
                // com.Parameters.AddWithValue("@MonthYearId", startMonth.ToString() + startYear.ToString());
                // com.Parameters.AddWithValue("@LockDate", endDate.ToString("dd-MMM-yyyy"));
                //com.Parameters.AddWithValue("@CreatedDate", comfun.dateISTstr());
                // com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());
                com.Parameters.AddWithValue("@FromDate", startDate.ToString("dd-MMM-yyyy"));
                com.Parameters.AddWithValue("@ToDate", endDate.ToString("dd-MMM-yyyy"));

                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@AttendanceTable";

                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = dt;
                com.Parameters.Add(parameter);
                com.CommandTimeout = 0;

                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();
                com.ExecuteNonQuery();
                conObj.con.Close();
                mes = "Arear has been Calculate successfully";
            }
            catch (Exception ex)
            {
                mes = "Error:" + ex.Message;
            }
            return Json(mes);
        }

        public ActionResult ArearPI()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            comfun.saveformname("ArearPI", "/Payroll/ArearPI", "ArearPI", "ArearPI", "Y", "ArearPI", "ArearPI", "List");
            return View("ArearPICalculation");
        }
        public ActionResult _ArearPIfilters()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            ViewData["CostCenterName"] = Session["CostCenterName"].ToString();
            ViewData["CostCenterId"] = Session["CostCenterId"].ToString();
            return PartialView("_ArearPIfilters");
        }
        public ActionResult _ArearPIFilterResult(FormCollection form)
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            decimal total_records = 0;
            decimal pageno = 1;
            decimal pagesize = 15;
            string fromDate = Request.Form["FromDate"].ToString();
            int startYear = Convert.ToDateTime(fromDate).Year;
            int startMonth = Convert.ToDateTime(fromDate).Month;

            DateTime startDate = new DateTime(startYear, startMonth, 1);
            DateTime endDate = startDate.AddMonths(1).AddDays(-1);
            string type = Request.Form["Type"].ToString();
            string Method = Request.Form["Method"].ToString();
            string proc = "";
            if (Request.Form["Method"].ToString() == "ArearPICal")
            {
                ViewData["Showhrs"] = "N";
                proc = "LockArearForArear_PI";

            }


            if (Request.Form["PageNo"] != null)
            {
                pageno = Convert.ToDecimal(Request.Form["PageNo"]);
            }
            if (Request.Form["PageNo"] != null)
            {
                pagesize = Convert.ToDecimal(Request.Form["PageSize"]);
            }
            SortedList list = new SortedList();
            //paging_get_downline1
            list.Add("@MonthYearId", startDate.Month.ToString() + startDate.Year.ToString());
            list.Add("@PageNo", pageno);
            list.Add("@PageSize", pagesize);
            //***************Parameters*******************************************
            list.Add("@FromDate", startDate.ToString("dd-MMM-yyyy"));
            list.Add("@ToDate", endDate.ToString("dd-MMM-yyyy"));
            list.Add("@loginId", Convert.ToInt32(Session["EmpId"].ToString()));
            list.Add("@Role", Enum.GetName(typeof(payrollFunctions.roleIds), Convert.ToInt32(Session["RoleId"].ToString())));
            if (Request.Form["CostCenterId"].ToString() != "0" && Request.Form["CostCenterId"].ToString() != "")
            {
                list.Add("@CostCenterId", Request.Form["CostCenterId"].ToString());
            }
            else
            {
                list.Add("@CostCenterId", Session["CostCenterId"].ToString());
            }
            list = payfun.globalParameterList(form, list);
            list.Remove("@Method");
            list.Remove("@Type");

            DataTable dt = comfun.fillDataTable(proc, "", list);
            /*
            if (dt.Rows.Count > 0)
                total_records = Convert.ToDecimal(dt.Rows[0]["total"]);

            string paging = comfun.create_paging_ajax(total_records, pagesize, pageno, "Attendance", "_salaryAttMonthly", "_salaryAttMonthly");

            HtmlString htm = new HtmlString(paging);
             ViewData["paging"] = htm;
            */
            ViewData["paging"] = "";
            ViewData["StartYear"] = startYear;
            ViewData["StartMonth"] = startMonth;
            ViewData["EmpId"] = Request.Form["EmpId"].ToString();

            //DateTime startDate = new DateTime(startYear, startMonth, 1);
            //DateTime endDate = startDate.AddMonths(1).AddDays(-1);

            return PartialView("_ArearPIFilterResult", dt);
        }
        public JsonResult _ArearPIInputSubmit(FormCollection form)
        {
            string mes = "";
            try
            {
                if (payfun.sessionRecreate() == "expires")
                {
                    return Json("Session expires");
                }
                SortedList list = new SortedList();
                decimal total_records = 0;
                decimal pageno = 1;
                decimal pagesize = 15;
                string fromDate = Request.Form["FromDate"].ToString();
                fromDate = Convert.ToDateTime(fromDate).ToString("dd-MMM-yyyy");
                //int startYear = Convert.ToDateTime(fromDate).Year;
                //int startMonth = Convert.ToDateTime(fromDate).Month;
                list.Add("@Role", Enum.GetName(typeof(payrollFunctions.roleIds), Convert.ToInt32(Session["RoleId"].ToString())));
                //DateTime startDate = new DateTime(startYear, startMonth, 1);
                DateTime endDate = Convert.ToDateTime(fromDate).AddMonths(1).AddDays(-1);

                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject(Request.Form["FreezeEmpIds"].ToString());

                DataTable dt = new DataTable();
                dt.Columns.Add("EmpId", typeof(int));
                dt.Columns.Add("AttendanceTypeId", typeof(int));
                DataRow dr;

                foreach (var item in jsonData)
                {
                    dr = dt.NewRow();
                    dr["EmpId"] = item.EmpId;
                    dr["AttendanceTypeId"] = 0;
                    dt.Rows.Add(dr);
                }
                list = payfun.globalParameterList(form, list);


                SqlCommand com = new SqlCommand();
                connection conObj = new connection();

                com.Connection = conObj.con;

                com.CommandType = CommandType.StoredProcedure;
                if (Request.Form["Method"].ToString() == "ArearPICalculation")
                {
                    com.CommandText = "CMRArearPICalculation";
                }


                foreach (string li in list.Keys)
                {
                    if (!li.Contains("FreezeEmpIds") && !li.Contains("FromDate") && !li.Contains("Method"))
                    {
                        com.Parameters.AddWithValue(li, list[li].ToString().Trim());
                    }
                }
                // com.Parameters.AddWithValue("@MonthYearId", startMonth.ToString() + startYear.ToString());
                // com.Parameters.AddWithValue("@LockDate", endDate.ToString("dd-MMM-yyyy"));
                //com.Parameters.AddWithValue("@CreatedDate", comfun.dateISTstr());
                // com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());

                com.Parameters.AddWithValue("@FromDate", fromDate);
                com.Parameters.AddWithValue("@ToDate", endDate.ToString("dd-MMM-yyyy"));
                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@AttendanceTable";

                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = dt;
                com.Parameters.Add(parameter);
                //com.CommandTimeout = 0;

                if (conObj.con.State == ConnectionState.Closed)
                    conObj.con.Open();
                com.ExecuteNonQuery();
                conObj.con.Close();
                // mes = "OT has been Calculate successfully";
            }
            catch (Exception ex)
            {
                mes = "Error:" + ex.Message;
            }
            return Json(mes);
        }
        public ActionResult LeaveEncashment()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return RedirectToAction("Login", "account");
            }
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            comfun.saveformname("LeaveEncashment", "/Payroll/LeaveEncashment", "LeaveEncashment", "LeaveEncashment", "Y", "LeaveEncashment", "LeaveEncashment", "List");

            return View();
        }
        public ActionResult _LeaveEncashment(FormCollection form)
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return PartialView("_sessionExpired");
            }

            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            //list.Add("@PageNo", Session["PageNo"].ToString());

            list.Add("@loginId", Convert.ToInt32(Session["EmpId"].ToString()));

            list.Add("@Role", Enum.GetName(typeof(payrollFunctions.roleIds), Convert.ToInt32(Session["RoleId"].ToString())));
            list = payfun.globalParameterList(form, list);
            dt = comfun.fillDataTable("stpVIKSATLeaveEncashment", "", list);
            return PartialView("_leaveEncashment", dt);
        }
        public ActionResult PaygroupList()
        {
            comfun.saveformname("PayGroup", "/Payroll/PaygroupList", "Emp Management/Pay Group", "Pay Group Main Form", "Y", "PayGroup", "PayGroup", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("PayGroup");
            return View();

        }

        public ActionResult _PartgroupJsonList()
        {
            comfun.saveformname("_PartgroupJsonList", "/Payroll/_PartgroupJsonList", "Emp Management/Pay Group", "Pay Group List", "N", "PayGroup", "PayGroup", "List", 4);
            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("stpViksatPayPartList", "", null);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult _PaygroupAddView()

        {

            ViewData["PartId"] = "0";
            ViewData["PartName"] = "";

            if (Request.Form["PartId"].ToString() != "0")
            {
                ViewData["PartId"] = Request.Form["PartId"].ToString();
                ViewData["PartName"] = Request.Form["PartName"].ToString();

            }
            return PartialView("_PaygroupAdd");

        }
        public JsonResult _PayPartSubmit()
        {
            comfun.saveformname("_PayPartSubmit", "/Payroll/_PayPartSubmit", "Emp Management/Pay Group", "Pay Group Add", "N", "PayGroup", "PayGroup", "Add", 2);
            comfun.saveformname("_PayPartSubmit", "/Payroll/_PayPartSubmit", "Emp Management/Pay Group", "Pay Group Edit", "N", "PayGroup", "PayGroup", "Edit", 3);

            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@PartId", Request.Form["PartId"].ToString());
                list.Add("@PartName", Request.Form["PartName"].ToString());
                //list.Add("@PartCode", Request.Form["PartCode"].ToString());
                list.Add("@CreatedBy", Session["EmpId"].ToString());
                list.Add("@CreatedDate", comfun.dateISTstr());
                mes = comfun.executeNonQueryWMessage("stpPayPart_Save", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _PayPartSubmit", ex.Message);
            }
            return Json(mes);
        }
        public JsonResult _PayPartStatusUpdate()
        {
            comfun.saveformname("_PayPartStatusUpdate", "/Payroll/_PayPartStatusUpdate", "Emp Management/Pay Group", "Pay Group Status", "N", "PayGroup", "PayGroup", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpPaypartStatusUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }

        //PayCharges
        public ActionResult PayChargesList()
        {
            comfun.saveformname("PayCharges", "/payroll/PayChargesList", "Emp Management/Pay Charges", "Pay Charges Main Form", "Y", "PayCharges", "PayCharges", "View", 1);
            ViewData["AccessRights"] = payfun.getAccessRights("PayCharges");
            //if (Request.QueryString["key"] != null)
            //{
            //    Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            //}
            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("stppaygroupddl", "", null);
            //dt = comfun.fillDataTable("Charges_Select", "", null);

            return View(dt);
        }
        public ActionResult _PayChargelist()
        {
            SortedList list = new SortedList();
            list.Add("@PartId", "0");
            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("Charges_SelectM", "", list);
            return PartialView("_PayChargelist", dt);

        }
        //public JsonResult _PaychargesStatusUpdate()
        //{

        //    string mes = string.Empty;
        //    try
        //    {
        //        SortedList list = new SortedList();
        //        list.Add("@Id", Request.Form["Id"].ToString());
        //        list.Add("@IsActive", Request.Form["IsActive"].ToString());
        //        mes = comfun.executeNonQueryWMessage("stpPaychargesStatusUpdate", "", list).ToString();
        //    }
        //    catch (Exception ex)
        //    {
        //        mes = ex.Message;
        //    }
        //    return Json(mes);
        //}

        public JsonResult _PaychargesJsonList()
        {
            comfun.saveformname("_PaychargesJsonList", "/payroll/_PaychargesJsonList", "Emp Management/Pay Charges", "Pay Charges List", "N", "PayCharges", "PayCharges", "List", 4);
            SortedList list = new SortedList();
            list.Add("@PartId", "0");
            DataTable dt = new DataTable();
            dt = comfun.fillDataTable("Charges_SelectM", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }


        public ActionResult _PayChargeAddView()
        {
            ViewData["ChargeId"] = "0";
            ViewData["ChargeName"] = "";

            if (Request.Form["PartId"].ToString() != "0")
            {
                ViewData["PartId"] = Request.Form["PartId"].ToString();
                ViewData["PartName"] = Request.Form["PartName"].ToString();

            }
            return PartialView("_PayChargeAdd");

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
                list.Add("@PartId", Request.Form["PartId"].ToString());
                list.Add("@PartName", Request.Form["PartName"].ToString());
                //list.Add("@PartCode", Request.Form["PartCode"].ToString());
                list.Add("@CreatedBy", Session["EmpId"]);
                list.Add("@CreatedDate", comfun.dateISTstr());
                mes = comfun.executeNonQueryWMessage("stpPayPart_Save", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("Admin _PayPartSubmit", ex.Message);
            }
            return Json(mes);
        }

        public JsonResult _PayChargesStatusUpdate()
        {
            comfun.saveformname("_PayChargesStatusUpdate", "/payroll/_PayChargesStatusUpdate", "Emp Management/Pay Charges", "Pay Charges Status", "N", "PayCharges", "PayCharges", "Status", 5);

            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpPaychargesStatusUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }
        public JsonResult _PayChargesAddSubmit()
        {
            comfun.saveformname("_PayChargesAddSubmit", "/payroll/_PayChargesAddSubmit", "Emp Management/Pay Charges", "Pay Charges Add", "N", "PayCharges", "PayCharges", "Add", 2);
            comfun.saveformname("_PayChargesEditSubmit", "/payroll/_PayChargesEditSubmit", "Emp Management/Pay Charges", "Pay Charges Edit", "N", "PayCharges", "PayCharges", "Edit", 3);
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@CreatedBy", Session["EmpId"]);
                list.Add("@CreatedDate", comfun.dateISTstr());
                list.Add("@ChargeName", Request.Form["ChargeName"].ToString());
                list.Add("@ChargeType", Request.Form["ChargeType"].ToString());
                list.Add("@CtcType", Request.Form["CtcType"].ToString());
                list.Add("@CalculationType", Request.Form["CalculationType"].ToString());
                list.Add("@IsPF", Request.Form["IsPF"].ToString());
                list.Add("@IsESI", Request.Form["IsESI"].ToString());
                list.Add("@IsLoan", Request.Form["IsLoan"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                list.Add("@Part", Request.Form["Part"].ToString());
                list.Add("@ChargeWEF", Request.Form["ChargeWEF"].ToString());
                list.Add("@IsGross", Request.Form["IsGross"].ToString());
                mes = comfun.executeNonQueryWMessage("Charges_Save", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_PayChargesAddSubmit", ex.Message);
            }
            return Json(mes);
        }

        //   [customAuthorize(Roles = payrollFunctions.roleAdmin + "," + payrollFunctions.roleManger + "," + payrollFunctions.rolePayrollTeam)]
        public ActionResult PayChargesEdit()
        {
            comfun.saveformname("PayChargesEdit", "/admin/PayChargesEdit", "Pay Charges Edit", "Pay Charges Edit", "N", "PayChargesEdit", "PayCharges", "Edit");

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
        public JsonResult _PaychargesEdit1()
        {
            string chargeid = "0";

            chargeid = Request.Form["Id"].ToString();
            SortedList list = new SortedList();
            list.Add("@ChargesId", chargeid);
            DataTable dt = comfun.fillDataTable("Charges_SelectWithId", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);

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
                list.Add("@IsLoan", Request.Form["IsLoan"].ToString());
                //list.Add("@IsActive", Request.Form["IsActive"].ToString());
                list.Add("@ChargesId", Request.Form["ChargeId"].ToString());
                list.Add("@ChargeWEF", Request.Form["ChargeWEF"].ToString());
                list.Add("@IsGross", Request.Form["IsGross"].ToString());
                list.Add("@Part", Request.Form["Part"].ToString());
                mes = comfun.executeNonQueryWMessage("Charges_Edit", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = comfun.errorMessage("_PayChargesEditSubmit", "Error:" + ex.Message);
            }
            return Json(mes);
        }

        public ActionResult SalaryCalculate()
        {
            comfun.saveformname("SalaryCalculate", "/Payroll/SalaryCalculate", "Payroll/SalaryCalculate", "Salary Calculate form", "Y", "SalaryCalculate", "SalaryCalculate", "View", 1);
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
                mes = comfun.executeNonQueryWMessage("stpSalaryCalculation", "", list).ToString();

            }
            catch (Exception ex)
            {

                mes = "Error:" + ex.Message;
            }

            return Json(mes);
        }
        public ActionResult ReimbursementSheet()
        {
            comfun.saveformname("ReimbursementSheet", "/Payroll/ReimbursementSheet", "Payroll/ReimbursementSheet", "Salary Reimbursement Sheet form", "Y", "ReimbursementSheet", "ReimbursementSheet", "View", 1);
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            return View();
        }


        public ActionResult _ReimbursementSheet()
        {
            SortedList list = new SortedList();

            if (Request.Form["FromDate"] != null)
            {
                list.Add("@fdDate", Request.Form["FromDate"].ToString());
            }
            //if (Request.Form["ToDate"] != null)
            //{
            //    list.Add("@EndDate", Convert.ToDateTime(Request.Form["ToDate"]).ToString("dd-MMM-yyyy"));
            //}
            list.Add("@fiMonthYear", Convert.ToDateTime(Request.QueryString["FromDate"]).ToString("MMyyyy"));
            // list.Add("@fiMonthYear", Convert.ToDateTime(Request.QueryString["FromDate"]).ToString("MMyyyy"));
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
            DataTable dt = comfun.fillDataTable("StpVIKSATReimbursementRegisterColumner ", "", list);


            return PartialView("_SalaryReimbursementSheet", dt);
        }
        public ActionResult ImportExcel()
        {
            return View();

        }
        public ActionResult _EmployeeExcelList()
        {

            SortedList list = new SortedList();
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            list.Add("@FromDate", Request.Form["FromDate"].ToString());
            list.Add("@Category", Request.Form["CategoryId"].ToString());
            list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
            list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
            list.Add("@BranchId", Request.Form["BranchId"].ToString());
            list.Add("@CircleId", Request.Form["CircleId"].ToString());
            DataTable dt = comfun.fillDataTable("stpEmployeeAttendanceImportList", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
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
            dtExcel.Columns.Add("TDS", typeof(string));
            dtExcel.Columns.Add("OT", typeof(string));
            dtExcel.Columns.Add("Expense", typeof(int));
            dtExcel.Columns.Add("Bonus", typeof(string));
            dtExcel.Columns.Add("Reward", typeof(string));
            dtExcel.Columns.Add("Vriable", typeof(string));
            dtExcel.Columns.Add("Insentive", typeof(string));

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

                        dr["TDS"] = item["TDS"];
                        dr["OT"] = item["OT"];
                        dr["Expense"] = item["Expense"];
                        dr["Bonus"] = item["Bonus"];
                        dr["Reward"] = item["Reward"];
                        dr["Vriable"] = item["Vriable"];
                        dr["Insentive"] = item["Insentive"];

                        dtExcel.Rows.Add(dr);
                    }
                    string sheetName = "Sheet1";

                    SqlCommand com = new SqlCommand();
                    connection conObj = new connection();

                    com.Connection = conObj.con;


                    com.CommandType = CommandType.StoredProcedure;
                    com.CommandText = "stpEmployeeAttendanceImportexcel_Accept";
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
                    com.CommandTimeout = 0;
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
        public ActionResult EmployeeReimbursementAdd()
        {
            return View();

        }
        public ActionResult _ChargeName()
        {
            DataTable dt = comfun.fillDataTable("stpHtisPayChargeddl", "", null);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }

        public ActionResult _ReimbursementJsonList()
        {
            DataTable dt = comfun.fillDataTable("stpReimbursementList", "", null);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }

        public JsonResult _ReimbursementSubmit()
        {
            string mes = string.Empty;
            try
            {

                SortedList list = new SortedList();
                list.Add("@fiChargeID", Request.Form["ChargeId"].ToString());
                list.Add("@fiEmpID", Request.Form["EmpId"].ToString());
                list.Add("@fnAmount", Request.Form["Amount"].ToString());
                list.Add("@fdWEFDate", Request.Form["WEF"].ToString());

                mes = comfun.executeNonQueryWMessage("StpEmployeeReimbursement_Accept", "", list).ToString();
            }
            catch (Exception ex)
            {

                mes = ex.Message;
            }

            return Json(mes);
        }
        public ActionResult ReimbursementCalculation()
        {
            comfun.saveformname("ReimbursementCalculation", "/Payroll/ReimbursementCalculation", "Payroll/ReimbursementCalculation", "Reimbursement Calculate form", "Y", "ReimbursementCalculation", "ReimbursementCalculation", "View", 1);
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            return View();

        }
        public JsonResult _ReimbursementCalculation()
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
                mes = comfun.executeNonQueryWMessage("stpReimbursementCalculation", "", list).ToString();

            }
            catch (Exception ex)
            {

                mes = "Error:" + ex.Message;
            }

            return Json(mes);
        }
        public ActionResult TDSSetting()
        {
            return View();
        }
        public ActionResult _TDSSettingJsonList()
        {
            DataTable dt = comfun.fillDataTable("stpTDSSettingList", "", null);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }

        public JsonResult _TDSSettingDetailSubmit(string TDSRate, string TDSId, string FemaleAmount, string MaleAmount, string WEF)
        {
            string mes = string.Empty;
            try
            {
                DataTable table = (DataTable)JsonConvert.DeserializeObject(TDSRate, (typeof(DataTable)));

                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;

                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "stpTDSSetting_Accept_Update";
                SortedList list = new SortedList();
                com.Parameters.AddWithValue("@TDSId", TDSId);
                com.Parameters.AddWithValue("@FemaleAmount", FemaleAmount);
                com.Parameters.AddWithValue("@MaleAmount", MaleAmount);
                com.Parameters.AddWithValue("@WEF", WEF);
                com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());
                com.Parameters.AddWithValue("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy"));
                com.Parameters.AddWithValue("@SessionId", Session["SessionId"].ToString());



                // com.Parameters.AddWithValue("@fdWeeklyoffDate", attendanceDate);
                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@TDSSetting";
                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table;
                com.Parameters.Add(parameter);

                SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
                sp.Direction = ParameterDirection.Output;
                com.Parameters.Add(sp);
                com.CommandTimeout = 0;

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
        //public ActionResult ddlCustomerTypeselect()
        //{
        //    DataTable dt = comfun.fillDataTable("stpSHIFT_Select", "", null);
        //    return PartialView("_ddlCustomerTypeselect", dt);
        //} 
        public JsonResult _TDSSettingStatusUpdate()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpTDSSetting_Status", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);

        }
        public ActionResult _TDSSettingEdit()
        {
            string Id = "0";
            Id = Request.Form["TdsId"].ToString();
            SortedList list = new SortedList();
            list.Add("@Id", Id);
            DataTable dt = comfun.fillDataTable("stpTDSSetting_Edit", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _TdsSettingDetailEdit()
        {
            string Id = "0";
            Id = Request.Form["TdsId"].ToString();
            SortedList list = new SortedList();
            list.Add("@Id", Id);
            DataTable dt = comfun.fillDataTable("stpTdsDetailSettingEdit", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult TDSCalculation()
        {
            return View();
        }
        public JsonResult _TDSCalculation()
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
                mes = comfun.executeNonQueryWMessage("stpEmployeeTDSCalculation", "", list).ToString();

            }
            catch (Exception ex)
            {

                mes = "Error:" + ex.Message;
            }

            return Json(mes);
        }
        public ActionResult _TDSCalculationJsonlist()
        {
            SortedList list = new SortedList();
            list.Add("@StartDate", Request.Form["StartDate"].ToString());
            list.Add("@EmpId", Request.Form["EmpId"].ToString());
            DataTable dt = comfun.fillDataTable("stpTDSCalculatedList", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public JsonResult _TdsPaidSubmit(string data)
        {
            string mes = string.Empty;


            try
            {
                DataTable table = (DataTable)JsonConvert.DeserializeObject(data, (typeof(DataTable)));


                SqlCommand com = new SqlCommand();
                connection conObj = new connection();

                com.Connection = conObj.con;

                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "stpEmployeeTDSDeduct_Accept";
                SortedList list = new SortedList();
                com.Parameters.AddWithValue("@SessionId", Session["SessionId"].ToString());



                // com.Parameters.AddWithValue("@fdWeeklyoffDate", attendanceDate);
                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@TDSAmount";
                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table;
                com.Parameters.Add(parameter);
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
                mes = "Error: " + ex.Message;

            }
            return Json(mes);
        }
        public ActionResult TrailBalanceList()
        {
            return View();
        }
        public ActionResult _ledgerjsonlist()
        {
            SortedList list = new SortedList();
            list.Add("@fdFromDate", Request.Form["StartDate"].ToString());
            list.Add("@fdToDate", Request.Form["EndDate"].ToString());
            list.Add("@fiAccountID", Request.Form["AccountId"].ToString());
            DataTable dt = comfun.fillDataTable("stpLedgerAc", "", list);
            var json1 = JsonConvert.SerializeObject(dt);
            return Json(json1);
        }
        public ActionResult EmployeeCTC()
        {
            //comfun.saveformname("Deductions", "/admin/Deductions", "Deductions", "Deductions", "Y", "Deductions", "Deductions", "List");

            return View();
        }
        public JsonResult _EmployeeCTCSubmit()
        {
            if (payfun.sessionRecreate() == "expires")
            {
                return Json("Session expires");
            }
            string mes = string.Empty;
            try
            {

                SortedList list = new SortedList();
                list.Add("@FromEmpID", Request.Form["FromEmpId"].ToString());
                list.Add("@ToEmpID", Request.Form["ToEmpId"].ToString());
                list.Add("@AsonDate", Request.Form["AsonDate"].ToString());

                mes = comfun.executeNonQueryWMessage("stpEmpUpdate_ctcFromAemp", "", list).ToString();
            }
            catch (Exception ex)
            {

                mes = comfun.errorMessage("Payroll _EmployeeCTCSubmit", ex.Message);
            }
            return Json(mes);
        }
        //public ActionResult FreezeAttendance()
        //{
        //    SortedList list = new SortedList();
        //    list.Add("@PageNo", 1);
        //    list.Add("@PageSize", 1000);
        //    DataTable dt = comfun.fillDataTable("Department_Select", "", list);
        //    return View(dt);
        //}
        //public JsonResult _FreezeAttendanceList()
        //{
        //    // comfun.saveformname("_HelpdeskGroupMappingList", "/admin/_HelpdeskGroupMappingList", "Master/Help Desk/Mapping&nbsp;&nbsp;", "HelpDesk Mapping List", "N", "HelpDeskMapping", "HelpDeskMapping", "List", 4);
        //    SortedList list = new SortedList();
        //    string fromDate = Request.Form["StartDate"].ToString();
        //    int startYear = Convert.ToDateTime(fromDate).Year;
        //    int startMonth = Convert.ToDateTime(fromDate).Month;

        //    DateTime startDate = new DateTime(startYear, startMonth, 1);
        //    DateTime endDate = startDate.AddMonths(1).AddDays(-1);
        //    list.Add("@StartDate", startDate.ToString("dd/MMM/yyyy"));
        //    list.Add("@EndDate", endDate.ToString("dd/MMM/yyyy"));
        //    list.Add("@BranchID", Request.Form["BranchId"].ToString());
        //    list.Add("@DepartmentID", Request.Form["DepartmentId"].ToString());
        //    list.Add("@CircleId", Request.Form["CircleId"].ToString());
        //    list.Add("@CategoryId", Request.Form["CategoryId"].ToString());
        //    list.Add("@MainCategoryId", Request.Form["MainCategoryId"].ToString());
        //    DataTable dt = comfun.fillDataTable("stpEmployeeAttendance_MonthlyLockList", "", list);
        //    //return PartialView("_HelpdeskGroupMappingList", dt);
        //    var json = JsonConvert.SerializeObject(dt);
        //    return Json(json);
        //}
        public ActionResult _Employeeddl()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@FromDate", Request.Form["FromDate"].ToString());
            dt = comfun.fillDataTable("stpEmployeeFreezeddl", null, list);
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

        //public ActionResult _FreezeEmployee()
        //{
        //    SortedList list = new SortedList();
        //    //string fromDate = Request.Form["FromDate"].ToString();
        //    string AsonDate = Convert.ToDateTime(Request.Form["FromDate"].ToString()).ToString("dd/MMM/yyyy");
        //    int startYear = Convert.ToDateTime(AsonDate).Year;
        //    int startMonth = Convert.ToDateTime(AsonDate).Month;
        //    string Type = Request.Form["Type"].ToString();
        //    DateTime startDate = new DateTime(startYear, startMonth, 1);
        //    list.Add("@StartDate", startDate.ToString("dd/MMM/yyyy"));
        //    //DateTime endDate = startDate.AddMonths(1).AddDays(-1);

        //    list.Add("@EndDate", AsonDate /*endDate.ToString("dd/MMM/yyyy")*/);
        //    list.Add("@DepartmentId", Request.Form["DepartmentId"].ToString());
        //    list.Add("@BranchId", Request.Form["BranchId"].ToString());
        //    list.Add("@EmpId", Request.Form["EmpId"].ToString());
        //    list.Add("@CategoryId", Request.Form["CategoryId"].ToString());
        //    list.Add("@MainCategoryId", Request.Form["MainCategoryId"].ToString());
        //    list.Add("@CircleId", Request.Form["CircleId"].ToString());
        //    list.Add("@DesignationId", Request.Form["DesignationId"].ToString());
        //    list.Add("@LoginID", Session["EmpId"].ToString());
        //    DataTable dt = new DataTable();
        //    if (Type == "Freeze")
        //    {
        //        dt = comfun.fillDataTable("stpEmployeeAttendance_MonthlyLock", "", list);

        //    }
        //    if (Type == "UnFreeze")
        //    {
        //        dt = comfun.fillDataTable("stpEmployeeAttendance_MonthlyUnlockList", "", list);

        //    }
        //    return PartialView("_Freezeinput", dt);
        //    //var json = JsonConvert.SerializeObject(dt);
        //    //return Json(json);
        //}

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

        //public JsonResult _FreezeSaveUpdateSubmit(string FromDate, string Type, string Emps)
        //{
        //    // comfun.saveformname("_HelpdeskMappingSaveUpdateSubmit", "/admin/_HelpdeskMappingSaveUpdateSubmit", "Master/Help Desk/Mapping&nbsp;&nbsp;", "HelpDesk Mapping Add", "N", "HelpDeskMapping", "HelpDeskMapping", "Add", 2);



        //    string mes = string.Empty;
        //    try
        //    {

        //        string fromDate = FromDate;
        //        int startYear = Convert.ToDateTime(fromDate).Year;
        //        int startMonth = Convert.ToDateTime(fromDate).Month;

        //        DateTime startDate = new DateTime(startYear, startMonth, 1);
        //        //DateTime endDate = startDate.AddMonths(1).AddDays(-1);



        //        dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(Emps);
        //        //dynamic jsonData1 = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(SaleOrderVendor);
        //        DataTable table1 = new DataTable();
        //        table1.Columns.Add("EmpId", typeof(string));
        //        table1.Columns.Add("PayDays", typeof(string));



        //        foreach (var Item1 in jsonData)
        //        {
        //            DataRow dr1 = table1.NewRow();

        //            dr1["EmpId"] = Item1.EmpId;
        //            dr1["PayDays"] = Item1.PayDays;

        //            table1.Rows.Add(dr1);
        //        }





        //        SqlCommand com = new SqlCommand();
        //        connection conObj = new connection();
        //        com.Connection = conObj.con;
        //        com.CommandType = CommandType.StoredProcedure;
        //        com.CommandText = "EmployeeAttendance_FreezeNewSubmit";
        //        SqlParameter parameter = new SqlParameter();
        //        com.Parameters.AddWithValue("@MonthYearId", startMonth.ToString() + startYear.ToString());
        //        com.Parameters.AddWithValue("@LockDate", FromDate /*endDate.ToString("dd-MMM-yyyy")*/);
        //        com.Parameters.AddWithValue("@CreatedDate", comfun.dateISTstr());
        //        com.Parameters.AddWithValue("@CreatedBy", Session["EmpId"].ToString());
        //        com.Parameters.AddWithValue("@Type", Type);
        //        //  com.Parameters.AddWithValue("@EmpIdexclude", Request.Form["Emps1"].ToString());
        //        //com.Parameters.AddWithValue("@VenderId", VenderId);
        //        parameter.ParameterName = "@EmpId";
        //        parameter.SqlDbType = System.Data.SqlDbType.Structured;
        //        parameter.Value = table1;
        //        com.Parameters.Add(parameter);
        //        SqlParameter sp = new SqlParameter("@Mes", SqlDbType.VarChar, 8000);
        //        sp.Direction = ParameterDirection.Output;
        //        com.Parameters.Add(sp);
        //        com.CommandTimeout = 0;

        //        if (conObj.con.State == ConnectionState.Closed)
        //        {
        //            conObj.con.Open();
        //            com.ExecuteNonQuery();
        //            conObj.con.Close();
        //        }
        //        mes = sp.Value.ToString();
        //    }
        //    catch (Exception ex)
        //    {
        //        mes = comfun.errorMessage("_FreezeSaveUpdateSubmit", "Error: " + ex.Message);
        //    }



        //    return Json(mes);




        //}


        public ActionResult FreezeAttendance()
        {
            SortedList list = new SortedList();
            list.Add("@PageNo", 1);
            list.Add("@PageSize", 1000);
            DataTable dt = comfun.fillDataTable("Department_Select", "", list);
            return View(dt);
        }
        public JsonResult _FreezeAttendanceList()
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
            return PartialView("_Freezeinput", dt);
            //var json = JsonConvert.SerializeObject(dt);
            //return Json(json);
        }

        public JsonResult _FreezeSaveUpdateSubmit(string FromDate, string Type, string Emps)
        {
            // comfun.saveformname("_HelpdeskMappingSaveUpdateSubmit", "/admin/_HelpdeskMappingSaveUpdateSubmit", "Master/Help Desk/Mapping&nbsp;&nbsp;", "HelpDesk Mapping Add", "N", "HelpDeskMapping", "HelpDeskMapping", "Add", 2);



            string mes = string.Empty;
            try
            {

                string fromDate = FromDate;
                int startYear = Convert.ToDateTime(fromDate).Year;
                int startMonth = Convert.ToDateTime(fromDate).Month;

                DateTime startDate = new DateTime(startYear, startMonth, 1);
                DateTime endDate = startDate.AddMonths(1).AddDays(-1);



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
                com.Parameters.AddWithValue("@LockDate", endDate.ToString("dd-MMM-yyyy"));
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
                mes = comfun.errorMessage("_FreezeSaveUpdateSubmit", "Error: " + ex.Message);
            }



            return Json(mes);




        }



        public JsonResult _UnFreezeSaveUpdateSubmit(string FromDate, string Type, string Emps)
        {
            // comfun.saveformname("_HelpdeskMappingSaveUpdateSubmit", "/admin/_HelpdeskMappingSaveUpdateSubmit", "Master/Help Desk/Mapping&nbsp;&nbsp;", "HelpDesk Mapping Add", "N", "HelpDeskMapping", "HelpDeskMapping", "Add", 2);

            string mes = string.Empty;
            try
            {

                string fromDate = FromDate;
                int startYear = Convert.ToDateTime(fromDate).Year;
                int startMonth = Convert.ToDateTime(fromDate).Month;

                DateTime startDate = new DateTime(startYear, startMonth, 1);
                DateTime endDate = startDate.AddMonths(1).AddDays(-1);



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
                com.Parameters.AddWithValue("@LockDate", endDate.ToString("dd-MMM-yyyy"));
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
                mes = comfun.errorMessage("_FreezeSaveUpdateSubmit", "Error: " + ex.Message);
            }



            return Json(mes);

        }
        public ActionResult UnlockSalary()
        {
            return View();
        }
        public ActionResult _SalaryList()
        {
            DateTime monthYear = Convert.ToDateTime(Request.Form["FromDate"]);
            SortedList list = new SortedList();

            //if (Request.Form["FromDate"] != null)
            //{
            //    list.Add("@StartDate", Request.Form["FromDate"].ToString());
            //}
            //if (Request.Form["ToDate"] != null)
            //{
            //    list.Add("@EndDate", Convert.ToDateTime(Request.Form["ToDate"]).ToString("dd-MMM-yyyy"));
            //}
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
            if (Request.Form["CircleId"] != "")
            {
                list.Add("@CircleID", Request.Form["CircleId"].ToString());
            }

            if (Request.Form["CategoryId"] != "")
            {
                list.Add("@CategoryID", Request.Form["CategoryId"].ToString());
            }
            if (Request.Form["MainCategoryId"] != "")
            {
                list.Add("@MainCategoryId", Request.Form["MainCategoryId"].ToString());
            }
            list.Add("@MonthYear", monthYear.ToString("MMyyyy"));
            list.Add("@CustomerId", "0");

            DataTable dt = comfun.fillDataTable("StpSalaryGenrate_List", "", list);
            return PartialView("_SalaryUnlock", dt);
        }



        //stpTrialBalance_Detail

        public JsonResult _SalaryUnlockSubmit(string FromDate, string Emps)
        {
            // comfun.saveformname("_HelpdeskMappingSaveUpdateSubmit", "/admin/_HelpdeskMappingSaveUpdateSubmit", "Master/Help Desk/Mapping&nbsp;&nbsp;", "HelpDesk Mapping Add", "N", "HelpDeskMapping", "HelpDeskMapping", "Add", 2);



            string mes = string.Empty;
            try
            {

                string fromDate = FromDate;
                int startYear = Convert.ToDateTime(fromDate).Year;
                int startMonth = Convert.ToDateTime(fromDate).Month;

                DateTime startDate = new DateTime(startYear, startMonth, 1);
                DateTime endDate = startDate.AddMonths(1).AddDays(-1);



                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(Emps);
                //dynamic jsonData1 = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(SaleOrderVendor);
                DataTable table1 = new DataTable();
                table1.Columns.Add("EmpID", typeof(string));

                foreach (var Item1 in jsonData)
                {
                    DataRow dr1 = table1.NewRow();
                    dr1["EmpID"] = Item1.EmpId;
                    table1.Rows.Add(dr1);
                }
                SqlCommand com = new SqlCommand();
                connection conObj = new connection();
                com.Connection = conObj.con;
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "stpSalaryGenrate_Unlock";
                SqlParameter parameter = new SqlParameter();
                com.Parameters.AddWithValue("@MonthYear", startMonth.ToString() + startYear.ToString());
                com.Parameters.AddWithValue("@LoginID", Session["EmpId"].ToString());

                //  com.Parameters.AddWithValue("@EmpIdexclude", Request.Form["Emps1"].ToString());
                //com.Parameters.AddWithValue("@VenderId", VenderId);
                parameter.ParameterName = "@SalaryGenrate";
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
                mes = comfun.errorMessage("_SalaryUnlockSubmit", "Error: " + ex.Message);
            }



            return Json(mes);




        }
        public ActionResult EmployeeActualCTC()
        {

            comfun.saveformname("EmployeeActualCTC", "/Payroll/EmployeeActualCTC", "EmployeeActualCTC", "EmployeeActualCTC  form", "Y", "", "EmployeeActualCTC", "List");
            if (Request.QueryString["key"] != null)
            {
                Session["SelectedMenu"] = comfun.decryptString(Request.QueryString["key"].ToString());
            }
            return View();
        }


        public ActionResult _EmployeeActualCTCjsonList()
        {
            //comfun.saveformname("_countryList", "/admin/_countryList", "Master/Address/HiringType", "Country  list", "N", "Country", "Country", "List", 4);
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
            DataTable dt = comfun.fillDataTable("stpEmployeeActualCTC_List", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _EmployeeActualCTCSubmit()
        {
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {

                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@RateDate", comfun.dateISTstr());
                list.Add("@EmpID ", Request.Form["EmpID"].ToString());

                list.Add("@Amount", Request.Form["Amount"].ToString());

                //  list.Add("@createdDate", comfun.dateISTstr());
                list.Add("@CreatedDate", DateTime.UtcNow.AddMinutes(330).ToString("dd-MMM-yyyy HH:mm:ss"));
                list.Add("@createdby", Session["EmpId"].ToString());

                mes = comfun.executeNonQueryWMessage("stpEmployeeActualCTC_Accept", "", list).ToString();

            }
            catch (Exception ex)
            {
                mes = "Error: " + ex.Message;
            }
            return Json(mes);
        }
        public ActionResult _EmployeeActualCTCEdit()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@Id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("stpEmployeeActualCTCEdit", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public ActionResult _EmployeeActualCTCStatusUpdate()
        {
            string mes = string.Empty;
            try
            {
                SortedList list = new SortedList();
                list.Add("@Id", Request.Form["id"].ToString());
                list.Add("@IsActive", Request.Form["IsActive"].ToString());
                mes = comfun.executeNonQueryWMessage("stpEmployeeActualCTCStatusUpdate", "", list).ToString();
            }
            catch (Exception ex)
            {
                mes = ex.Message;
            }
            return Json(mes);
        }

        public ActionResult ActualRate()
        {
            return View();

        }
        public JsonResult ActualRateList()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@ID", Request.Form["ID"].ToString());
           // list.Add("@EmpID", Request.Form["EmpID"].ToString());
            dt = comfun.fillDataTable("stpSalaryEmpRate_List", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult ActualRateEdit()        
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@id", Request.Form["Id"].ToString());
            dt = comfun.fillDataTable("stpSalaryEmpRate_Edit", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }
        public JsonResult ActualRateSubmit()
        {
            //comfun.saveformname("HiringExecutiveSubmit", "/HRNew/HiringExecutiveSubmit", "HRNew/HiringExecutive", "HiringExecutive  Add", "N", "HiringExecutive", "HiringExecutive", "Add", 2);
            //comfun.saveformname("HiringExecutiveSubmit", "/HRNew/HiringExecutiveSubmit", "HRNew/HiringExecutive", "HiringExecutive", "N", "HiringExecutive", "HiringExecutive", "Submit", 3);
            string mes = string.Empty;
            SortedList list = new SortedList();
            try
            {
                list.Add("@Id", Request.Form["Id"].ToString());
                list.Add("@EmpID", Request.Form["EmpID"].ToString());
                list.Add("@Amount", Request.Form["Amount"].ToString());
                list.Add("@RateDate", Request.Form["RateDate"].ToString());
                list.Add("@CreatedBy", Session["EmpId"]);               

                mes = comfun.executeNonQueryWMessage("stpSalaryEmpRate_Accept", "", list).ToString();
            }

            catch (Exception ex)
            {

                mes = "Error:" + ex.Message;
            }

            return Json(mes);
        }

        public JsonResult ActualRateDelete()
        {
            DataTable dt = new DataTable();
            SortedList list = new SortedList();
            list.Add("@id", Request.Form["mId"].ToString());
            String mes = comfun.executeNonQueryWMessage("stpSalaryEmpRate_Del", "", null).ToString();


            var json = JsonConvert.SerializeObject(mes);
            return Json(json);
        }


        #region Document wise report

        public ActionResult DocumentWiseReport()
        {
            return View();
        }
        public ActionResult _DocumentTypeddl()
        {
            SortedList list = new SortedList();
            //list.Add("@StatusId", Request.Form["StatusId"].ToString());
            list.Add("@Id", Request.Form["Id"]);
            DataTable dt = comfun.fillDataTable("stpDocumentcategoryddl", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult _DocumentTypelist(string DocumentTypeId, string EmpID,string DepartmentId,string DesignationId)
        {
            SortedList list = new SortedList();
            //list.Add("@StatusId", Request.Form["StatusId"].ToString());
            list.Add("@EmpID", EmpID);
            list.Add("@DesignationId", DesignationId);
            list.Add("@DepartmentId", DepartmentId);
            list.Add("@DocumentTypeId", DocumentTypeId);
            DataTable dt = comfun.fillDataTable("sp_EmpDocWise", "", list);
            //return PartialView("_SalarySheet", dt);
            var json = JsonConvert.SerializeObject(dt);
           return Json(json);
        }

        #endregion

        #region Vishu
        public ActionResult EmployeeUnPaidList()
        {
            return View();
        }

        [HttpPost]
        public JsonResult _EmployeeAllPaidSubmit(string emps,string date,string type)
        {

            string mes = string.Empty;

            try
            {
                dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(emps);

                DataTable table = new DataTable();         
                table.Columns.Add("Empid", typeof(string));
                table.Columns.Add("NetPay", typeof(string));
                table.Columns.Add("Remarks", typeof(string));

                foreach (var Item1 in jsonData)
                {
                    DataRow dr1 = table.NewRow();              
                    dr1["Empid"] = Item1.id;
                    dr1["NetPay"] = Item1.NetPay;
                    dr1["Remarks"] = Item1.Remarks;

                    table.Rows.Add(dr1);
                }
                if(type== "1")
                {
                    type = "Unpaid";
                }
                else
                {
                    type = "Paid";
                }

                SqlCommand com = new SqlCommand();
                connection conObj = new connection();

                com.Connection = conObj.con;

                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "stpUnPaidEmployee";

                SqlParameter parameter = new SqlParameter();
                SortedList list = new SortedList();
               //list.Add("@Id", Request.Form["Id"].ToString());               
                com.Parameters.AddWithValue("@monthyear", date);
                com.Parameters.AddWithValue("@type", type);
                com.Parameters.AddWithValue("@Createdby", Session["EmpId"]);

                parameter.ParameterName = "@tbltypeUnPaidEmployee";

                parameter.SqlDbType = System.Data.SqlDbType.Structured;
                parameter.Value = table;

                com.Parameters.Add(parameter);

                SqlParameter sp = new SqlParameter("@mes", SqlDbType.VarChar, 8000);

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
                mes = comfun.errorMessage("_EmployeeAllPaidSubmit", "Error: " + ex.Message);
            }
            return Json(mes);
        }
        public ActionResult _EmployeeAllList()
        {
            SortedList list = new SortedList();
            //string MonthyearId = Convert.ToDateTime(Request.Form["Monthyear"].ToString()).ToString("myyyy");
            list.Add("@Monthyear", Request.Form["Monthyear"].ToString());
            list.Add("@Empid",Session["Empid"].ToString());
            DataTable dt = comfun.fillDataTable("stpPaidEmployeeList", "", list);
            return PartialView("_EmployeeAllList", dt);
        }

        public ActionResult _EmployeeUnPaidList(string Month , string Year)
        {
            SortedList list = new SortedList();

            DateTime dtDate = new DateTime(2000, Convert.ToInt32(Month), 1);
            string MonthName = dtDate.ToString("MMM");
            var MonthYear = MonthName + '-' + Year;
            //string MonthyearId = Convert.ToDateTime(Request.Form["Monthyear"].ToString()).ToString("myyyy");
            list.Add("@monthyear", MonthYear);
            list.Add("@Empid", Session["Empid"].ToString());
            DataTable dt = comfun.fillDataTable("stpUnPaidEmployee_List", "", list);
            return PartialView("_EmployeeUnPaidList", dt);
        }
        #endregion

        public ActionResult LoanApply()
        {
            DataTable dt = comfun.fillDataTable("stpLoanTypeddl", "", null);
            return View(dt);
        }

        public JsonResult _LoanApplySubmit(string IssueDate, string EmployeeName, int LoantypeId, string Amount, string Noofinstallments, string InstallmentAmount, string Deductiondate, string PayCharges)
        {
            string mes = string.Empty;
            SortedList list = new SortedList();
            list.Add("@IssueDate", Convert.ToDateTime(IssueDate).ToString("dd-MMM-yyyy"));
            list.Add("@EmployeeName", EmployeeName);
            list.Add("@LoanTypeId", LoantypeId);
            list.Add("@Amount", Amount);
            list.Add("@Noofinstallments", Noofinstallments);
            list.Add("@InstallmentAmount", InstallmentAmount);
            list.Add("@DeductionDate", Convert.ToDateTime(Deductiondate).ToString("dd-MMM-yyyy"));
            list.Add("@PayCharges", PayCharges);
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            mes = comfun.executeNonQueryWMessage("stpLoan_Accept", "", list).ToString();
            return Json(mes);
        }

        public ActionResult _PayChargeddl()
        {
            DataTable dt = comfun.fillDataTable("stpHtisPayAllChargeddl", "", null);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public JsonResult _LoanApplyList()
        {
            SortedList list = new SortedList();
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            DataTable dt = comfun.fillDataTable("stpLoanApplylist", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public ActionResult LoanApproval()
        {
            return View();
        }

        public JsonResult _LoanApprovalList()
        {
            SortedList list = new SortedList();
            list.Add("@CreatedBy", Session["EmpId"].ToString());
            DataTable dt = comfun.fillDataTable("stpLoanApprovallist", "", list);
            var json = JsonConvert.SerializeObject(dt);
            return Json(json);
        }

        public JsonResult _LoanApprovalSubmit(int Id , string Status , string Remarks)
        {
            SortedList list = new SortedList();
            string mes = string.Empty;
            list.Add("@Id", Id);
            list.Add("@Status", Status);
            list.Add("@Remarks", Remarks);
            mes = comfun.executeNonQueryWMessage("stpLoanApprovalSubmit", "", list).ToString();
            return Json(mes);
        }
    }
}

