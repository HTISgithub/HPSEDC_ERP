using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CrystalDecisions.Web; 
using System.Collections;
using CrystalDecisions.ReportAppServer;
using System.Data;
using System.Data.SqlClient;
using System.Device.Location;
using System.Drawing;
using System.Globalization;
using System.IO;

using System.Configuration;
using System.Net.Mail;
using CrystalDecisions.Shared;
using System.Net;
using Newtonsoft.Json;

namespace Payroll.portal.CrystalReports
{
   
   
    public partial class PIprint : System.Web.UI.Page
    {

        commonFunctions comfun = new commonFunctions();
        protected void Page_Load(object sender, EventArgs e)
        {
            string Id = string.Empty;
            string Method = string.Empty;
            string MonthYear = string.Empty;
            string Asondate = string.Empty;
            string DesignationId = string.Empty;
            string CategoryId = string.Empty;
            string DepartmentId = string.Empty;
            string BranchId = string.Empty;
           
            DataSet ds = new DataSet();
            
            //CrystalReportViewer1.Dispose();
            if (Request.QueryString["id"].ToString() != null)
            {
                Id = Request.QueryString["id"].ToString();
            }
           
            if (Request.QueryString["Method"].ToString() != null)
            {
           Method = Request.QueryString["Method"].ToString();
            }
            
                //stpIndentDetail_print
                SortedList list = new SortedList();
            if (Method == "PI")
            {
                list.Add("@PIID", Id);
                ds = comfun.fillDataSet("stpPerformaInvoice_Print", null, list);
                ds.Tables[0].TableName = "ReportHead";
                ds.Tables[1].TableName = "PIMaster";
                ds.Tables[2].TableName = "PIDetail";
                ds.Tables[3].TableName = "PITerms";
            }
            if(Method == "IP")
            {
                list.Add("@Id", Id);
                ds = comfun.fillDataSet("stpIndentDetail_Print", null, list);
                ds.Tables[0].TableName = "ReportHead";
                ds.Tables[1].TableName = "IndentPart1";
                ds.Tables[2].TableName = "IndentPart2";
                
            }
            if (Method == "Invoice")
            {
                list.Add("@InvoiceID", Id);
                ds = comfun.fillDataSet("stpDepartmentInvoice_Ptint", null, list);
                ds.Tables[0].TableName = "ReportHead";
                ds.Tables[1].TableName = "CustomerInv";
                ds.Tables[2].TableName = "InvDetail";

            }
           

            if (Method == "CustomerPrint")
            {
                list.Add("@InvoiceID", Id);
                ds = comfun.fillDataSet("StpBillDetailExcel", null, list);
                ds.Tables[0].TableName = "INVEmpDetail";
               
            }
            if (Method == "CoveringLetter")
            {
                list.Add("@InvoiceID", Id);
                ds = comfun.fillDataSet("stpCustomerInviceCoveringLetter", null, list);
                ds.Tables[0].TableName = "CoveringLetter";

            }
            
            // ds.Tables[1].TableName = "Data";
            //CrystalReportViewer1.ToolPanelView = CrystalDecisions.Web.ToolPanelViewType.None;
           
            //if (Method == "PI")
            //{
            //    CrystalReport1 CrystalReport1 = new CrystalReport1();
            //    CrystalReport1.SetDataSource(ds);
            //    CrystalReportViewer1.ReportSource = CrystalReport1;
            //}
            if (Method == "Invoice")
            {
                CustomerINV CustomerINV = new CustomerINV();
                CustomerINV.SetDataSource(ds);
                CrystalReportViewer1.ReportSource = CustomerINV;
                using (MailMessage mm = new MailMessage("kuldeep.viksat@gmail.com", "kuldeep.viksat@gmail.com"))
                {
                    mm.Subject = "Crystal Report PDF";
                    mm.Body = "Attachment: Customer's Crystal Report PDF";
                    mm.CC.Add(new MailAddress("munish.kumar@horizontelecom.in"));
                    mm.Attachments.Add(new Attachment(CustomerINV.ExportToStream(ExportFormatType.PortableDocFormat), "Crystal.pdf"));
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
                    }
                }
            }
            if (Method == "StudentInfo")
            {


                DataTable table = (DataTable)JsonConvert.DeserializeObject(Id, (typeof(DataTable)));

                for (int i = 0; i <= table.Rows.Count - 1; i++)
                {
                   
                     Id = table.Rows[i]["Id"].ToString();

                    //  Response.Redirect("~/CrystalReports/PIprint.aspx?Id=" + ids + "&&" + "Method=" + "StudentInfo");

                    list.Clear();
                    list.Add("@Studentid", Id);
                ds = comfun.fillDataSet("stpStudentInfo_Print", null, list);
                ds.Tables[0].TableName = "StudentInfo";
                StudentInfo StudentInfo = new StudentInfo();
                StudentInfo.SetDataSource(ds);
                CrystalReportViewer1.ReportSource = StudentInfo;
                using (MailMessage mm = new MailMessage("Wfms@horizontelecom.in", ds.Tables[0].Rows[0]["Email"].ToString()))
                {
                    mm.Subject = "Laptop Form";
                    mm.Body = "Attachment: Student Laptop Receipt Form";
                    mm.Bcc.Add(new MailAddress("munish.kumar@horizontelecom.in"));
                    mm.Attachments.Add(new Attachment(StudentInfo.ExportToStream(ExportFormatType.PortableDocFormat), "laptopform.pdf"));
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
                    }
                    }

                }
            }
            if (Method == "CustomerPrint")
            {
                InvoiceEmpList invoiceEmpList = new InvoiceEmpList();
                invoiceEmpList.SetDataSource(ds);  
                CrystalReportViewer1.ReportSource = invoiceEmpList;
            }
             if (Method == "CoveringLetter")
            {
                ManpowerCoveringLetter manpowerCoveringLetter = new ManpowerCoveringLetter();
                manpowerCoveringLetter.SetDataSource(ds);
                CrystalReportViewer1.ReportSource = manpowerCoveringLetter;
            }

            //if (Method == "IP")
            //{
            //    CrystalIndent crystalIndent = new CrystalIndent();
            //    crystalIndent.SetDataSource(ds);
            //    CrystalReportViewer1.ReportSource = crystalIndent;
            //}


            if (Method == "EmpSalarySlip")
            {
                MonthYear = Request.QueryString["MonthYear"].ToString();
                list.Clear();
                list.Add("@EmpID", Id);
                list.Add("@MonthYear", MonthYear);
                ds = comfun.fillDataSet("stp_EmpSalarySlip_Print", null, list);
                ds.Tables[0].TableName = "SalarySlip";
                ds.Tables[1].TableName = "Particulars";
                EmpSalarySlip EmpSalarySlip = new EmpSalarySlip();
                EmpSalarySlip.SetDataSource(ds);
                CrystalReportViewer1.ReportSource = EmpSalarySlip;
            }
            if (Method == "SaleOrderPrint")
            {
                //MonthYear = Request.QueryString["MonthYear"].ToString();
                list.Clear();
                list.Add("@InvoiceID", Id);
                ds = comfun.fillDataSet("stpAccountingSaleInvoice_Print", null, list);
                ds.Tables[0].TableName = "Companyinfo";
                ds.Tables[1].TableName = "Customerinfo";
                ds.Tables[2].TableName = "SaleBillDetail";
                rptSaleOrder rptSaleOrder = new rptSaleOrder();
                rptSaleOrder.SetDataSource(ds);
                rptSaleOrder.SetDatabaseLogon("sa", "data@5_htis#9002");
                CrystalReportViewer1.ReportSource = rptSaleOrder;
            }
            if (Method == "PurchaseOrderPrint")
            {
                list.Clear();
                list.Add("@Id", Id);
                ds = comfun.fillDataSet("stpPurchaseOrder_Print", null, list);
                ds.Tables[0].TableName = "Reporthead";
                ds.Tables[1].TableName = "PurchaseOrder";
                ds.Tables[2].TableName = "PurchaseOrderDetail";
                ds.Tables[3].TableName = "PurchaseOrderTerms";
                rptPurchaseOrderPrint rptPurchaseOrderPrint = new rptPurchaseOrderPrint();
                rptPurchaseOrderPrint.SetDataSource(ds);
                rptPurchaseOrderPrint.SetDatabaseLogon("sa", "data@5_htis#9002");
                CrystalReportViewer1.ReportSource = rptPurchaseOrderPrint;
            }
            if (Method == "SaleBillPrint")
            {
                //MonthYear = Request.QueryString["MonthYear"].ToString();
                list.Clear();
                list.Add("@InvoiceID", Id);
                ds = comfun.fillDataSet("stpAccountingSaleInvoice_Print", null, list);
                ds.Tables[0].TableName = "Companyinfo";
                ds.Tables[1].TableName = "Customerinfo";
                ds.Tables[2].TableName = "SaleBillDetail";
                rptSaleInvoice rptSaleInvoice = new rptSaleInvoice();
                rptSaleInvoice.SetDataSource(ds);
                rptSaleInvoice.SetDatabaseLogon("sa", "data@5_htis#9002");
                 CrystalReportViewer1.ReportSource = rptSaleInvoice;
               
            }


            if (Method == "OVSheet")
            {
                //MonthYear = Request.QueryString["MonthYear"].ToString();
                list.Clear();
                list.Add("@InvoiceID", Id);
                ds = comfun.fillDataSet("stpCreditNote_Print", null, list);
                ds.Tables[0].TableName = "Companyinfo";
                ds.Tables[1].TableName = "CreditNote1";
                ds.Tables[2].TableName = "CreditNote2";
                rptCreditNote rptCreditNote = new rptCreditNote();
                rptCreditNote.SetDataSource(ds);
                rptCreditNote.SetDatabaseLogon("sa", "data@5_htis#9002");
                CrystalReportViewer1.ReportSource = rptCreditNote;

            }

            if (Method == "CreditNotePrint")
            {
                //MonthYear = Request.QueryString["MonthYear"].ToString();
                list.Clear();
                list.Add("@InvoiceID", Id);
                ds = comfun.fillDataSet("stpCreditNote_Print", null, list);
                ds.Tables[0].TableName = "Companyinfo";
                ds.Tables[1].TableName = "CreditNote1";
                ds.Tables[2].TableName = "CreditNote2";
                //rptCreditNote rptCreditNote = new rptCreditNote();
                //rptCreditNote.SetDataSource(ds);
                //rptCreditNote.SetDatabaseLogon("sa", "data@5_htis#9002");
                //CrystalReportViewer1.ReportSource = rptCreditNote;

            }


            if (Method == "PurchaseBill")
            {
                //MonthYear = Request.QueryString["MonthYear"].ToString();
                list.Clear();
                list.Add("@InvoiceID", Id);
                ds = comfun.fillDataSet("stpAccountingSaleInvoice_Print", null, list);
                ds.Tables[0].TableName = "Companyinfo";
                ds.Tables[1].TableName = "Customerinfo";
                ds.Tables[2].TableName = "SaleBillDetail";
                rptPurchaseInvoice rptPurchaseInvoice = new rptPurchaseInvoice();
                rptPurchaseInvoice.SetDataSource(ds);
                CrystalReportViewer1.ReportSource = rptPurchaseInvoice;
                //rptPurchaseInvoice rptPurchaseInvoice = new rptPurchaseInvoice();
                //rptPurchaseInvoice.SetDataSource(ds);
                //CrystalReportViewer1.ReportSource = rptPurchaseInvoice;
            }
            if (Method == "Accountledger")
            {
                //MonthYear = Request.QueryString["MonthYear"].ToString();
                list.Clear();
                list.Add("@fiAccountID", Id);
                list.Add("@fdFromDate", Request.QueryString["FromDate"].ToString());
                list.Add("@fiSessionId", Session["SessionId"].ToString());
                list.Add("@fdToDate", Request.QueryString["ToDate"].ToString());
                ds = comfun.fillDataSet("stpAccountLedgerPrint", null, list);
                ds.Tables[0].TableName = "ReportHead";
                ds.Tables[1].TableName = "Accountinfo";
                ds.Tables[2].TableName = "LedgerAC";
                rptLedgerAc rptLedgerAc = new rptLedgerAc();
                rptLedgerAc.SetDataSource(ds);
                rptLedgerAc.SetDatabaseLogon("sa", "data@5_htis#9002"); 
                CrystalReportViewer1.ReportSource = rptLedgerAc;
                   
                //AccountLedger AccountLedger = new AccountLedger();
                //AccountLedger.SetDataSource(ds);
                //CrystalReportViewer1.ReportSource = AccountLedger;
            }


            if (Method == "Payroll")
            {
                Asondate = Request.QueryString["Asondate"].ToString();
                DepartmentId = Request.QueryString["DepartmentId"].ToString();
                DesignationId = Request.QueryString["DesignationId"].ToString();
                BranchId = Request.QueryString["BranchId"].ToString();
                CategoryId = Request.QueryString["CategoryId"].ToString();

                list.Clear();
                list.Add("@AsonDate" , "01/Jun/2023");
                list.Add("@EmpId" , Id);
                list.Add("@DesignationId" , 0);
                list.Add("@CategoryId" , 0);
                list.Add("@DepartmentId" , 0);
                list.Add("@BranchId", 0);
                ds = comfun.fillDataSet("StpSalarySheetHudeiPrint", null, list);
                ds.Tables[0].TableName = "ReportHead";
                ds.Tables[1].TableName = "PayRoll";
                ds.Tables[2].TableName = "PayRollSumm";
                rptPayRollPrint rptPayRollPrint = new rptPayRollPrint();
                rptPayRollPrint.SetDataSource(ds);
                rptPayRollPrint.SetDatabaseLogon("sa", "data@5_htis#9002");
                CrystalReportViewer1.ReportSource = rptPayRollPrint;

               
            }

            CrystalReportViewer1.Visible = true;
            CrystalReportViewer1.RefreshReport();
            //CrystalReportViewer1 = null;
           
            //myDataReport.Close(); // I can't remember if this is part of the reportDucment class
            //myDataReport.Dispose();
            //myDataReport = null;
            //CrystalReportViewer1.Dispose();

            GC.Collect();
            GC.WaitForPendingFinalizers();
            

        }
    }
}