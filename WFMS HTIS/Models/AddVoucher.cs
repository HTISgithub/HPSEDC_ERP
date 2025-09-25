using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Payroll.portal.Models
{
    public class AddVoucher
    {
        public int AccountId { get; set; }

        public int SubAccountId { get; set; }


        public string Narration { get; set; }
        public string DebitAmount { get; set; }
        public string CreditAmount { get; set; }
        public string TransactionType { get; set; }
        public string ReceiptChequeNo { get; set; }
        public string ReceiptChequeDate { get; set; }
        public int CostTypeId { get; set; }

        public int ProjectId { get; set; }
        public int CircleId { get; set; }

    }
    public class NewAddReceiptVoucher
    {
        public string fiVoucherID { get; set; }
        public string fiSaleID { get; set; }
        public string BillNumber { get; set; }
        public string PaidAmount { get; set; }
        public string TDSAmount { get; set; }
        public string GSTAmount { get; set; }
        public string Deduction { get; set; }
        public string GSTPaymentPending { get; set; }
        public string RetentionAmt { get; set; }
    }
    public class NewAddExpensesVoucher
        {
       // public string fiExpBillDetailId { get; set; }
        public string fiExpBillID { get; set; }
        public string fiAccountID { get; set; }
        public string fvHsnCode { get; set; }
        public string fnAmount { get; set; }
        public string fnDiscount { get; set; }
         public string fnNetAmount { get; set; }
        public string fnIGSTRate { get; set; }
        public string fnCGSTRate { get; set; }
        public string fnSGSTRate { get; set; }
        public string fnIGSTAmount { get; set; }
        public string fnCGSTAmount { get; set; }
        public string fnSGSTAmount { get; set; }
        public string fiCostcenter { get; set; }
        public string fiCircleID { get; set; }
        public string fiCostypeID { get; set; }
        public string fiSaleteamID { get; set; }
        public string SaleTeam { get; set; }
        public string CostPoint { get; set; }
        public string SiteId { get; set; }


    }
    public class ExpenseVoucherEntryDetail
    {
        // public string fiExpBillDetailId { get; set; }
        
             public string ID { get; set; }
        public string fiEmployeeID { get; set; }
        public string EmployeeCode { get; set; }
        public string EmployeeName { get; set; }
        public string ClaimDate { get; set; }
        public string ClaimSubmit { get; set; }
        public string ApprovedByPM { get; set; }
        public string RejectedByPM { get; set; }
        public string ApprovedByPMO { get; set; }
        public string RejectedByPMO { get; set; }
        public string PendingByPMO { get; set; }
        public string Remarks { get; set; }
        public string ProjectID { get; set; }
        public string SiteID { get; set; }
        public string CircleID { get; set; }

    }
    public class ExpenseVoucherEntry
    {
        // public string fiExpBillDetailId { get; set; }
     
        public string fiEmployeeID { get; set; }
        public string fvEmployeeCode { get; set; }
        public string fvEmployeeName { get; set; }
        public string ClaimAmount { get; set; }
        public string PMApproved { get; set; }
        public string PMOApproved { get; set; }
        public string Payable { get; set; }
       
    }
}