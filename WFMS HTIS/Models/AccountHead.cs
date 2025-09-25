using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Payroll.portal.Models
{
    public class AccountHead
    {
        public string HeadId { get; set; }
        public string HeadName { get; set; }
    }
    public class ReqPostModal
    {
        public int RequisitionDetailId { get; set; }
        public int NoOfPost { get; set; }
        public string[] DeploymentOffice { get; set; }
        public string[] EmpType { get; set; }

        public string[] Designation { get; set; }

        public string[] Tenure { get; set; }

        public string[] Renumeration { get; set; }

        public string[] Rate { get; set; }

        public string[] AVOName { get; set; }

        public string[] AVODesignation { get; set; }

        public string[] AVOMobile { get; set; }

        public string[] AVOEmail { get; set; }
    }

    public class LedgerAc
    {
        public string SrNo { get; set; }
        public string Voucherdate { get; set; }
        public string RefNo { get; set; }
        public string Particulars { get; set; }
        public string Debit { get; set; }
        public string Credit { get; set; }
        public string Balance { get; set; }
        public string BalType { get; set; }

    }
}