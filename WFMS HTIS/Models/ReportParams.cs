using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.IO;
using System.Data.SqlClient;
namespace Payroll.portal.Models
{
    public class ReportParams
    {
        public string RptfileName { get; set; }
        public string ReportTitle { get; set; }
        public string ReportType { get; set; }
       public  DataTable DataSource { get; set; }
        public bool IsHasparams { get; set; }
        public string DataSetName { get; internal set; }
    }
}
