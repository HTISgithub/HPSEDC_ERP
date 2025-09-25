using ClosedXML.Excel;
using Microsoft.Reporting.WebForms;
using System;
using System.Collections;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using Newtonsoft.Json;
using System.Collections.Generic;
//using System.Text;
using Payroll.portal.Models;
using static Payroll.portal.Controllers.AdminController;
using System.Data.SqlClient;
//using System.Reflection.Emit;
using Excelcon = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Runtime.InteropServices;
namespace Payroll.portal
{
    public partial class RdlcFromate : System.Web.UI.Page
    {
        commonFunctions comfun = new commonFunctions();
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                _getData();
            }
           

        }
        public void _getData()
        {
            string Id = string.Empty;
          
            if (Request.QueryString["Id"] != null)
            {
                Id = Request.QueryString["Id"].ToString();
            }

            //   List<Customer> customers = null;
            DataSet ds = new DataSet();
            SortedList list = new SortedList();
                list.Add("@Id",Id);
            ds = comfun.fillDataSet("stpCustomerSaleInvoice_Print", null,list);
           // ds = ds.Tables[0]["SaleInvoice"];
          
            ReportViewer1.LocalReport.ReportPath = Server.MapPath("~/rdlcReports/RDLC/Report1.rdlc");
            ReportViewer1.LocalReport.DataSources.Clear();
            ReportDataSource rdc = new ReportDataSource("SaleInvoice", ds.Tables[0]);
            ReportViewer1.LocalReport.DataSources.Add(rdc);
            ReportViewer1.LocalReport.Refresh();
            ReportViewer1.DataBind();

        }
    }
}