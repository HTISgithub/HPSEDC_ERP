<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Report1.aspx.cs" Inherits="Payroll.portal.Views.Shared.Report1" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=12.0.0.0, Culture=neutral, PublicKeyToken=89845dcd8080cc91" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<%


%>
<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script runat="server">
        commonFunctions comfun = new commonFunctions();
        void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {

                ReportViewer1.LocalReport.ReportPath = Server.MapPath("~/Reports/Report1.rdlc");
                ReportViewer1.LocalReport.DataSources.Clear();
                System.Data.DataTable dt = comfun.fillDataTable("", "select * from tblViksatPayEmployees", null);
                ReportDataSource rd = new ReportDataSource("DataSet1", dt);
                ReportViewer1.LocalReport.DataSources.Add(rd);
                ReportViewer1.LocalReport.Refresh();

            }
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server"></asp:ScriptManager>
        <div>
            <rsweb:ReportViewer ID="ReportViewer1" runat="server"></rsweb:ReportViewer>
        </div>
    </form>
</body>
</html>
