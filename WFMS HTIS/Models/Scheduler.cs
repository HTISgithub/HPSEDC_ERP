using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Data;
using System.Collections;
using System.Net;
using System.Net.Mail;

namespace Payroll.portal.Models
{
    public class Secheduler
    {
        private readonly object _lock = new object();
        private static readonly Secheduler _MailSecheduler = new Secheduler();
        string connectionString = ConfigurationManager.ConnectionStrings["cnLocalHost"].ConnectionString;

        public static Secheduler Instance
        {
            get
            {
                return _MailSecheduler;
            }
        }

        #region Auto Assign Leave

        private DataTable GetUnAssignedEmployee()
        {
            using (SqlConnection conn = new SqlConnection(connectionString))

            //Create the SqlConnection object
            using (SqlCommand cmd = new SqlCommand("sp_GetEmployeeNotAssignedLeave", conn))
            {
                //Create the SqlCommand object by passing the stored procedure name and connection object as parameters
                SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                adapt.SelectCommand.CommandType = CommandType.StoredProcedure;
                DataTable dt = new DataTable();
                adapt.Fill(dt);

                return dt;
            }
        }
        public DataTable getLeaveType(string fiEmployeeID)
        {
            DataTable dt = new DataTable();
            using (SqlConnection conn = new SqlConnection(connectionString))

            //Create the SqlConnection object
            using (SqlCommand cmd = new SqlCommand("stpLeaveAssignedEmployeewise", conn))
            {
                cmd.Parameters.Add("@EmpId", SqlDbType.Int).Value = fiEmployeeID;
                //Create the SqlCommand object by passing the stored procedure name and connection object as parameters
                SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                adapt.SelectCommand.CommandType = CommandType.StoredProcedure;
                adapt.Fill(dt);
            }
            return dt;
        }
        public void AssignLeaveAftersixMonths()
        {
            DataTable dt = GetUnAssignedEmployee();
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    DataTable dtLeaveType = getLeaveType(row["fiEmployeeID"].ToString());
                    foreach (DataRow item in dtLeaveType.Rows)
                    {
                        SqlConnection con = new SqlConnection(connectionString);
                        SqlCommand cmd = new SqlCommand();
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandText = "stpLeaveAssigned_Accept";
                        cmd.Parameters.Add("@Id", SqlDbType.Int).Value = row["fiEmployeeID"].ToString();
                        cmd.Parameters.Add("@LeaveType", SqlDbType.Int).Value = item["Id"].ToString();
                        cmd.Parameters.Add("@FromDate", SqlDbType.DateTime).Value = DateTime.Now.ToString("yyyy-MM-dd");
                        cmd.Parameters.Add("@ToDate", SqlDbType.DateTime).Value = DateTime.Now.Year + "-12-31";
                        cmd.Parameters.Add("@Approver", SqlDbType.VarChar).Value = row["fiReportingManager"].ToString();
                        cmd.Parameters.Add("@Remarks", SqlDbType.VarChar).Value = "Auto update after six month";
                        cmd.Parameters.Add("@Day", SqlDbType.VarChar).Value = "F";
                        cmd.Parameters.Add("@mes", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output;

                        cmd.Connection = con;
                        try
                        {
                            con.Open();
                            cmd.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                        finally
                        {
                            con.Close();
                            con.Dispose();
                        }
                    }
                }
            }
        }
        #endregion

        #region Auto Notification Probation Mail

        private DataTable GetEmployeeNearExpireProbation()
        {
            using (SqlConnection conn = new SqlConnection(connectionString))

            //Create the SqlConnection object
            using (SqlCommand cmd = new SqlCommand("StpSendandsaveProbationPeriod", conn))
            {
                //Create the SqlCommand object by passing the stored procedure name and connection object as parameters
                SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                adapt.SelectCommand.CommandType = CommandType.StoredProcedure;
                DataTable dt = new DataTable();
                adapt.Fill(dt);

                return dt;
            }
        }
        private string SendEmail(string displayName, string toEmail, string subject, string message)
        {
            string m = "";
            try
            {
                string email = "Wfms@horizontelecom.in";
                string password = "wfms@9002"; ;
                displayName = "WFMS";

                var loginInfo = new NetworkCredential(email, password);
                var msg = new MailMessage();
                var smtpClient = new SmtpClient("smtp.rediffmailpro.com");
                //var smtpClient = new SmtpClient("smtp.gmail.com");
                smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtpClient.Port = 587;

                msg.From = new MailAddress("hrms@wfms.in", displayName);
                msg.To.Add(new MailAddress(toEmail));
                //msg.To.Add(new MailAddress("jagvirjb@gmail.com"));
                //msg.Bcc.Add(new MailAddress("htis.comp111@gmail.com"));
                //msg.Bcc.Add(new MailAddress("htis.comp111@gmail.com"));
                //msg.CC.Add(new MailAddress("info@horizontelecom.in"));
                //msg.CC.Add(new MailAddress("anish.dhiman@horizontelecom.in"));
                HtmlString htmlString = new HtmlString(message);
                //       AlternateView htmlView =
                //AlternateView.CreateAlternateViewFromString(message, Encoding.UTF8, "text/html");
                //       msg.AlternateViews.Add(htmlView); // And a html attachment to make sure.

                msg.Subject = subject;
                msg.Body = message;
                msg.ReplyTo = new MailAddress("hrms@wfms.in");
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
        public void ProbationMailAftersixMonths()
        {
            DataTable dt = GetEmployeeNearExpireProbation();
            if (dt.Rows.Count > 0)
            {
                SqlConnection conn;
                SqlCommand comm;
                string toEmail = "info@horizontelecom.in"; //info@horizontelecom.in
                foreach (DataRow row in dt.Rows)
                {
                    string mes1 = string.Empty;
                    string mes = string.Empty;
                    int i = 0;
                    string body = "<table>" +
                        "<tr>" +
                            "<td style=\"color:#5a78ad !important;\"> Dear HR,<br /><br /></td>" +
                        "</tr>" +
                        "<tr>" +
                            "<td style=\"color:#5a78ad !important;\">" +
                                 "Hope this email finds you well. This is to remind you about the upcoming probation period review for "+ dt.Rows[i]["fvEmployeeName"].ToString() + "("+ dt.Rows[i]["fvEmployeeCode"].ToString() + "), who joined our team on "+ dt.Rows[i]["fdDateOfJoining"].ToString() + ". As per our company's policy, the probation period is set to conclude today.<br /><br />" +
                            "</td>" +
                        "</tr>" +
                        "<tr>" +
                            "<td style=\"color:#5a78ad !important;\">" +
                                "<b> Thanks</b>" +
                            "</td>" +
                        "</tr>" +
                    "</table>";
                    mes1 = string.Empty;
                    mes = string.Empty;

                    string subject = "Probation Period Notification";

                    mes1 = SendEmail(subject, toEmail, subject, body);
                    // update in database after sent mail
                    if (mes1 == "email sent")
                    {
                        conn = new SqlConnection(connectionString);
                        conn.Open();
                        comm = new SqlCommand("update tblVIKSATpayEmployees set sendProvisionMail=1 where fiEmployeeID=" + Convert.ToInt32(dt.Rows[i]["fiEmployeeID"]) + "", conn);
                        try
                        {
                            comm.ExecuteNonQuery();
                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            conn.Close();
                        }
                    }
                    mes1 = mes + " and " + mes1;
                    i++;
                }
            }
        }
        #endregion
    }
}