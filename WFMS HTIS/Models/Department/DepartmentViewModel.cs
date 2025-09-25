using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Payroll.portal.Models.Department
{
    public class AdditionalChargeDetailDTO
    {
        public string departmentId { get; set; }
        public string departmentName { get; set; }
        public string desigId { get; set; }
        public string desigName { get; set; }
        public string officeTypeId { get; set; }
        public string officeTypeName { get; set; }
        public string officeId { get; set; }
        public string officeName { get; set; }
        public string chargeAssignedOn { get; set; }
    }

    public class Root
    {
        public string empId { get; set; }
        public string employeeName { get; set; }
        public string applicationId { get; set; }
        public string applicationName { get; set; }
        public string serviceId { get; set; }
        public string serviceName { get; set; }
        public string emailId { get; set; }
        public string orgId { get; set; }
        public string orgName { get; set; }
        public string departmentId { get; set; }
        public string departmentName { get; set; }
        public string desigId { get; set; }
        public string desigName { get; set; }
        public string officeTypeId { get; set; }
        public string officeTypeName { get; set; }
        public string officeId { get; set; }
        public string officeName { get; set; }
        public string postId { get; set; }
        public string postName { get; set; }
        public string roleId { get; set; }
        public string roleName { get; set; }
        public string pmisCode { get; set; }
        public string salaryCode { get; set; }
        public List<AdditionalChargeDetailDTO> additionalChargeDetailDTO { get; set; }
    }
}