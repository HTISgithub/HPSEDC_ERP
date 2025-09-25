using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Payroll.portal.Models
{
    public class Graph
    {
    }

    public class DatasetDashboard
    {
        public string label { get; set; }
        public List<int> data { get; set; }
        public List<string> backgroundColor { get; set; }
    }

    public class RootDashboard
    {
        public List<string> labels { get; set; }
        public List<DatasetDashboard> datasets { get; set; }
    }

}