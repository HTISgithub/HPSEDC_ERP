using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Payroll.portal.Models
{
    public class DeviceLogQueryModel
    {
        // Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(myJsonResponse); 
        public class QueryInfoList
        {
            public int QryType { get; set; }
            public int QryCondition { get; set; }
            public string QryData { get; set; }
        }

        public class RootQuery
        {
            public int Num { get; set; }
            public List<QueryInfoList> QueryInfoList { get; set; }
            public int Limit { get; set; }
            public int Offset { get; set; }
        }


    }
}