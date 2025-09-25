using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Globalization;

namespace Payroll.portal.classes
{
    public static class Dropdowns
    {
        public static string dropdown(DataTable dt,string val,string txt )
        {
            string ddl = string.Empty;
            for (int i = 0; i < dt.Rows.Count; i++)
            {

                ddl += "<option value='" + dt.Rows[i][val] + "'>" + CultureInfo.CurrentCulture.TextInfo.ToTitleCase(dt.Rows[i][txt].ToString().ToLower()) + "</option>";
            }
            return ddl;
        }
    }
}