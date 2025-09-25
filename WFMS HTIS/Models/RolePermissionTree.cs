using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Payroll.portal.Models
{
    //public class RolePermissionTree
    //{
    //    public string id { get; set; }
    //    public string parentid { get; set; }
    //    public string Mainmenu { get; set; }
    //    public string SubMenu { get; set; }
    //    public FlatObject(string name, Int64 id, Int64 parentId)
    //    {
    //        data = name;
    //        Id = id;
    //        ParentId = parentId;
    //    }
    //}
    public class FlatObject
    {
        public int Id { get; set; }
        public int ParentId { get; set; }
        public string data { get; set; }
        public FlatObject(string name, int id, int parentId)
        {
            data = name;
            Id = id;
            ParentId = parentId;
        }
    }
    public class RecursiveObject
    {
        public string data { get; set; }
        public int id { get; set; }
        public FlatTreeAttribute attr { get; set; }
        public List<RecursiveObject> children { get; set; }
    }
    public class FlatTreeAttribute
    {
        public string id;
        public bool selected;
    }

    public class Forms
    {
        public string FormName { get; set; }
        public string FormPath{ get; set; }
        public string Title { get; set; }
    }


}