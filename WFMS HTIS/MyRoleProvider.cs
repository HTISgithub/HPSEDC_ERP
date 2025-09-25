using System;
using System.Collections.Generic;
using System.Linq;

using System.Web;
using System.Web.Caching;
using System.Web.Security;
using System.Collections;
using System.Data;

namespace Payroll.portal
{
    public class MyRoleProvider : RoleProvider
    {
        private int _cacheTimeoutInMinute = 20;
        commonFunctions comfun = new commonFunctions();
        public override string ApplicationName
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public override void AddUsersToRoles(string[] usernames, string[] roleNames)
        {
            throw new NotImplementedException();
        }

        public override void CreateRole(string roleName)
        {
            throw new NotImplementedException();
        }

        public override bool DeleteRole(string roleName, bool throwOnPopulatedRole)
        {
            throw new NotImplementedException();
        }

        public override string[] FindUsersInRole(string roleName, string usernameToMatch)
        {
            throw new NotImplementedException();
        }

        public override string[] GetAllRoles()
        {
            var cacheKey = string.Format("{0}_AllRole", "AllRoles");
            if (HttpRuntime.Cache[cacheKey] != null)
            {
                return (string[])HttpRuntime.Cache[cacheKey];
            }
            using (var MyDatabaseEntities = new PayrollEntities())
            {
                string[] roles = new string[] { };
                roles = MyDatabaseEntities.tbUserRoleMasters.Select(r => r.RoleName).ToArray();
                if (roles.Count() > 0)
                {    
                    HttpRuntime.Cache.Insert(cacheKey, roles, null, DateTime.Now.AddMinutes(_cacheTimeoutInMinute), Cache.NoSlidingExpiration);

                }
                return roles;
            }
        }

        public override string[] GetRolesForUser(string username)
        {
            if (!HttpContext.Current.User.Identity.IsAuthenticated)
            {
                return null;
            }

            //var cacheKey = string.Format("{0}_role", username);
            //if (HttpRuntime.Cache[cacheKey] != null)
            //{
            //    return (string[])HttpRuntime.Cache[cacheKey];
            //}
            //string[] roles = new string[] { };
            //using (PayrollEntities dc = new PayrollEntities())
            //{
            //    roles = (from a in dc.tbUserRoleMasters
            //             join b in dc.tblVIKSATUserLoginPart2 on a.Id equals b.fiEmpCategoryId
            //             join c in dc.tblVIKSATUserLoginPart1 on b.EmpId equals c.fiEmployeeId 

            //             where c.fvUserName.Equals(username)
            //             select a.RoleName.Trim()).ToArray<string>();
            //    if (roles.Count() > 0)
            //    {
            //        HttpRuntime.Cache.Insert(cacheKey, roles, null, DateTime.Now.AddMinutes(_cacheTimeoutInMinute), Cache.NoSlidingExpiration); 
            //    }
            //} 
            SortedList list = new SortedList();
            list.Add("@userName", username);
            // list.Add("@Password", L.Password);
            DataTable dt = comfun.fillDataTable("stpHiringUserRole_List", "",list);
            string[] result = new string[dt.Rows.Count];
            int index = 0;
            foreach (DataRow dr in dt.Rows)
            {
                result[index] = dr[0].ToString();
                index++;
            }
            return result;
        }

        public override string[] GetUsersInRole(string roleName)
        {
            throw new NotImplementedException();
        }

        public override bool IsUserInRole(string username, string roleName)
        {
            var userRoles = GetRolesForUser(username);
            return userRoles.Contains(roleName);    
        }

        public override void RemoveUsersFromRoles(string[] usernames, string[] roleNames)
        {
            throw new NotImplementedException();
        }

        public override bool RoleExists(string roleName)
        {     
            throw new NotImplementedException();
        }
       
    }
}