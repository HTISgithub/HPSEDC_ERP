using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Payroll.portal
{
    public class customAuthorize : AuthorizeAttribute
    {
        commonFunctions comfun = new commonFunctions();
        public override void OnAuthorization(AuthorizationContext filterContext)
        {
            // If they are authorized, handle accordingly
            if (this.AuthorizeCore(filterContext.HttpContext))
            { 
                SortedList list = new SortedList();
                list.Add("@UserName", filterContext.HttpContext.User.Identity.Name);
                list.Add("@ActivityName", filterContext.ActionDescriptor.ActionName);
                object hasPermission = comfun.executeScaler("[UserPermissionCheck]", "", list);
                if (hasPermission == null)
                {
                    filterContext.Controller.TempData["Error"] = "Permission not defined";
                    filterContext.Result = new RedirectResult("~/Home/Unauthorized");
                }
                else
                if (hasPermission.ToString() != "Y")
                {
                    filterContext.Result = new RedirectResult("~/Home/Unauthorized");
                } 
                base.OnAuthorization(filterContext);
            }
            else
            {
                // Otherwise redirect to your specific authorized area
                filterContext.Result = new RedirectResult("~/Home/Unauthorized");
            }
        }
    }
}