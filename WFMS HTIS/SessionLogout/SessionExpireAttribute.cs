using System;
using System.Web;
using System.Web.Mvc;

namespace Payroll.portal.SessionLogout
{
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Method)]
    public class SessionExpireAttribute : ActionFilterAttribute
    {
        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            if (HttpContext.Current.Session["Empid"] == null)
            {
                // filterContext.Result = new RedirectResult("~/"); 
                filterContext.Result = new PartialViewResult
                {
                    // Specify the path to your partial view
                    ViewName = "~/Views/Shared/_SessionExpired.cshtml"

                };
            }
            base.OnActionExecuting(filterContext);
        }
    }
}
