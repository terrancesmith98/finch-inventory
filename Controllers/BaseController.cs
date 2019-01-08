using CodeNode.ActiveDirectory;
using Finch_Inventory.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Finch_Inventory.Controllers
{
    public class BaseController : Controller
    {
        private static FinchDbContext db = new FinchDbContext();

        protected override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            db.Dispose();
            db = new FinchDbContext();
            var inventory = db.Clothings.ToList();

            var manager = new ActiveDirectoryManager();
            var currWinUser = manager.GetCurrentWindowUser();
            if (currWinUser != null)
            {
                var currUser = db.Users.SingleOrDefault(u => u.UserName == currWinUser.UserPrincipalName);
                var roles = db.UserRoles.Where(r => r.UserID == currUser.ID).Select(r => r.RoleID).ToList();
                ViewBag.CurrUser = currUser;
                ViewBag.UserRoles = roles;

            }


            
            
            ViewBag.RolesList = new MultiSelectList(db.Roles.OrderBy(r => r.ID).Select(r => r.Role1).ToList());
            ViewBag.Inventory = inventory;
        }
    }
}