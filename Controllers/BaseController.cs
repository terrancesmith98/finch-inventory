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
            var inventory = db.Clothing.ToList();
#if (!DEBUG)
            var manager = new ActiveDirectoryManager();
            var currWinUser = manager.GetCurrentWindowUser();
#endif
#if (DEBUG)
            var currWinUser = "tsmith@otiservices.com";
#endif
            ViewBag.CurrUser = currWinUser;
            ViewBag.Inventory = inventory;
        }
    }
}