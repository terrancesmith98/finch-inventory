using CodeNode.ActiveDirectory;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Finch_Inventory.Controllers
{
    public class BaseController : Controller
    {
        // GET: Base
        public ActionResult Index()
        {
            var manager = new ActiveDirectoryManager();
            var currWinUser = manager.GetCurrentWindowUser();
            return View();
        }
    }
}