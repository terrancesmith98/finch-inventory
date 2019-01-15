using Finch_Inventory.Models;
using System;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Web.Mvc;

namespace Finch_Inventory.Controllers
{
    public class BaseController : Controller
    {
        private static FinchDbContext db = new FinchDbContext();

        protected override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            //string UserEmail = "terry.smith@finchpaper.com";

            var UserDomain = Environment.UserDomainName.ToString();
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain);
            UserPrincipal user = UserPrincipal.FindByIdentity(ctx, User.Identity.Name);
            string UserEmail = user.EmailAddress;

            if (user.Name == "Terrance Smith")
                UserEmail = "terry.smith@finchpaper.com";

            db.Dispose();
            db = new FinchDbContext();
            var inventory = db.Clothings.ToList();

            if (UserEmail != null)
            {
                var currUser = db.Users.SingleOrDefault(u => u.UserName == UserEmail);
                if (currUser != null)
                {
                    var roles = db.UserRoles.Where(r => r.UserID == currUser.ID).Select(r => r.RoleID).ToList();
                    ViewBag.CurrUser = currUser;
                    ViewBag.UserRoles = roles;
                    ViewBag.RolesList = new MultiSelectList(db.Roles.OrderBy(r => r.ID).Select(r => r.Role1).ToList());
                    ViewBag.Admins = db.Admins.ToList(); 
                }
                else
                {
                    ViewBag.ErrorMessage = $"There is no user account found that matches the current logged in user.";
                    ViewBag.ErrorDetails = $@"The current logged in user has the email address {UserEmail}.
                                            Please contact an admin for assistance. Current admins include:";
                }
            }
            else
            {
                ViewBag.ErrorMessage = $"There is no user account found that matches the current logged in user.";
                ViewBag.ErrorDetails = $"The current logged in user has the email address {UserEmail}.";
            }

            ViewBag.Inventory = inventory;
        }
    }
}