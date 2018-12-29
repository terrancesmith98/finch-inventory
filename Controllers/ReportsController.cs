using Finch_Inventory.Custom_Classes;
using Finch_Inventory.Models;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web.Mvc;

namespace Finch_Inventory.Controllers
{

    public class ReportsController : Controller
    {
        private static FinchDbContext db = new FinchDbContext();
        // GET: Reports
        public ActionResult Index()
        {
            return View();
        }


        public string CreateInventoryAuditReport()
        {
            var inventoryItems = db.Clothings.Where(c => c.StatusID != 2 && c.StatusID != 3).ToList();
            //create Migradoc Document
            Document document = Documents.WeeklyInventoryAudit(inventoryItems);
            //create PDF Renderer
            PdfDocumentRenderer renderer = new PdfDocumentRenderer(true);
            renderer.Document = document;
            renderer.RenderDocument();
            //var fileName = "Finch_Weekly_Inventory_Audit_" + DateTime.Now.ToShortDateString().Replace('/', '-').Replace(':', '_').Replace(' ', '_') + ".pdf";
            var fileName = "Finch_Weekly_Inventory_Audit_" + DateTime.Now.ToShortDateString().Replace('/', '-') + ".pdf";
            var filePath = Path.Combine(Server.MapPath("~/Content/Reports"), fileName);
            renderer.PdfDocument.Save(filePath);

            return fileName;
        }
    }
}