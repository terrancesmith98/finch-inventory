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
            Document document = Documents.InventoryAudit(inventoryItems);
            //create PDF Renderer
            PdfDocumentRenderer renderer = new PdfDocumentRenderer(true);
            renderer.Document = document;
            renderer.RenderDocument();
            var fileName = "Finch_Inventory_Audit.pdf";
            var filePath = Path.Combine(Server.MapPath("~/Content/Reports"), fileName);
            renderer.PdfDocument.Save(filePath);

            return fileName;
        }

        public string CreateWeeklyPMReport()
        {
            var currentClothing = db.Clothings.Where(c => c.StatusID == 2 && c.RollTypeID == 2).ToList();
            var currentRolls = db.Clothings.Where(c => c.StatusID == 2 && c.RollTypeID == 1).ToList();
            //create Migradoc Document
            Document document = Documents.WeeklyPMReport(currentClothing, currentRolls);
            PageSetup pageSetup = document.DefaultPageSetup.Clone();
            // set orientation
            pageSetup.Orientation = Orientation.Landscape;
            //create PDF renderer
            PdfDocumentRenderer renderer = new PdfDocumentRenderer(true);
            renderer.Document = document;
            renderer.RenderDocument();
            var fileName = "Finch_Weekly_PM_Report.pdf";
            var filePath = Path.Combine(Server.MapPath("~/Content/Reports"), fileName);
            renderer.PdfDocument.Save(filePath);

            return fileName;
        }
    }
}