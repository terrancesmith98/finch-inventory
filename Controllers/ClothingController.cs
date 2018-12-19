using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using System.Net;
using System.Web;
using System.Web.Mvc;
using Finch_Inventory.Models;

namespace Finch_Inventory.Controllers
{
    public class ClothingController : Controller
    {
        private FinchDbContext db = new FinchDbContext();

        // GET: Clothing
        public async Task<ActionResult> Index()
        {
            var clothing = db.Clothing.Include(c => c.Location).Include(c => c.Position).Include(c => c.Status).Include(c => c.Type);
            return View(await clothing.ToListAsync());
        }

        // GET: Clothing/Details/5
        public async Task<ActionResult> Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Clothing clothing = await db.Clothing.FindAsync(id);
            if (clothing == null)
            {
                return HttpNotFound();
            }
            return View(clothing);
        }

        // GET: Clothing/Create
        public ActionResult Create()
        {
            ViewBag.LocationID = new SelectList(db.Location, "ID", "Location1");
            ViewBag.PositionID = new SelectList(db.Position, "ID", "Position1");
            ViewBag.StatusID = new SelectList(db.Status, "ID", "Status1");
            ViewBag.TypeID = new SelectList(db.Type, "ID", "Type1");
            return View();
        }

        // POST: Clothing/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Create([Bind(Include = "ID,PM_Number,PositionID,Manufacturer,TypeID,Serial_Number,Date_Received,Date_Placed_On_Mac,Date_Removed_From_Mac,StatusID,LocationID,Comments")] Clothing clothing)
        {
            if (ModelState.IsValid)
            {
                db.Clothing.Add(clothing);
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }

            ViewBag.LocationID = new SelectList(db.Location, "ID", "Location1", clothing.LocationID);
            ViewBag.PositionID = new SelectList(db.Position, "ID", "Position1", clothing.PositionID);
            ViewBag.StatusID = new SelectList(db.Status, "ID", "Status1", clothing.StatusID);
            ViewBag.TypeID = new SelectList(db.Type, "ID", "Type1", clothing.TypeID);
            return View(clothing);
        }

        // GET: Clothing/Edit/5
        public async Task<ActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Clothing clothing = await db.Clothing.FindAsync(id);
            if (clothing == null)
            {
                return HttpNotFound();
            }
            ViewBag.LocationID = new SelectList(db.Location, "ID", "Location1", clothing.LocationID);
            ViewBag.PositionID = new SelectList(db.Position, "ID", "Position1", clothing.PositionID);
            ViewBag.StatusID = new SelectList(db.Status, "ID", "Status1", clothing.StatusID);
            ViewBag.TypeID = new SelectList(db.Type, "ID", "Type1", clothing.TypeID);
            ViewBag.Machines = db.Machines.ToList();

            return View(clothing);
        }

        // POST: Clothing/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Edit([Bind(Include = "ID,PM_Number,PositionID,Manufacturer,TypeID,Serial_Number,Date_Received,Date_Placed_On_Mac,Date_Removed_From_Mac,StatusID,LocationID,Comments")] Clothing clothing)
        {
            if (ModelState.IsValid)
            {
                db.Entry(clothing).State = EntityState.Modified;
                await db.SaveChangesAsync();
                return RedirectToAction("Index", "Home");
            }
            ViewBag.LocationID = new SelectList(db.Location, "ID", "Location1", clothing.LocationID);
            ViewBag.PositionID = new SelectList(db.Position, "ID", "Position1", clothing.PositionID);
            ViewBag.StatusID = new SelectList(db.Status, "ID", "Status1", clothing.StatusID);
            ViewBag.TypeID = new SelectList(db.Type, "ID", "Type1", clothing.TypeID);

            return RedirectToAction("Index", "Home");
        }

        // GET: Clothing/Delete/5
        public async Task<ActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Clothing clothing = await db.Clothing.FindAsync(id);
            if (clothing == null)
            {
                return HttpNotFound();
            }
            return View(clothing);
        }

        // POST: Clothing/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> DeleteConfirmed(int id)
        {
            Clothing clothing = await db.Clothing.FindAsync(id);
            db.Clothing.Remove(clothing);
            await db.SaveChangesAsync();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
