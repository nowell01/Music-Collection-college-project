using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using MVC_Music.CustomControllers;
using MVC_Music.Data;
using MVC_Music.Models;
using MVC_Music.Utilities;
using MVC_Music.ViewModels;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;

namespace MVC_Music.Controllers
{
    [Authorize(Roles = "Admin, Supervisor, Staff")]
    public class PerformancesController : ElephantController
    {
        private readonly MusicContext _context;

        public PerformancesController(MusicContext context)
        {
            _context = context;
        }

        // GET: Performances
        public async Task<IActionResult> Index()
        {
            var musicContext = _context.Performances.Include(p => p.Instrument).Include(p => p.Musician).Include(p => p.Song);
            return View(await musicContext.ToListAsync());
        }

        // GET: Performances/Details/5
        [Authorize(Roles = "Admin, Supervisor")]
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var performance = await _context.Performances
                .Include(p => p.Instrument)
                .Include(p => p.Musician)
                .Include(p => p.Song)
                .FirstOrDefaultAsync(m => m.ID == id);
            if (performance == null)
            {
                return NotFound();
            }

            return View(performance);
        }

        // GET: Performances/Create
        [Authorize(Roles = "Admin, Supervisor")]
        public IActionResult Create()
        {
            ViewData["InstrumentID"] = new SelectList(_context.Instruments, "ID", "Name");
            ViewData["MusicianID"] = new SelectList(_context.Musicians, "ID", "FirstName");
            ViewData["SongID"] = new SelectList(_context.Songs, "ID", "Title");
            return View();
        }

        // POST: Performances/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize(Roles = "Admin, Supervisor")]
        public async Task<IActionResult> Create([Bind("ID,Comments,FeePaid,SongID,MusicianID,InstrumentID")] Performance performance)
        {
            if (ModelState.IsValid)
            {
                _context.Add(performance);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }
            ViewData["InstrumentID"] = new SelectList(_context.Instruments, "ID", "Name", performance.InstrumentID);
            ViewData["MusicianID"] = new SelectList(_context.Musicians, "ID", "FirstName", performance.MusicianID);
            ViewData["SongID"] = new SelectList(_context.Songs, "ID", "Title", performance.SongID);
            return View(performance);
        }

        // GET: Performances/Edit/5
        [Authorize(Roles = "Admin, Supervisor")]
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var performance = await _context.Performances.FindAsync(id);
            if (performance == null)
            {
                return NotFound();
            }
            ViewData["InstrumentID"] = new SelectList(_context.Instruments, "ID", "Name", performance.InstrumentID);
            ViewData["MusicianID"] = new SelectList(_context.Musicians, "ID", "FirstName", performance.MusicianID);
            ViewData["SongID"] = new SelectList(_context.Songs, "ID", "Title", performance.SongID);
            return View(performance);
        }

        // POST: Performances/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize(Roles = "Admin, Supervisor")]
        public async Task<IActionResult> Edit(int id, [Bind("ID,Comments,FeePaid,SongID,MusicianID,InstrumentID")] Performance performance)
        {
            if (id != performance.ID)
            {
                return NotFound();
            }

            if (ModelState.IsValid)
            {
                try
                {
                    _context.Update(performance);
                    await _context.SaveChangesAsync();
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!PerformanceExists(performance.ID))
                    {
                        return NotFound();
                    }
                    else
                    {
                        throw;
                    }
                }
                return RedirectToAction(nameof(Index));
            }
            ViewData["InstrumentID"] = new SelectList(_context.Instruments, "ID", "Name", performance.InstrumentID);
            ViewData["MusicianID"] = new SelectList(_context.Musicians, "ID", "FirstName", performance.MusicianID);
            ViewData["SongID"] = new SelectList(_context.Songs, "ID", "Title", performance.SongID);
            return View(performance);
        }

        // GET: Performances/Delete/5
        [Authorize(Roles = "Admin, Supervisor")]
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var performance = await _context.Performances
                .Include(p => p.Instrument)
                .Include(p => p.Musician)
                .Include(p => p.Song)
                .FirstOrDefaultAsync(m => m.ID == id);
            if (performance == null)
            {
                return NotFound();
            }

            return View(performance);
        }

        // POST: Performances/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        [Authorize(Roles = "Admin, Supervisor")]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var performance = await _context.Performances.FindAsync(id);
            if (performance != null)
            {
                _context.Performances.Remove(performance);
            }

            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }
        public async Task<IActionResult> PerformanceReport(int? page, int? pageSizeID)
        {
            var sumQ = GetPerformanceReports();

            int pageSize = PageSizeHelper.SetPageSize(HttpContext, pageSizeID, "PerformanceReportVM");
            ViewData["pageSizeID"] = PageSizeHelper.PageSizeList(pageSize);
            var pagedData = await PaginatedList<PerformanceReportVM>.CreateAsync(sumQ.AsNoTracking(), page ?? 1, pageSize);

            return View(pagedData);
        }
        public IActionResult DownloadPerformances()
        {

            //Get the performances
            var perf = from p in _context.Performances
                        .Include(p => p.Musician)
                        .AsEnumerable()
                        .GroupBy(p => p.Musician.Summary)
                       select new
                       {
                           Musician = p.Key,
                           AverageFee = p.Average(p => p.FeePaid),
                           HighestFee = p.Max(p => p.FeePaid),
                           LowestFee = p.Min(p => p.FeePaid),
                           TotalPerformances = p.Count(),
                       };
            int numRows = perf.Count();

            if (numRows > 0)
            {
                using (ExcelPackage excel = new ExcelPackage())
                {
                    var workSheet = excel.Workbook.Worksheets.Add("Performances");

                    workSheet.Cells[3, 1].LoadFromCollection(perf, true);

                    //Style fee column for currency
                    workSheet.Column(2).Style.Numberformat.Format = "###,##0.00";
                    workSheet.Column(3).Style.Numberformat.Format = "###,##0.00";
                    workSheet.Column(4).Style.Numberformat.Format = "###,##0.00";

                    workSheet.Cells[4, 1, numRows + 3, 1].Style.Font.Bold = true;

                    using (ExcelRange totalfees = workSheet.Cells[numRows + 4, 5])
                    {
                        totalfees.Formula = "Sum(" + workSheet.Cells[4, 5].Address + ":" + workSheet.Cells[numRows + 3, 5].Address + ")";
                        totalfees.Style.Font.Bold = true;
                        //Set backgroung color for the total
                        var fill = totalfees.Style.Fill;
                        fill.PatternType = ExcelFillStyle.Solid;
                        fill.BackgroundColor.SetColor(Color.LightPink);
                    }

                    using (ExcelRange totalfees = workSheet.Cells[numRows + 4, 1])
                    {
                        totalfees.Formula = "Counta(" + workSheet.Cells[4, 1].Address + ":" + workSheet.Cells[numRows + 3, 1].Address + ")";
                        totalfees.Style.Font.Bold = true;
                        //Set backgroung color for the total
                        var fill = totalfees.Style.Fill;
                        fill.PatternType = ExcelFillStyle.Solid;
                        fill.BackgroundColor.SetColor(Color.LightPink);
                    }

                    //Set Style and backgound colour of headings
                    using (ExcelRange headings = workSheet.Cells[3, 1, 3, 5])
                    {
                        headings.Style.Font.Bold = true;
                        var fill = headings.Style.Fill;
                        fill.PatternType = ExcelFillStyle.Solid;
                        fill.BackgroundColor.SetColor(Color.LightCoral);
                    }

                    workSheet.Cells.AutoFitColumns();

                    workSheet.Cells[1, 1].Value = "Performance Report";
                    using (ExcelRange Rng = workSheet.Cells[1, 1, 1, 5])
                    {
                        Rng.Merge = true;
                        Rng.Style.Font.Bold = true;
                        Rng.Style.Font.Size = 18;
                        Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                    DateTime utcDate = DateTime.UtcNow;
                    TimeZoneInfo esTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
                    DateTime localDate = TimeZoneInfo.ConvertTimeFromUtc(utcDate, esTimeZone);
                    using (ExcelRange Rng = workSheet.Cells[2, 5])
                    {
                        Rng.Value = "Created: " + localDate.ToShortTimeString() + " on " +
                            localDate.ToShortDateString();
                        Rng.Style.Font.Bold = true;
                        Rng.Style.Font.Size = 12;
                        Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    }

                    try
                    {
                        Byte[] theData = excel.GetAsByteArray();
                        string filename = "Performances.xlsx";
                        string mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                        return File(theData, mimeType, filename);
                    }
                    catch (Exception)
                    {
                        return BadRequest("Could not build and download the file.");
                    }
                }
            }
            return NotFound("No data.");
        }

        private IQueryable<PerformanceReportVM> GetPerformanceReports()
        {
            return _context.PerformanceReports.AsNoTracking();
        }
        private bool PerformanceExists(int id)
        {
            return _context.Performances.Any(e => e.ID == id);
        }
    }
}
