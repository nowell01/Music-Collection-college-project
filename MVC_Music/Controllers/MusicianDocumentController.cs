using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using MVC_Music.CustomControllers;
using MVC_Music.Data;
using MVC_Music.Models;
using MVC_Music.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MVC_Music.Controllers
{
    [Authorize(Roles = "Admin, Supervisor, Staff")]
    public class MusicianDocumentController : ElephantController
    {
        private readonly MusicContext _context;

        public MusicianDocumentController(MusicContext context)
        {
            _context = context;
        }

        // GET: MusicianDocument
        public async Task<IActionResult> Index(int? MusicianID, string SearchString, int? page, int? pageSizeID)
        {
            //Clear the sort/filter/paging URL Cookie for Controller
            CookieHelper.CookieSet(HttpContext, ControllerName() + "URL", "", -1);

            PopulateDropDownLists();

            //Toggle the Open/Closed state of the collapse depending on if we are filtering
            ViewData["Filtering"] = "btn-outline-secondary"; //Asume not filtering
            //Then in each "test" for filtering, add ViewData["Filtering"] = "btn-danger" if true;
            int numberFilters = 0;

            var musicianDocuments = _context.MusicianDocuments
                .Include(m => m.Musician)
                .OrderBy(m => m.FileName)
                .AsNoTracking();

            if (MusicianID.HasValue)
            {
                musicianDocuments = musicianDocuments.Where(p => p.MusicianID == MusicianID);
                numberFilters++;
            }
            if (!String.IsNullOrEmpty(SearchString))
            {
                musicianDocuments = musicianDocuments.Where(p => p.FileName.ToUpper().Contains(SearchString.ToUpper()));
                numberFilters++;
            }
            //Give feedback about the state of the filters
            if (numberFilters != 0)
            {
                //Toggle the Open/Closed state of the collapse depending on if we are filtering
                ViewData["Filtering"] = " btn-danger";
                //Show how many filters have been applied
                ViewData["numberFilters"] = "(" + numberFilters.ToString()
                    + " Filter" + (numberFilters > 1 ? "s" : "") + " Applied)";
                //Keep the Bootstrap collapse open
                @ViewData["ShowFilter"] = " show";
            }

            int pageSize = PageSizeHelper.SetPageSize(HttpContext, pageSizeID, ControllerName());
            ViewData["pageSizeID"] = PageSizeHelper.PageSizeList(pageSize);
            var pagedData = await PaginatedList<MusicianDocument>.CreateAsync(musicianDocuments.AsNoTracking(), page ?? 1, pageSize);

            return View(pagedData);
        }

        // GET: MusicianDocument/Edit/5
        [Authorize(Roles = "Admin, Supervisor")]
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null || _context.MusicianDocuments == null)
            {
                return NotFound();
            }

            var musicianDocument = await _context.MusicianDocuments
                .Include(m => m.Musician)
                .AsNoTracking()
                .FirstOrDefaultAsync(m => m.ID == id);
            if (musicianDocument == null)
            {
                return NotFound();
            }
            PopulateDropDownLists(musicianDocument);
            return View(musicianDocument);
        }

        // POST: MusicianDocument/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize(Roles = "Admin, Supervisor")]
        public async Task<IActionResult> Edit(int id)
        {
            var musicianDocumentToUpdate = await _context.MusicianDocuments
                .Include(m => m.Musician)
                .FirstOrDefaultAsync(m => m.ID == id);

            //Check that you got it or exit with a not found error
            if (musicianDocumentToUpdate == null)
            {
                return NotFound();
            }

            if (await TryUpdateModelAsync<MusicianDocument>(musicianDocumentToUpdate, "",
                d => d.FileName, d => d.Description))
            {
                try
                {
                    await _context.SaveChangesAsync();
                    return Redirect(ViewData["returnURL"].ToString());
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!MusicianDocumentExists(musicianDocumentToUpdate.ID))
                    {
                        return NotFound();
                    }
                    else
                    {
                        throw;
                    }
                }
                catch (DbUpdateException)
                {
                    ModelState.AddModelError("", "Unable to save the update. Try again, and if the problem persists see your system administrator.");
                }
            }
            PopulateDropDownLists(musicianDocumentToUpdate);
            return View(musicianDocumentToUpdate);
        }

        // GET: MusicianDocument/Delete/5
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null || _context.MusicianDocuments == null)
            {
                return NotFound();
            }

            var musicianDocument = await _context.MusicianDocuments
                .Include(m => m.Musician)
                .FirstOrDefaultAsync(m => m.ID == id);

            if (musicianDocument == null)
            {
                return NotFound();
            }

            return View(musicianDocument);
        }

        // POST: MusicianDocument/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {

            if (_context.MusicianDocuments == null)
            {
                return Problem("Entity set 'MusicContext.MusicianDocuments'  is null.");
            }
            var musicianDocument = await _context.MusicianDocuments.FindAsync(id);
            try
            {
                if (musicianDocument != null)
                {
                    _context.MusicianDocuments.Remove(musicianDocument);
                }

                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateException)
            {
                ModelState.AddModelError("", "Unable to save the update. Try again, and if the problem persists see your system administrator.");
            }
            return View(musicianDocument);
        }

        public async Task<FileContentResult> Download(int id)
        {
            var theFile = await _context.UploadedFiles
                .Include(d => d.FileContent)
                .Where(f => f.ID == id)
                .FirstOrDefaultAsync();

            if (theFile?.FileContent?.Content == null || theFile.MimeType == null)
            {
                return new FileContentResult(Array.Empty<byte>(), "application/octet-stream");
            }

            return File(theFile.FileContent.Content, theFile.MimeType, theFile.FileName);
        }

        private SelectList MusicianSelectList(int? id)
        {
            var dQuery = from d in _context.Musicians
                         orderby d.LastName, d.FirstName
                         select d;
            return new SelectList(dQuery, "ID", "Summary", id);
        }
        private void PopulateDropDownLists(MusicianDocument musicianDocument = null)
        {
            ViewData["MusicianID"] = MusicianSelectList(musicianDocument?.MusicianID);
        }

        private bool MusicianDocumentExists(int id)
        {
            return _context.MusicianDocuments.Any(e => e.ID == id);
        }
    }
}
