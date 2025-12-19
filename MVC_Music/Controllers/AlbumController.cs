using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Storage;
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
    [Authorize]
    public class AlbumController : ElephantController
    {
        private readonly MusicContext _context;

        public AlbumController(MusicContext context)
        {
            _context = context;
        }

        // GET: Album
        public async Task<IActionResult> Index(string? SearchString, int? GenreID, string? actionButton, int? page, int? pageSizeID, string sortDirection = "asc", string sortField = "Name")
        {
            string[] sortOptions = new[] { "Name", "Year Produced", "Price", "Genre" };

            ViewData["Filtering"] = "btn-outline-secondary";
            int numberFilters = 0;
            //Then in each "test" for filtering, add to the count of Filters applied

            PopulateDropDownLists();    //Genre DDL

            var albums = _context.Albums
                .Include(a => a.Genre)
                .Include(a => a.Songs)
                .AsNoTracking();

            if (GenreID.HasValue)
            {
                albums = albums.Where(a => a.GenreID == GenreID);
                numberFilters++;
            }

            if (!string.IsNullOrEmpty(SearchString))
            {
                albums = albums.Where(a => a.Name.ToUpper().Contains(SearchString.ToUpper()));
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

            //Before we sort, see if we have called for a change of filtering or sorting
            if (!String.IsNullOrEmpty(actionButton)) //Form Submitted!
            {
                page = 1;//Reset page to start

                if (sortOptions.Contains(actionButton))//Change of sort is requested
                {
                    if (actionButton == sortField) //Reverse order on same field
                    {
                        sortDirection = sortDirection == "asc" ? "desc" : "asc";
                    }
                    sortField = actionButton;//Sort by the button clicked
                }
            }

            //Now we know which field and direction to sort by
            if (sortField == "Name")
            {
                if (sortDirection == "asc")
                {
                    albums = albums
                        .OrderBy(a => a.Name);
                }
                else
                {
                    albums = albums
                        .OrderByDescending(a => a.Name);
                }
            }
            if (sortField == "Year Produced")
            {
                if (sortDirection == "asc")
                {
                    albums = albums.OrderBy(a => a.YearProduced);
                }
                else
                {
                    albums = albums.OrderByDescending(a => a.YearProduced);
                }
            }
            if (sortField == "Price")
            {
                if (sortDirection == "asc")
                {
                    albums = albums.OrderBy(a => a.Price);
                }
                else
                {
                    albums = albums.OrderByDescending(a => a.Price);
                }
            }
            if (sortField == "Genre")
            {
                if (sortDirection == "asc")
                {
                    albums = albums.OrderBy(a => a.Genre.Name);
                }
                else
                {
                    albums = albums.OrderByDescending(a => a.Genre.Name);
                }
            }

            //Set sort for next time
            ViewData["sortField"] = sortField;
            ViewData["sortDirection"] = sortDirection;

            //Handle Paging
            int pageSize = PageSizeHelper.SetPageSize(HttpContext, pageSizeID, ControllerName());
            ViewData["pageSizeID"] = PageSizeHelper.PageSizeList(pageSize);
            var pagedData = await PaginatedList<Album>.CreateAsync(albums.AsNoTracking(), page ?? 1, pageSize);

            return View(pagedData);
        }

        // GET: Album/Details/5
        [Authorize(Roles = "Admin, Supervisor, Staff")]
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var album = await _context.Albums
                .Include(a => a.Genre)
                .AsNoTracking()
                .FirstOrDefaultAsync(m => m.ID == id);

            if (album == null)
            {
                return NotFound();
            }

            return View(album);
        }

        // GET: Album/Create
        [Authorize(Roles = "Admin, Supervisor, Staff")]
        public IActionResult Create()
        {
            var album = new Album();
            PopulateDropDownLists(album);
            return View(album);
        }

        // POST: Album/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize(Roles = "Admin, Supervisor, Staff")]
        public async Task<IActionResult> Create([Bind("ID,Name,YearProduced,Price,GenreID")] Album album)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    _context.Add(album);
                    await _context.SaveChangesAsync();
                    var returnUrl = ViewData["returnURL"]?.ToString();
                    if (string.IsNullOrEmpty(returnUrl))
                    {
                        return RedirectToAction(nameof(Index));
                    }
                    return Redirect(returnUrl);
                }
            }
            catch (RetryLimitExceededException /* dex */)
            {
                ModelState.AddModelError("", "Unable to save changes after multiple attempts. Try again, and if the problem persists, see your system administrator.");
            }
            catch (DbUpdateException)
            {
                ModelState.AddModelError("", "Unable to save changes. Try again, and if the problem persists see your system administrator.");
            }
            PopulateDropDownLists(album);
            return View(album);
        }

        // GET: Album/Edit/5
        [Authorize(Roles = "Admin, Supervisor, Staff")]
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var album = await _context.Albums
                .Include(a => a.Genre)
                .AsNoTracking()
                .FirstOrDefaultAsync(m => m.ID == id);

            if (album == null)
            {
                return NotFound();
            }
            PopulateDropDownLists(album);
            return View(album);
        }

        // POST: Album/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize(Roles = "Admin, Supervisor, Staff")]
        public async Task<IActionResult> Edit(int id, Byte[] RowVersion)
        {
            var albumToUpdate = await _context.Albums
                .Include(a => a.Songs).ThenInclude(a => a.Genre)
                .FirstOrDefaultAsync(m => m.ID == id);
            if (albumToUpdate == null)
            {
                return NotFound();
            }

            //Put the original RowVersion value in the OriginalValues collection for the entity
            _context.Entry(albumToUpdate).Property("RowVersion").OriginalValue = RowVersion;

            if (await TryUpdateModelAsync<Album>(albumToUpdate, "",
                d => d.Name, d => d.YearProduced, d => d.Price, d => d.GenreID))
            {
                try
                {
                    _context.Update(albumToUpdate);
                    await _context.SaveChangesAsync();
                    return RedirectToAction(nameof(Index));
                }
                catch (DbUpdateConcurrencyException ex)
                {
                    var exceptionEntry = ex.Entries.Single();
                    var clientValues = (Album)exceptionEntry.Entity;
                    var databaseEntry = exceptionEntry.GetDatabaseValues();
                    if (databaseEntry == null)
                    {
                        ModelState.AddModelError("",
                            "Unable to save changes. The Album was deleted by another user.");
                    }
                    else
                    {
                        var databaseValues = (Album)databaseEntry.ToObject();
                        if (databaseValues.Name != clientValues.Name)
                            ModelState.AddModelError("Name", "Current value: "
                                + databaseValues.Name);
                        if (databaseValues.YearProduced != clientValues.YearProduced)
                            ModelState.AddModelError("YearProduced", "Current value: "
                                + databaseValues.YearProduced);
                        if (databaseValues.Price != clientValues.Price)
                            ModelState.AddModelError("Price", "Current value: "
                                + String.Format("{0:c}", databaseValues.Price));
                        //For the foreign key, we need to go to the database to get the information to show
                        if (databaseValues.GenreID != clientValues.GenreID)
                        {
                            Genre databaseGenre = await _context.Genres.FirstAsync(i => i.ID == databaseValues.GenreID);
                            ModelState.AddModelError("GenreID", $"Current value: {databaseGenre.Name}");
                        }
                        ModelState.AddModelError(string.Empty, "The record you attempted to edit "
                                + "was modified by another user after you received your values. The "
                                + "edit operation was canceled and the current values in the database "
                                + "have been displayed. If you still want to save your version of this record, click "
                                + "the Save button again. Otherwise click the 'Back to Album List' hyperlink.");
                        albumToUpdate.RowVersion = databaseValues.RowVersion ?? Array.Empty<byte>();
                        ModelState.Remove("RowVersion");
                    }
                }
                catch (DbUpdateException)
                {
                    ModelState.AddModelError("", "Unable to save changes to the Album. Try again, and if the problem persists see your system administrator.");
                }
            }
            PopulateDropDownLists(albumToUpdate);
            return View(albumToUpdate);
        }

        // GET: Album/Delete/5
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var album = await _context.Albums
                .Include(a => a.Genre)
                .FirstOrDefaultAsync(m => m.ID == id);

            if (album == null)
            {
                return NotFound();
            }

            return View(album);
        }

        // POST: Album/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> DeleteConfirmed(int id, Byte[] RowVersion)
        {
            var album = await _context.Albums.FirstOrDefaultAsync(a => a.ID == id); ;
            try
            {
                if (album != null)
                {
                    _context.Entry(album).Property("RowVersion").OriginalValue = RowVersion;
                    _context.Albums.Remove(album);
                }
                await _context.SaveChangesAsync();

                var returnUrl = ViewData["returnURL"]?.ToString();
                if (string.IsNullOrEmpty(returnUrl))
                {
                    return RedirectToAction(nameof(Index));
                }
                return Redirect(returnUrl);
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "The record you attempted to delete "
                    + "has been modified by another user. Please go back and refresh.");
                ViewData["CannotSave"] = "disabled='disabled'";
            }
            catch (DbUpdateException dex)
            {
                if (dex.GetBaseException().Message.ToLower().Contains("foreign key constraint failed"))
                {
                    ModelState.AddModelError("", "Unable to delete Album. An Album that is linked to a Song cannot be deleted.");
                }
                else
                {
                    ModelState.AddModelError("", "Unable to save Changes. Try again, if the problem persists try contacting with the system administration");
                }
            }
            return View(album);
        }

        private SelectList GenreList(int? selectedId)
        {
            return new SelectList(_context.Genres
                .OrderBy(a => a.Name), "ID", "Name", selectedId);
        }

        private void PopulateDropDownLists(Album? album = null)
        {
            ViewData["GenreID"] = GenreList(album?.GenreID);
        }

        private bool AlbumExists(int id)
        {
            return _context.Albums.Any(e => e.ID == id);
        }
    }
}
