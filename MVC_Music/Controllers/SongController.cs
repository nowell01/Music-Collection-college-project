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
    public class SongController : ElephantController
    {
        private readonly MusicContext _context;

        public SongController(MusicContext context)
        {
            _context = context;
        }

        // GET: Song
        public async Task<IActionResult> Index(string searchString, int? albumID, int? genreID, int? page, int? pageSizeID, string actionButton, string sortDirection = "asc", string sortField = "Title")
        {
            //NOTE: make sure this array has matching values to the column headings
            string[] sortOptions = new[] { "Title", "Date Recorded", "Album", "Genre" };

            //Count the number of filters applied - start by assuming no filters
            ViewData["Filtering"] = "btn-outline-secondary";
            int numberFilters = 0;
            //Then in each "test" for filtering, add to the count of Filters applied

            PopulateDropDownLists();    //Data for Genre and Album Filter DDL

            var songs = _context.Songs
                .Include(s => s.Album)
                .Include(s => s.Genre)
                .AsNoTracking();

            //Add Filtering
            if (albumID.HasValue)
            {
                songs = songs.Where(s => s.AlbumID == albumID);
                numberFilters++;
            }
            if (genreID.HasValue)
            {
                songs = songs.Where(s => s.GenreID == genreID);
                numberFilters++;
            }
            if (!string.IsNullOrWhiteSpace(searchString))
            {
                songs = songs.Where(s => s.Title.ToUpper().Contains(searchString.ToUpper()));
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
            if (sortField == "Title")
            {
                if (sortDirection == "asc")
                {
                    songs = songs.OrderBy(s => s.Title);
                }
                else
                {
                    songs = songs.OrderByDescending(s => s.Title);
                }
            }

            if (sortField == "Date Recorded")
            {
                if (sortDirection == "asc")
                {
                    songs = songs.OrderBy(s => s.DateRecorded);
                }
                else
                {
                    songs = songs.OrderByDescending(s => s.DateRecorded);
                }
            }
            if (sortField == "Album")
            {
                if (sortDirection == "asc")
                {
                    songs = songs.OrderBy(s => s.Album.Name);
                }
                else
                {
                    songs = songs.OrderByDescending(s => s.Album.Name);
                }
            }

            if (sortField == "Genre")
            {
                if (sortDirection == "asc")
                {
                    songs = songs.OrderBy(s => s.Genre.Name);
                }
                else
                {
                    songs = songs.OrderByDescending(s => s.Genre.Name);
                }
            }

            //Set sort for next time
            ViewData["sortField"] = sortField;
            ViewData["sortDirection"] = sortDirection;

            //Handle Paging
            int pageSize = PageSizeHelper.SetPageSize(HttpContext, pageSizeID, ControllerName());
            ViewData["pageSizeID"] = PageSizeHelper.PageSizeList(pageSize);
            var pagedData = await PaginatedList<Song>.CreateAsync(songs.AsNoTracking(), page ?? 1, pageSize);

            return View(pagedData);
        }

        // GET: Song/Details/5
        [Authorize(Roles = "Admin, Supervisor, Staff")]
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var song = await _context.Songs
                .Include(s => s.Album)
                .Include(s => s.Genre)
                .FirstOrDefaultAsync(m => m.ID == id);

            if (song == null)
            {
                return NotFound();
            }

            return View(song);
        }

        // GET: Song/Create
        [Authorize(Roles = "Admin, Supervisor, Staff")]
        public IActionResult Create()
        {
            var song = new Song();
            PopulateDropDownLists();
            return View();
        }

        // POST: Song/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize(Roles = "Admin, Supervisor, Staff")]
        public async Task<IActionResult> Create([Bind("ID,Title,DateRecorded,AlbumID,GenreID")] Song song)
        {
            try
            {

                if (ModelState.IsValid)
                {
                    _context.Add(song);
                    await _context.SaveChangesAsync();

                    return RedirectToAction("Index", "SongPerformance", new { SongID = song.ID });
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
            PopulateDropDownLists();
            return View(song);
        }

        // GET: Song/Edit/5
        [Authorize(Roles = "Admin, Supervisor, Staff")]
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var song = await _context.Songs
                .Include(s => s.Genre)
                .Include(s => s.Album)
                .AsNoTracking()
                .FirstOrDefaultAsync(s => s.ID == id);

            if (song == null)
            {
                return NotFound();
            }

            PopulateDropDownLists();
            return View(song);
        }

        // POST: Song/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize(Roles = "Admin, Supervisor, Staff")]
        public async Task<IActionResult> Edit(int id, Byte[] RowVersion)
        {
            var songsToUpdate = await _context.Songs
                .Include(s => s.Genre)
                .Include(s => s.Album)
                .FirstOrDefaultAsync(s => s.ID == id);

            if (songsToUpdate == null)
            {
                return NotFound();
            }
            //Put the original RowVersion value in the OriginalValues collection for the entity
            _context.Entry(songsToUpdate).Property("RowVersion").OriginalValue = RowVersion;

            if (await TryUpdateModelAsync<Song>(songsToUpdate, "", s => s.ID, s => s.Title, s => s.DateRecorded, s => s.AlbumID, s => s.GenreID))
            {
                try
                {
                    await _context.SaveChangesAsync();

                    return RedirectToAction("Index", "SongPerformance", new { SongID = songsToUpdate.ID });
                }
                catch (RetryLimitExceededException /* dex */)
                {
                    ModelState.AddModelError("", "Unable to save changes after multiple attempts. Try again, and if the problem persists, see your system administrator.");
                }
                catch (DbUpdateConcurrencyException ex)// Added for concurrency
                {
                    var exceptionEntry = ex.Entries.Single();
                    var clientValues = (Song)exceptionEntry.Entity;
                    var databaseEntry = exceptionEntry.GetDatabaseValues();
                    if (databaseEntry == null)
                    {
                        ModelState.AddModelError("",
                            "Unable to save changes. The Instrument was deleted by another user.");
                    }
                    else
                    {
                        var databaseValues = (Song)databaseEntry.ToObject();
                        if (databaseValues.Title != clientValues.Title)
                            ModelState.AddModelError(nameof(clientValues.Title), "Current value: "
                                + databaseValues.Title);

                        if (databaseValues.DateRecorded != clientValues.DateRecorded)
                            ModelState.AddModelError("DateRecorded", "Current value: "
                                + databaseValues.DateRecorded);

                        //For the foreign key, we need to go to the database to get the information to show
                        Genre? databaseGenre = await _context.Genres.FirstOrDefaultAsync(i => i.ID == databaseValues.GenreID);

                        if (databaseGenre != null)
                        {
                            ModelState.AddModelError("GenreID", $"Current value: {databaseGenre.Name}");
                        }
                        if (databaseValues.AlbumID != clientValues.AlbumID)
                        {
                            Album databaseAlbum = await _context.Albums.FirstAsync(i => i.ID == databaseValues.AlbumID);
                            ModelState.AddModelError("AlbumID", $"Current value: {databaseAlbum.Name}");
                        }

                        ModelState.AddModelError(string.Empty, "The record you attempted to edit "
                                                 + "was modified by another user after you received your values. The "
                                                 + "edit operation was canceled and the current values in the database "
                                                 + "have been displayed. If you still want to save your version of this record, click "
                                                 + "the Save button again. Otherwise click the 'Back to Song List' hyperlink.");

                        //Final steps before redisplaying: Update RowVersion from the Database
                        //and remove the RowVersion error from the ModelState
                        songsToUpdate.RowVersion = databaseValues.RowVersion ?? Array.Empty<byte>();
                        ModelState.Remove("RowVersion");
                    }
                }
                catch (DbUpdateException)
                {

                    ModelState.AddModelError("!", "Unable to save Changes. Try again, and if the problem persists see your system Administation.");

                }
            }
            PopulateDropDownLists();

            return View(songsToUpdate);
        }

        // GET: Song/Delete/5
        [Authorize(Roles = "Admin, Supervisor")]
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var song = await _context.Songs
                .Include(s => s.Album)
                .Include(s => s.Genre)
                .FirstOrDefaultAsync(m => m.ID == id);

            if (song == null)
            {
                return NotFound();
            }

            return View(song);
        }

        // POST: Song/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        [Authorize(Roles = "Admin, Supervisor")]
        public async Task<IActionResult> DeleteConfirmed(int id, Byte[] RowVersion)
        {
            var song = await _context.Songs
                .Include(s => s.Album)
                .Include(s => s.Genre)
                .FirstOrDefaultAsync(m => m.ID == id);
            try
            {
                if (song != null)
                {
                    _context.Entry(song).Property("RowVersion").OriginalValue = RowVersion;
                    _context.Songs.Remove(song);
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
            catch (DbUpdateException)
            {
                ModelState.AddModelError("", "Unable to save Changes. Try again, if the problem persists try contacting with the system administration");

            }
            return View(song);
        }

        private SelectList GenreList(int? selectedId)
        {
            return new SelectList(_context.Genres
                .OrderBy(s => s.Name), "ID", "Name", selectedId);
        }

        private SelectList AlbumList(int? selectedId)
        {
            return new SelectList(_context.Albums
                .OrderBy(s => s.Name), "ID", "Name", selectedId);
        }

        private void PopulateDropDownLists(Song? song = null)
        {
            ViewData["AlbumID"] = AlbumList(song?.AlbumID);
            ViewData["GenreID"] = GenreList(song?.GenreID);
        }
        private bool SongExists(int id)
        {
            return _context.Songs.Any(e => e.ID == id);
        }
    }
}
