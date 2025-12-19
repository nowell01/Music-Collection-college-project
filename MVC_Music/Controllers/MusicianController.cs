using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Storage;
using MVC_Music.CustomControllers;
using MVC_Music.Data;
using MVC_Music.Models;
using MVC_Music.Utilities;
using MVC_Music.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Threading.Tasks;

namespace MVC_Music.Controllers
{
    [Authorize]
    public class MusicianController : ElephantController
    {
        private readonly MusicContext _context;

        public MusicianController(MusicContext context)
        {
            _context = context;
        }

        // GET: Musician
        public async Task<IActionResult> Index(string SearchName, string SearchPhone, int? InstrumentID, int? OtherInstrumentID, int? page, int? pageSizeID, string actionButton, string sortDirection = "asc", string sortField = "Musician")
        {
            //List of sort options.
            //NOTE: make sure this array has matching values to the column headings
            string[] sortOptions = new[] { "Musician", "Phone", "Age", "Instruments" };

            //Count the number of filters applied - start by assuming no filters
            ViewData["Filtering"] = "btn-outline-secondary";
            int numberFilters = 0;
            //Then in each "test" for filtering, add to the count of Filters applied

            PopulateDropDownLists();    //Data for Doctor and MedicalTrial Filter DDL
            ViewData["OtherInstrumentID"] = ViewData["InstrumentID"];

            var musicians = _context.Musicians
                .Include(m => m.Instrument)
                .Include(m=>m.MusicianDocuments)
                .Include(m => m.MusicianThumbnail)
                .Include(m=>m.Plays).ThenInclude(m=>m.Instrument)
                .AsNoTracking();

            //Add as many filters as needed
            if (InstrumentID.HasValue)
            {
                musicians = musicians.Where(p => p.InstrumentID == InstrumentID);
                numberFilters++;
            }
            if (OtherInstrumentID.HasValue)
            {
                musicians = musicians.Where(p => p.Plays.Any(p => p.InstrumentID == OtherInstrumentID));
                numberFilters++;
            }
            if (!String.IsNullOrEmpty(SearchName))
            {
                musicians = musicians.Where(p => p.LastName.ToUpper().Contains(SearchName.ToUpper())
                                       || p.FirstName.ToUpper().Contains(SearchName.ToUpper()));
                numberFilters++;
            }
            if (!String.IsNullOrEmpty(SearchPhone))
            {
                musicians = musicians.Where(p => p.Phone.ToUpper().Contains(SearchPhone.ToUpper()));
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
            if (sortField == "Phone")
            {
                if (sortDirection == "asc")
                {
                    musicians = musicians
                        .OrderBy(p => p.Phone)
                        .ThenBy(p => p.LastName)
                        .ThenBy(p => p.FirstName);
                }
                else
                {
                    musicians = musicians
                        .OrderByDescending(p => p.Phone)
                        .ThenBy(p => p.LastName)
                        .ThenBy(p => p.FirstName);
                }
            }
            else if (sortField == "Age")
            {
                if (sortDirection == "asc")
                {
                    musicians = musicians
                        .OrderByDescending(p => p.DOB)
                        .ThenBy(p => p.LastName)
                        .ThenBy(p => p.FirstName);
                }
                else
                {
                    musicians = musicians
                        .OrderBy(p => p.DOB)
                        .ThenBy(p => p.LastName)
                        .ThenBy(p => p.FirstName);
                }
            }
            else if (sortField == "Instruments")
            {
                if (sortDirection == "asc")
                {
                    musicians = musicians
                        .OrderBy(p => p.Instrument.Name)
                        .ThenBy(p => p.LastName)
                        .ThenBy(p => p.FirstName);
                }
                else
                {
                    musicians = musicians
                        .OrderByDescending(p => p.Instrument.Name)
                        .ThenBy(p => p.LastName)
                        .ThenBy(p => p.FirstName);
                }
            }
            else //Sorting by Musician Name
            {
                if (sortDirection == "asc")
                {
                    musicians = musicians
                        .OrderBy(p => p.LastName)
                        .ThenBy(p => p.FirstName);
                }
                else
                {
                    musicians = musicians
                        .OrderByDescending(p => p.LastName)
                        .ThenByDescending(p => p.FirstName);
                }
            }
            //Set sort for next time
            ViewData["sortField"] = sortField;
            ViewData["sortDirection"] = sortDirection;

            //Handle Paging
            int pageSize = PageSizeHelper.SetPageSize(HttpContext, pageSizeID, ControllerName());
            ViewData["pageSizeID"] = PageSizeHelper.PageSizeList(pageSize);

            var pagedData = await PaginatedList<Musician>.CreateAsync(musicians.AsNoTracking(), page ?? 1, pageSize);

            return View(pagedData);
        }

        // GET: Musician/Details/5
        [Authorize(Roles = "Admin, Supervisor, Staff")]
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var musician = await _context.Musicians
                .Include(m => m.Instrument)
                .Include(m => m.MusicianPhoto)
                .Include(m => m.MusicianDocuments)
                .Include(m => m.Plays).ThenInclude(m => m.Instrument)
                .AsNoTracking()
                .FirstOrDefaultAsync(m => m.ID == id);
            if (musician == null)
            {
                return NotFound();
            }

            return View(musician);
        }

        // GET: Musician/Create
        [Authorize(Roles = "Admin, Supervisor, Staff")]
        public IActionResult Create()
        {
            Musician musician = new Musician();
            PopulateAssignedPlaysData(musician);
            PopulateDropDownLists();
            return View();
        }

        // POST: Musician/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize(Roles = "Admin, Supervisor, Staff")]
        public async Task<IActionResult> Create([Bind("FirstName,MiddleName,LastName," +
            "Phone,DOB,SIN,InstrumentID")] Musician musician, string[] selectedOptions, List<IFormFile> theFiles, IFormFile? thePicture)
        {
            try
            {
                //Add the selected instruments
                if (selectedOptions != null)
                {
                    foreach (var instrument in selectedOptions)
                    {
                        var instrumentToAdd = new Play { MusicianID = musician.ID, InstrumentID = int.Parse(instrument) };
                        musician.Plays.Add(instrumentToAdd);
                    }
                }

                if (ModelState.IsValid)
                {
                    await AddDocumentsAsync(musician, theFiles);
                    if (thePicture != null) await AddPicture(musician, thePicture);
                    _context.Add(musician);
                    await _context.SaveChangesAsync();
                    return RedirectToAction("Details", new { musician.ID });
                }
            }
            catch (RetryLimitExceededException /* dex */)//This is a Transaction in the Database!
            {
                ModelState.AddModelError("", "Unable to save changes after multiple attempts. " +
                    "Try again, and if the problem persists, see your system administrator.");
            }
            catch (DbUpdateException dex)
            {
                string message = dex.GetBaseException().Message;
                if (message.Contains("UNIQUE") && message.Contains("Musicians.SIN"))
                {
                    ModelState.AddModelError("SIN", "Unable to save changes. Remember, " +
                        "you cannot have duplicate SIN numbers.");
                }
                else
                {
                    ModelState.AddModelError("", "Unable to save changes. Try again, and if the problem persists see your system administrator.");
                }
            }

            PopulateAssignedPlaysData(musician);
            PopulateDropDownLists(musician);
            return View(musician);
        }

        // GET: Musician/Edit/5
        [Authorize(Roles = "Admin, Supervisor, Staff")]
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var musician = await _context.Musicians
                .Include(p => p.MusicianPhoto)
                .Include(p => p.MusicianDocuments)
                .Include(p => p.Plays).ThenInclude(p => p.Instrument)
                .FirstOrDefaultAsync(p => p.ID == id);

            if (musician == null)
            {
                return NotFound();
            }

            PopulateAssignedPlaysData(musician);
            PopulateDropDownLists(musician);
            return View(musician);
        }


        // POST: Musician/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize(Roles = "Admin, Supervisor, Staff")]
        public async Task<IActionResult> Edit(int id, string[] selectedOptions, Byte[] RowVersion, string? chkRemoveImage, List<IFormFile> theFiles, IFormFile? thePicture)
        {
            //Go get the musician to update
            var musicianToUpdate = await _context.Musicians
                .Include(p => p.MusicianPhoto)
                .Include(p => p.MusicianDocuments)
                .Include(p => p.Plays).ThenInclude(p => p.Instrument)
                .FirstOrDefaultAsync(p => p.ID == id);

            //Check that you got it or exit with a not found error
            if (musicianToUpdate == null)
            {
                return NotFound();
            }

            //Update the additional instruments playes
            UpdatePlays(selectedOptions, musicianToUpdate);

            //Put the original RowVersion value in the OriginalValues collection for the entity
            _context.Entry(musicianToUpdate).Property("RowVersion").OriginalValue = RowVersion;

            //Try updating it with the values posted
            if (await TryUpdateModelAsync<Musician>(musicianToUpdate, "",
                p => p.FirstName, p => p.MiddleName, p => p.LastName, p => p.Phone, p => p.DOB,
                p => p.SIN, p => p.InstrumentID))
            {
                try
                {
                    //For the image
                    if (chkRemoveImage != null)
                    {
                        //If we are just deleting the two versions of the photo, we need to make sure the Change Tracker knows
                        //about them both so go get the Thumbnail since we did not include it.
                        musicianToUpdate.MusicianThumbnail = _context.MusicianThumbnails
                            .Where(p => p.MusicianID == musicianToUpdate.ID).FirstOrDefault();
                        //Then, setting them to null will cause them to be deleted from the database.
                        musicianToUpdate.MusicianPhoto = null;
                        musicianToUpdate.MusicianThumbnail = null;
                    }
                    else
                    {
                        if (thePicture != null) await AddPicture(musicianToUpdate, thePicture);
                    }

                    await AddDocumentsAsync(musicianToUpdate, theFiles);
                    await _context.SaveChangesAsync();
                    return RedirectToAction("Details", new { musicianToUpdate.ID });
                }
                catch (DbUpdateConcurrencyException ex)
                {
                    var exceptionEntry = ex.Entries.Single();
                    var clientValues = (Musician)exceptionEntry.Entity;
                    var databaseEntry = exceptionEntry.GetDatabaseValues();
                    if (databaseEntry == null)
                    {
                        ModelState.AddModelError("",
                            "Unable to save changes. The Musician was deleted by another user.");
                    }
                    else
                    {
                        var databaseValues = (Musician)databaseEntry.ToObject();
                        if (databaseValues.FirstName != clientValues.FirstName)
                            ModelState.AddModelError("FirstName", "Current value: "
                                + databaseValues.FirstName);
                        if (databaseValues.MiddleName != clientValues.MiddleName)
                            ModelState.AddModelError("MiddleName", "Current value: "
                                + databaseValues.MiddleName);
                        if (databaseValues.LastName != clientValues.LastName)
                            ModelState.AddModelError("LastName", "Current value: "
                                + databaseValues.LastName);
                        if (databaseValues.SIN != clientValues.SIN)
                            ModelState.AddModelError("SIN", "Current value: "
                                + databaseValues.SINFormatted);
                        if (databaseValues.DOB != clientValues.DOB)
                            ModelState.AddModelError("DOB", "Current value: "
                                + String.Format("{0:d}", databaseValues.DOB));
                        if (databaseValues.Phone != clientValues.Phone)
                            ModelState.AddModelError("Phone", "Current value: "
                                + databaseValues.PhoneFormatted);
                        //For the foreign key, we need to go to the database to get the information to show
                        if (databaseValues.InstrumentID != clientValues.InstrumentID)
                        {
                            Instrument? databaseInstrument = await _context.Instruments.FirstOrDefaultAsync(i => i.ID == databaseValues.InstrumentID);
                            ModelState.AddModelError("InstrumentID", $"Current value: {databaseInstrument?.Name}");
                        }
                        ModelState.AddModelError(string.Empty, "The record you attempted to edit "
                                + "was modified by another user after you received your values. The "
                                + "edit operation was canceled and the current values in the database "
                                + "have been displayed. If you still want to save your version of this record, click "
                                + "the Save button again. Otherwise click the 'Back to Musician List' hyperlink.");
                        //Final steps before redisplaying: Update RowVersion from the Database
                        //and remove the RowVersion error from the ModelState
                        musicianToUpdate.RowVersion = databaseValues.RowVersion ?? Array.Empty<byte>();
                        ModelState.Remove("RowVersion");
                    }
                }
                catch (RetryLimitExceededException /* dex */)
                {
                    ModelState.AddModelError("", "Unable to save changes after multiple attempts. " +
                        "Try again, and if the problem persists, see your system administrator.");
                }
                catch (DbUpdateException dex)
                {
                    string message = dex.GetBaseException().Message;
                    if (message.Contains("UNIQUE") && message.Contains("Musicians.SIN"))
                    {
                        ModelState.AddModelError("SIN", "Unable to save changes. Remember, " +
                            "you cannot have duplicate SIN numbers.");
                    }
                    else
                    {
                        ModelState.AddModelError("", "Unable to save changes. Try again, and if the problem persists see your system administrator.");
                    }
                }

            }

            PopulateAssignedPlaysData(musicianToUpdate);
            PopulateDropDownLists(musicianToUpdate);
            return View(musicianToUpdate);
        }

        // GET: Musician/Delete/5
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var musician = await _context.Musicians
                .Include(m => m.Instrument)
                .Include(m=>m.MusicianPhoto)
                .Include(m => m.MusicianDocuments)
                .Include(m => m.Plays).ThenInclude(p => p.Instrument)
                .AsNoTracking()
                .FirstOrDefaultAsync(m => m.ID == id);

            if (musician == null)
            {
                return NotFound();
            }

            return View(musician);
        }

        // POST: Musician/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> DeleteConfirmed(int id, Byte[] RowVersion)
        {
            var musician = await _context.Musicians
                .Include(p => p.Instrument)
                .Include(m => m.MusicianPhoto)
                .Include(m => m.MusicianDocuments)
                .Include(p => p.Plays).ThenInclude(p => p.Instrument)
                .FirstOrDefaultAsync(m => m.ID == id);

            try
            {
                if (musician != null)
                {
                    _context.Entry(musician).Property("RowVersion").OriginalValue = RowVersion;
                    _context.Musicians.Remove(musician);
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
                //Note: there is really no reason a delete should fail if you can "talk" to the database.
                ModelState.AddModelError("", "Unable to delete record. Try again, and if the problem persists see your system administrator.");
            }

            return View(musician);

        }

        private SelectList InstrumentList(int? selectedId)
        {
            return new SelectList(_context.Instruments
                .OrderBy(d => d.Name), "ID", "Name", selectedId);
        }

        private void PopulateDropDownLists(Musician? musician = null)
        {
            ViewData["InstrumentID"] = InstrumentList(musician?.InstrumentID);
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

        private async Task AddDocumentsAsync(Musician musician, List<IFormFile> theFiles)
        {
            if (musician == null || theFiles == null)
                return;

            foreach (var f in theFiles)
            {
                if (f != null)
                {
                    string mimeType = f.ContentType;
                    string fileName = Path.GetFileName(f.FileName);
                    long fileLength = f.Length;

                    //Note: you could filter for mime types if you only want to allow
                    //certain types of files.  I am allowing everything.
                    if (!string.IsNullOrWhiteSpace(fileName) && fileLength > 0)//Looks like we have a file!!!
                    {
                        using var memoryStream = new MemoryStream();
                        await f.CopyToAsync(memoryStream);

                        var document = new MusicianDocument
                        {
                            MimeType = mimeType,
                            FileName = fileName,
                            FileContent = new FileContent
                            {
                                Content = memoryStream.ToArray()
                            }
                        };

                        musician.MusicianDocuments.Add(document);
                    }
                }
            }
        }

        private async Task AddPicture(Musician musician, IFormFile? thePicture)
        {
            //Get the picture and save it with the Patient (2 sizes)
            if (thePicture != null)
            {
                string mimeType = thePicture.ContentType;
                long fileLength = thePicture.Length;
                if (!(mimeType == "" || fileLength == 0))//Looks like we have a file!!!
                {
                    if (mimeType.Contains("image"))
                    {
                        using var memoryStream = new MemoryStream();
                        await thePicture.CopyToAsync(memoryStream);
                        var pictureArray = memoryStream.ToArray();//Gives us the Byte[]

                        //Check if we are replacing or creating new
                        if (musician.MusicianPhoto != null)
                        {
                            //We already have pictures so just replace the Byte[]
                            musician.MusicianPhoto.Content = ResizeImage.ShrinkImageWebp(pictureArray, 500, 600);

                            //Get the Thumbnail so we can update it.  Remember we didn't include it
                            musician.MusicianThumbnail = _context.MusicianThumbnails.Where(p => p.MusicianID == musician.ID).FirstOrDefault();
                            if (musician.MusicianThumbnail != null)
                            {
                                musician.MusicianThumbnail.Content = ResizeImage.ShrinkImageWebp(pictureArray, 75, 90);
                            }
                        }
                        else //No pictures saved so start new
                        {
                            musician.MusicianPhoto = new MusicianPhoto
                            {
                                Content = ResizeImage.ShrinkImageWebp(pictureArray, 500, 600),
                                MimeType = "image/webp"
                            };
                            musician.MusicianThumbnail = new MusicianThumbnail
                            {
                                Content = ResizeImage.ShrinkImageWebp(pictureArray, 75, 90),
                                MimeType = "image/webp"
                            };
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Prepares a collection of checkbox ViewModel objects for each instrument.
        /// Sets Assigned = true for instruments already played by the musician.
        /// </summary>
        /// <param name="musician">The Musician entity with included Plays</param>
        private void PopulateAssignedPlaysData(Musician musician)
        {
            //Ensure Play records are loaded with .Include before calling this method
            var allInstruments = _context.Instruments.ToList();

            var assignedInstrumentIds = new HashSet<int>(
                musician.Plays.Select(pc => pc.InstrumentID));

            var instrumentOptions = allInstruments.Select(instrument => new CheckOptionVM
            {
                ID = instrument.ID,
                DisplayText = instrument.Name,
                Assigned = assignedInstrumentIds.Contains(instrument.ID)
            }).ToList();

            ViewData["InstrumentOptions"] = instrumentOptions;
        }

        /// <summary>
        /// Updates the Plays for a Musician to match the selected checkboxes.
        /// </summary>
        /// <param name="selectedOptions">IDs of the selected instruments</param>
        /// <param name="musicianToUpdate">The Musician entity to update</param>
        private void UpdatePlays(string[]? selectedOptions, Musician musicianToUpdate)
        {
            // If no options are selected, clear all instruments
            if (selectedOptions == null || selectedOptions.Length == 0)
            {
                musicianToUpdate.Plays = new List<Play>();
                return;
            }

            var selectedIds = new HashSet<string>(selectedOptions);
            var currentInstrumentIds = new HashSet<int>(
                musicianToUpdate.Plays.Select(pc => pc.InstrumentID));

            foreach (var instrument in _context.Instruments)
            {
                string instrumentIdStr = instrument.ID.ToString();

                bool isSelected = selectedIds.Contains(instrumentIdStr);
                bool isCurrentlyAssigned = currentInstrumentIds.Contains(instrument.ID);

                if (isSelected && !isCurrentlyAssigned)
                {
                    // Add new instrument
                    musicianToUpdate.Plays.Add(new Play
                    {
                        MusicianID = musicianToUpdate.ID,
                        InstrumentID = instrument.ID
                    });
                }
                else if (!isSelected && isCurrentlyAssigned)
                {
                    // Remove unselected instrument
                    var instrumentToRemove = musicianToUpdate.Plays
                        .SingleOrDefault(pc => pc.InstrumentID == instrument.ID);

                    if (instrumentToRemove != null)
                    {
                        _context.Remove(instrumentToRemove);
                    }
                }
            }
        }

        private bool MusicianExists(int id)
        {
            return _context.Musicians.Any(e => e.ID == id);
        }
    }
}
