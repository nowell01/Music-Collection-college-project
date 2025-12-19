using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using MVC_Music.CustomControllers;
using MVC_Music.Data;
using MVC_Music.ViewModels;
using System.ComponentModel.DataAnnotations;

namespace MVC_Music.Controllers
{
    [Authorize(Roles = "Security")]
    public class UserRoleController : CognizantController
    {
        private readonly ApplicationDbContext _context;
        private readonly UserManager<IdentityUser> _userManager;

        public UserRoleController(ApplicationDbContext context, UserManager<IdentityUser> userManager)
        {
            _context = context;
            _userManager = userManager;
        }

        // GET: User
        public async Task<IActionResult> Index()
        {
            var users = await (from u in _context.Users
                               .OrderBy(u => u.UserName)
                               select new UserVM
                               {
                                   Id = u.Id,
                                   UserName = u.UserName ?? ""
                               }).ToListAsync();

            foreach (var u in users)
            {
                var user = await _userManager.FindByIdAsync(u.Id);
                if (user != null)
                {
                    u.UserRoles = (List<string>)await _userManager.GetRolesAsync(user);
                    //Note: we needed the explicit cast above because GetRolesAsync() returns an IList<string>
                }
            }
            ;
            return View(users);
        }

        // GET: Users/Edit/5
        [Authorize(Roles = "Security")]
        public async Task<IActionResult> Edit(string id)
        {
            if (id == null)
            {
                return new BadRequestResult();
            }
            var _user = await _userManager.FindByIdAsync(id);//IdentityRole
            if (_user == null)
            {
                return NotFound();
            }
            UserVM user = new UserVM
            {
                Id = _user.Id,
                UserName = _user.UserName ?? "",
                UserRoles = (List<string>)await _userManager.GetRolesAsync(_user)
            };
            PopulateAssignedRoleData(user);

            if(user.UserName == User.Identity!.Name)
            {
                ModelState.AddModelError("", "You cannot change your own role assignments.");
                ViewData["NoSubmit"] = "disabled=disabled";
            }

            return View(user);
        }

        // POST: Users/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize(Roles = "Security")]
        public async Task<IActionResult> Edit(string Id, string[] selectedRoles)
        {
            var _user = await _userManager.FindByIdAsync(Id);//IdentityRole
            if (_user == null) return View(); // Quick null check

            UserVM user = new UserVM
            {
                Id = _user.Id,
                UserName = _user.UserName ?? "",
                UserRoles = (List<string>)await _userManager.GetRolesAsync(_user)
            };
            if (user.UserName == User.Identity!.Name)
            {
                ModelState.AddModelError("", "You cannot change your own role assignments.");
                ViewData["NoSubmit"] = "disabled=disabled";
            }
            else
            {
                try
                {
                    await UpdateUserRoles(selectedRoles, user);
                    return RedirectToAction("Index");
                }
                catch (Exception)
                {
                    ModelState.AddModelError(string.Empty,
                                    "Unable to save changes.");
                }
            }
            PopulateAssignedRoleData(user);
            return View(user);
        }

        private void PopulateAssignedRoleData(UserVM user)
        {//Prepare checkboxes for all Roles
            var allRoles = _context.Roles;
            var currentRoles = user.UserRoles;
            var viewModel = new List<RoleVM>();
            foreach (var r in allRoles)
            {
                var roleName = r.Name ?? string.Empty;
                viewModel.Add(new RoleVM
                {
                    RoleId = r.Id,
                    RoleName = roleName,
                    Assigned = currentRoles.Contains(roleName)
                });
            }
            ViewBag.Roles = viewModel;
        }

        private async Task UpdateUserRoles(string[] selectedRoles, UserVM userToUpdate)
        {
            var UserRoles = userToUpdate.UserRoles;//Current roles user is in
            var _user = await _userManager.FindByIdAsync(userToUpdate.Id);//IdentityUser
            if (_user == null) return; // Quick null check

            if (selectedRoles == null)
            {
                //No roles selected so just remove any currently assigned
                foreach (var r in UserRoles)
                {
                    await _userManager.RemoveFromRoleAsync(_user, r);
                }
            }
            else
            {
                //At least one role checked so loop through all the roles
                //and add or remove as required

                //We need to do this next line because foreach loops don't always work well
                //for data returned by EF when working async.  Pulling it into an IList<>
                //first means we can safely loop over the colleciton making async calls and avoid
                //the error 'New transaction is not allowed because there are other threads running in the session'
                IList<IdentityRole> allRoles = _context.Roles.ToList<IdentityRole>();

                foreach (var r in allRoles)
                {
                    var roleName = r.Name ?? string.Empty; // Null-safe
                    if (selectedRoles.Contains(roleName))
                    {
                        if (!UserRoles.Contains(roleName))
                        {
                            await _userManager.AddToRoleAsync(_user, roleName);
                        }
                    }
                    else
                    {
                        if (UserRoles.Contains(roleName))
                        {
                            await _userManager.RemoveFromRoleAsync(_user, roleName);
                        }
                    }
                }
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _context.Dispose();
                _userManager.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
