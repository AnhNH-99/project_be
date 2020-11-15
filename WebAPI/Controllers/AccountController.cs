using WebAPI.Services.Accounts;
using WebAPI.Models;
using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;
using WebAPI.ViewModels;
using System.Linq;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using System;
using WebAPI.Models.Token;

namespace WebAPI.Controllers
{
    [Authorize]
    [Route("api/[controller]/[action]/{id?}")]
    [ApiController]
    public class AccountController : BaseApiController<AccountViewModel>
    {
        private IAccountService _accountService;
        public AccountController(IAccountService accountService)
        {
            this._accountService = accountService;
        }

        [AllowAnonymous]
        [HttpPost]
        public async Task<IActionResult> Authenticate([FromBody] AuthenticateRequest model)
        {
            var response = await _accountService.Authenticate(model, IpAddress());

            if (response == null)
                return BadRequest(new { message = "Username or password is incorrect" });

            SetTokenCookie(response.RefreshToken);

            return Ok(response);
        }

        [AllowAnonymous]
        [HttpPost]
        public async Task<IActionResult> RefreshToken()
        {
            var refreshToken = Request.Cookies["refreshToken"];
            var response = await _accountService.RefreshToken(refreshToken, IpAddress());

            if (response == null)
                return Unauthorized(new { message = "Invalid token" });

            SetTokenCookie(response.RefreshToken);

            return Ok(response);
        }

        /// <summary>
        /// Get all Account
        /// </summary>
        /// <returns></returns>
        [Authorize(Roles = Role.Admin + "," + Role.Mod)]
        [HttpGet]
        public async Task<IActionResult> GetAllAccount()
        {
            var employees = await _accountService.GetAllAccount();
            apiResult.AppResult.DataResult = employees.AsEnumerable<AccountInfo>();
            return Ok(apiResult.AppResult);
        }

        /// <summary>
        /// Get Account by Id
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [HttpGet]
        public async Task<IActionResult> GetAccountById(int id)
        {
            var app = await _accountService.GetAccountById(id);
            apiResult.AppResult.DataResult = app;
            return Ok(apiResult.AppResult);
        }

        /// <summary>
        /// Create Account
        /// </summary>
        /// <param name="accountInfo"></param>
        /// <returns></returns>
        [Authorize(Roles = Role.Admin + "," + Role.Mod)]
        [HttpPost]
        public async Task<IActionResult> InsertAccount([FromBody] AccountInfo accountInfo)
        {
            var app = await _accountService.InsertAccount(accountInfo);
            return Ok(app.AppResult);
        }

        /// <summary>
        /// Update Account
        /// </summary>
        /// <param name="accountInfo"></param>
        /// <returns></returns>
        [Authorize(Roles = Role.Admin + "," + Role.Mod)]
        [HttpPut]
        public async Task<IActionResult> UpdateAccount([FromBody] AccountInfo accountInfo)
        {
            var app = await _accountService.UpdateAccount(accountInfo);
            return Ok(app.AppResult);
        }

        /// <summary>
        /// Delete Account
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [Authorize(Roles = Role.Admin)]
        [HttpDelete]
        public async Task<IActionResult> DeleteAccount(int id)
        {
            var app = await _accountService.DeleteAccount(id);
            return Ok(app.AppResult);
        }

        /// <summary>
        /// Set token into cookie
        /// </summary>
        /// <param name="token"></param>
        private void SetTokenCookie(string token)
        {
            var cookieOptions = new CookieOptions
            {
                HttpOnly = true,
                Expires = DateTime.UtcNow.AddDays(7)
            };
            Response.Cookies.Append("refreshToken", token, cookieOptions);
        }

        /// <summary>
        /// Get IP address
        /// </summary>
        /// <returns></returns>
        private string IpAddress()
        {
            if (Request.Headers.ContainsKey("X-Forwarded-For"))
                return Request.Headers["X-Forwarded-For"];
            else
                return HttpContext.Connection.RemoteIpAddress.MapToIPv4().ToString();
        }
    }
}
