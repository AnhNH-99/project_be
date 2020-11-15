using Microsoft.Extensions.Options;
using Microsoft.IdentityModel.Tokens;
using System;
using System.Collections.Generic;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using WebAPI.Models;
using WebAPI.Models.Token;
using WebAPI.Repositories.Accounts;
using WebAPI.ViewModels;

namespace WebAPI.Services.Accounts
{
    public class AccountService : BaseService<AccountViewModel>, IAccountService
    {
        private readonly AppSettings _appSettings;
        IAccountRepository _accountRepository;
        public AccountService(IAccountRepository accountRepository, IOptions<AppSettings> appSettings)
        {
            this._accountRepository = accountRepository;
            this._appSettings = appSettings.Value;
        }

        /// <summary>
        /// Get all Account
        /// </summary>
        /// <returns></returns>
        public async Task<IEnumerable<AccountInfo>> GetAllAccount()
        {
            return await _accountRepository.GetAllAsync();
        }

        /// <summary>
        /// Get Account by Id
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<AccountInfo> GetAccountById(int id)
        {
            return await _accountRepository.FindAsync(id);
        }

        /// <summary>
        /// Create Account
        /// </summary>
        /// <param name="account"></param>
        /// <returns></returns>
        public async Task<AccountViewModel> InsertAccount(AccountInfo account) {
            model.AccountInfo = account;
            model.AccountInfo.CREATED_BY = WebAPI.Helpers.HttpContext.CurrentUser;
            await _accountRepository.InsertAsync(model.AccountInfo);
            return model;
        }

        /// <summary>
        /// Update Account
        /// </summary>
        /// <param name="account"></param>
        /// <returns></returns>
        public async Task<AccountViewModel> UpdateAccount(AccountInfo account)
        {
            model.AccountInfo = account;
            model.AccountInfo.UPDATED_BY = WebAPI.Helpers.HttpContext.CurrentUser;
            await _accountRepository.UpdateAsync(model.AccountInfo);
            return model;
        }

        /// <summary>
        /// Delete Account
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<AccountViewModel> DeleteAccount(int id)
        {
            await _accountRepository.DeleteAsync(id);
            return model;
        }

        /// <summary>
        /// Authenticate
        /// </summary>
        /// <param name="model"></param>
        /// <param name="ipAddress"></param>
        /// <returns></returns>
        public async Task<AuthenticateResponse> Authenticate(AuthenticateRequest model, string ipAddress)
        {
            model.Password = Convert.ToBase64String(Encoding.UTF8.GetBytes(model.Password));
            var user = await _accountRepository.AuthenticateAsync(model.Username, model.Password);

            // return null if user not found
            if (user == null) return null;

            // authentication successful so generate jwt and refresh tokens
            var jwtToken = GenerateJwtToken(user);
            var refreshToken = GenerateRefreshToken(ipAddress);
            // save refresh token
            user.RefreshToken = refreshToken;
            await _accountRepository.UpdateRefreshTokenAsync(user.RefreshToken.Token, user.USER_NO);

            return new AuthenticateResponse(user, jwtToken, refreshToken.Token);
        }

        /// <summary>
        /// RefreshToken
        /// </summary>
        /// <param name="token"></param>
        /// <param name="ipAddress"></param>
        /// <returns></returns>
        public async Task<AuthenticateResponse> RefreshToken(string token, string ipAddress)
        {
            var user = await _accountRepository.RefreshTokenAsync(token);

            // return null if no user found with token
            if (user == null) return null;

            var refreshToken = user.RefreshToken;

            // return null if token is no longer active
            if (!refreshToken.IsActive) return null;

            // replace old refresh token with a new one and save
            var newRefreshToken = GenerateRefreshToken(ipAddress);
            refreshToken.Revoked = DateTime.UtcNow;
            refreshToken.RevokedByIp = ipAddress;
            refreshToken.ReplacedByToken = newRefreshToken.Token;
            user.RefreshToken = newRefreshToken;
            await _accountRepository.UpdateRefreshTokenAsync(user.RefreshToken.Token, user.USER_NO);

            // generate new jwt
            var jwtToken = GenerateJwtToken(user);

            return new AuthenticateResponse(user, jwtToken, newRefreshToken.Token);
        }

        /// <summary>
        /// Generate Jwt Token
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        private string GenerateJwtToken(AccountInfo user)
        {
            var tokenHandler = new JwtSecurityTokenHandler();
            var key = Encoding.ASCII.GetBytes(_appSettings.Secret);
            var tokenDescriptor = new SecurityTokenDescriptor
            {
                Subject = new ClaimsIdentity(new Claim[]
                {
                    new Claim(ClaimTypes.Name, user.USER_NAME.ToString()),
                    new Claim(ClaimTypes.Role, user.ROLE.ToString())
                }),
                Expires = DateTime.UtcNow.AddMinutes(15),
                SigningCredentials = new SigningCredentials(new SymmetricSecurityKey(key), SecurityAlgorithms.HmacSha256Signature)
            };
            var token = tokenHandler.CreateToken(tokenDescriptor);
            return tokenHandler.WriteToken(token);
        }

        /// <summary>
        /// Generate Refresh Token
        /// </summary>
        /// <param name="ipAddress"></param>
        /// <returns></returns>
        private RefreshToken GenerateRefreshToken(string ipAddress)
        {
            using (var rngCryptoServiceProvider = new RNGCryptoServiceProvider())
            {
                var randomBytes = new byte[64];
                rngCryptoServiceProvider.GetBytes(randomBytes);
                return new RefreshToken
                {
                    Token = Convert.ToBase64String(randomBytes),
                    Expires = DateTime.UtcNow.AddDays(7),
                    Created = DateTime.UtcNow,
                    CreatedByIp = ipAddress
                };
            }
        }
    }
}
