using System.Threading.Tasks;
using WebAPI.Models;
using WebAPI.Repositories.Interfaces;

namespace WebAPI.Repositories.Accounts
{
    public interface IAccountRepository : IRepositoryBase<AccountInfo>
    {
        /// <summary>
        /// Authenticate Account
        /// </summary>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public Task<AccountInfo> AuthenticateAsync(string userName, string password);

        /// <summary>
        /// Get Refresh Token
        /// </summary>
        /// <param name="token"></param>
        /// <returns></returns>
        public Task<AccountInfo> RefreshTokenAsync(string token);

        /// <summary>
        /// Update Refresh Token
        /// </summary>
        /// <param name="token"></param>
        /// <param name="userNo"></param>
        /// <returns></returns>
        public Task UpdateRefreshTokenAsync(string token, int userNo);
    }
}
