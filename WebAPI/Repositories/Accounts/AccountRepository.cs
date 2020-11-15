using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WebAPI.Models;
using Dapper;
using System.Text;
using System.Data;

namespace WebAPI.Repositories.Accounts
{
    public class AccountRepository : RepositoryBase, IAccountRepository
    {
        public AccountRepository(string connectionString) : base(connectionString) { }

        /// <summary>
        /// Get all Account
        /// </summary>
        /// <returns></returns>
        public async Task<IEnumerable<AccountInfo>> GetAllAsync()
        {
            using (var conn = OpenDBConnection())
            {
                var sql = "SELECT * FROM dbo.AccountInfo";
                return await conn.QueryAsync<AccountInfo>(sql);
            }
        }

        /// <summary>
        /// Check Account
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="userNo"></param>
        /// <returns></returns>
        public bool CheckAccount(IDbConnection conn, int userNo)
        {
            StringBuilder sql = new StringBuilder();
            sql.Length = 0;
            sql.AppendLine("SELECT SUM(CNT)                           ");
            sql.AppendLine("FROM (                                    ");
            sql.AppendLine("        SELECT                            ");
            sql.AppendLine("           COUNT(*) AS CNT                ");
            sql.AppendLine("        FROM dbo.AccountInfo              ");
            sql.AppendLine("        WHERE USER_NO = @USER_NO          ");
            sql.AppendLine("     )                                    ");

            int count = conn.ExecuteScalar<int>(sql.ToString(), new { user_no = userNo });
            return count > 0;
        }

        /// <summary>
        /// Find Account
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<AccountInfo> FindAsync(int id)
        {
            using (var conn = OpenDBConnection())
            {
                var sql = "SELECT * FROM dbo.AccountInfo WHERE USER_NO = @user_no";
                var param = new { user_no = id };
                return await conn.QueryFirstOrDefaultAsync<AccountInfo>(sql, param);
            }
        }

        /// <summary>
        /// Create Account
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        public async Task<int> InsertAsync(AccountInfo entity)
        {
            using (var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;

                sql.AppendLine(" INSERT INTO                            ");
                sql.AppendLine("      dbo.Accounts (USER_CD,            ");
                sql.AppendLine("                    USER_NAME,          ");
                sql.AppendLine("                    PASSWORD,           ");
                sql.AppendLine("                    FULL_NAME,          ");
                sql.AppendLine("                    EMAIL,              ");
                sql.AppendLine("                    PHONE,              ");
                sql.AppendLine("                    ADDRESS,            ");
                sql.AppendLine("                    ROLE,               ");
                sql.AppendLine("                    STATUS_FLG,         ");
                sql.AppendLine("                    CREATED_AT,         ");
                sql.AppendLine("                    CREATED_BY)         ");
                sql.AppendLine("      dbo.AccountInfo                   ");
                sql.AppendLine(" VALUES (@USER_CD,                      ");
                sql.AppendLine("         @USER_NAME,                    ");
                sql.AppendLine("         @PASSWORD,                     ");
                sql.AppendLine("         @FULL_NAME,                    ");
                sql.AppendLine("         @EMAIL,                        ");
                sql.AppendLine("         @PHONE,                        ");
                sql.AppendLine("         @ADDRESS,                      ");
                sql.AppendLine("         @ROLE,                         ");
                sql.AppendLine("         @STATUS_FLG,                   ");
                sql.AppendLine("         getdate(),                     ");
                sql.AppendLine("         @CREATED_BY)                   ");
                entity.USER_CD = Guid.NewGuid().ToString().GetHashCode().ToString("x");
                entity.PASSWORD = Convert.ToBase64String(Encoding.UTF8.GetBytes(entity.PASSWORD));
                var param = new
                {
                    user_cd = entity.USER_CD,
                    user_name = entity.USER_NAME,
                    password = entity.PASSWORD,
                    full_name = entity.FULL_NAME,
                    email = entity.EMAIL,
                    phone = entity.PHONE,
                    address = entity.ADDRESS,
                    role = entity.ROLE,
                    status_flg = entity.STATUS_FLG,
                    created_by = entity.CREATED_BY
                };
                return await conn.ExecuteAsync(sql.ToString(), param);
            }
        }

        /// <summary>
        /// Update Account
        /// </summary>
        /// <param name="entityToUpdate"></param>
        /// <returns></returns>
        public async Task<int> UpdateAsync(AccountInfo entity)
        {
            using (var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;
                sql.AppendLine(" UPDATE dbo.Accounts                    ");
                sql.AppendLine("    SET PASSWORD    = @PASSWORD,        ");
                sql.AppendLine("        FULL_NAME   = @FULL_NAME,       ");
                sql.AppendLine("        EMAIL       = @EMAIL,           ");
                sql.AppendLine("        PHONE       = @PHONE,           ");
                sql.AppendLine("        ADDRESS     = @ADDRESS,         ");
                sql.AppendLine("        ROLE        = @ROLE,            ");
                sql.AppendLine("        STATUS_FLG  = @STATUS_FLG,      ");
                sql.AppendLine("        UPDASTED_AT = getdate(),        ");
                sql.AppendLine("        UPDASTED_BY = @UPDASTED_BY      ");
                sql.AppendLine(" WHERE USER_NO = @USER_NO               ");

                var param = new
                {
                    password    = entity.PASSWORD,
                    full_name   = entity.FULL_NAME,
                    email       = entity.EMAIL,
                    phone       = entity.PHONE,
                    address     = entity.ADDRESS,
                    role        = entity.ROLE,
                    status_flg  = entity.STATUS_FLG,
                    updated_by  = entity.CREATED_BY,
                    user_no     = entity.USER_NO
                };
                return await conn.ExecuteAsync(sql.ToString(), param);
            }
        }

        /// <summary>
        /// Delete Account
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<int> DeleteAsync(int id)
        {
            using (var conn = OpenDBConnection())
            {
                var sql = "DELETE FROM dbo.AccountInfo WHERE USER_NO = @USER_NO";
                return await conn.ExecuteAsync(sql.ToString(), new { user_no = id });
            }
        }

        /// <summary>
        /// Authenticate Account
        /// </summary>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public async Task<AccountInfo> AuthenticateAsync(string username, string password)
        {
            using (var conn = OpenDBConnection())
            {
                var sql = "SELECT * FROM dbo.Accounts WHERE USER_NAME = @USER_NAME AND PASSWORD = @PASSWORD";
                var param = new { user_name = username, password = password };
                return await conn.QuerySingleOrDefaultAsync<AccountInfo>(sql, param);
            }
        }

        /// <summary>
        /// Get Refresh Token
        /// </summary>
        /// <param name="token"></param>
        /// <returns></returns>
        public async Task<AccountInfo> RefreshTokenAsync(string token)
        {
            using (var conn = OpenDBConnection())
            {
                var sql = "SELECT * FROM dbo.Accounts WHERE REFRESH_TOKEN = @JWT_TOKEN";
                var param = new { jwt_token = token };
                return await conn.QuerySingleOrDefaultAsync<AccountInfo>(sql, param);
            }
        }

        /// <summary>
        /// Update Refresh Token
        /// </summary>
        /// <param name="token"></param>
        /// <param name="userNo"></param>
        /// <returns></returns>
        public async Task UpdateRefreshTokenAsync(string token, int userNo)
        {
            StringBuilder sql = new StringBuilder();
            sql.Length = 0;

            using (var conn = OpenDBConnection())
            {
                sql.AppendLine(" UPDATE dbo.Accounts                             ");
                sql.AppendLine("    SET REFRESH_TOKEN   = @REFRESH_TOKEN         ");
                sql.AppendLine(" WHERE USER_NO = @USER_NO                        ");
                var param = new { refresh_token = token, user_no = userNo };
                await conn.QueryAsync(sql.ToString(), param);
            }
        }
    }
}
