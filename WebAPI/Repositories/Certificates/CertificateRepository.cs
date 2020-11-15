using Dapper;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using WebAPI.Models;
using WebAPI.Repositories.Certificates;
using System.Text.RegularExpressions;

namespace WebAPI.Repositories.Certificates
{
    public class CertificateRepository : RepositoryBase, ICertificateRepository
    {
        public CertificateRepository(string connectionString) : base(connectionString) { }

        /// <summary>
        /// Get all Certificate
        /// </summary>
        /// <returns></returns>
        public async Task<IEnumerable<CertificateInfo>> GetAllAsync()
        {
            using (var conn = OpenDBConnection())
            {
                var sql = "SELECT * FROM dbo.Certificate";
                return await conn.QueryAsync<CertificateInfo>(sql);
            }
        }

        ///// <summary>
        ///// Check Certificate
        ///// </summary>
        ///// <param name="conn"></param>
        ///// <param name="userNo"></param>
        ///// <returns></returns>
        //public bool CheckCertificate(IDbConnection conn, int userNo)
        //{
        //    StringBuilder sql = new StringBuilder();
        //    sql.Length = 0;
        //    sql.AppendLine("SELECT SUM(CNT)                           ");
        //    sql.AppendLine("FROM (                                    ");
        //    sql.AppendLine("        SELECT                            ");
        //    sql.AppendLine("           COUNT(*) AS CNT                ");
        //    sql.AppendLine("        FROM dbo.Certificates                 ");
        //    sql.AppendLine("        WHERE USER_NO = @USER_NO          ");
        //    sql.AppendLine("     )                                    ");

        //    int count = conn.ExecuteScalar<int>(sql.ToString(), new { user_no = userNo });
        //    return count > 0;
        //}

        /// <summary>
        /// Find Certificate
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<CertificateInfo> FindAsync(int id)
        {
            using (var conn = OpenDBConnection())
            {
                var sql = "SELECT * FROM dbo.Certificate WHERE Id = @_id";
                var param = new { _id = id };
                return await conn.QueryFirstOrDefaultAsync<CertificateInfo>(sql, param);
            }
        }

        /// <summary>
        /// Create Certificate
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        public async Task<int> InsertAsync(CertificateInfo entity)
        {
            using (var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;

                sql.AppendLine(" INSERT INTO                          ");
                sql.AppendLine("      dbo.Certificate (Name,          ");
                sql.AppendLine("                    Provider,         ");
                sql.AppendLine("                    OrderIndex,       ");
                sql.AppendLine("                    Status,           ");
                sql.AppendLine("                    StartDate,        ");
                sql.AppendLine("                    EndDate,          ");
                sql.AppendLine("                    CreatedBy,        ");
                sql.AppendLine("                    PersonId)         ");
                sql.AppendLine(" VALUES (@Name,                       ");
                sql.AppendLine("         @Provider,                   ");
                sql.AppendLine("         @OrderIndex,                 ");
                sql.AppendLine("         @Status,                     ");
                sql.AppendLine("         @StartDate,                        ");
                sql.AppendLine("         @EndDate,                        ");
                sql.AppendLine("         @CreatedBy,                      ");
                sql.AppendLine("         @PersonId)                      ");
                
                var param = new
                {
                    collegeName = entity.Name,
                    major = entity.Provider,
                    orderIndex = entity.OrderIndex,
                    status = entity.Status,
                    startDate = entity.StartDate,
                    endDate = entity.EndDate,
                    createdBy = entity.CreatedBy,
                    personId = entity.PersonId
                };
                return await conn.ExecuteAsync(sql.ToString(), param);
            }
        }

        /// <summary>
        /// Update Certificate
        /// </summary>
        /// <param name="entityToUpdate"></param>
        /// <returns></returns>
        public async Task<int> UpdateAsync(CertificateInfo entity)
        {
            using (var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;

                sql.AppendLine(" UPDATE dbo.Certificate                  ");
                sql.AppendLine("    SET Name   = @Name,                    ");
                sql.AppendLine("        Provider  = @Provider,       ");
                sql.AppendLine("        OrderIndex      = @OrderIndex,           ");
                sql.AppendLine("        Status      = @Status,           ");
                sql.AppendLine("        StartDate    = @StartDate          ");
                sql.AppendLine("        EndDate    = @EndDate          ");
                sql.AppendLine("        UpdatedBy    = @UpdatedBy          ");
                sql.AppendLine("        PersonId    = @PersonId          ");
                sql.AppendLine(" WHERE Id = @Id                       ");

                var param = new
                {
                    collegeName = entity.Name,
                    major = entity.Provider,
                    orderIndex = entity.OrderIndex,
                    status = entity.Status,
                    startDate = entity.StartDate,
                    endDate = entity.EndDate,
                    updatedBy = entity.UpdatedBy,
                    personId = entity.PersonId
                };
                return await conn.ExecuteAsync(sql.ToString(), param);
            }
        }

        /// <summary>
        /// Delete Certificate
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<int> DeleteAsync(int id)
        {
            using (var conn = OpenDBConnection())
            {
                var sql = "DELETE FROM dbo.Certificate WHERE Id = @Id";
                return await conn.ExecuteAsync(sql.ToString(), new { Id = id });
            }
        }
    }
}
