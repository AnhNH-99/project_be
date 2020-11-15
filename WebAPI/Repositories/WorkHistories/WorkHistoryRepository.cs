using Dapper;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using WebAPI.Models;

namespace WebAPI.Repositories.WorkHistories
{
    public class WorkHistoryRepository : RepositoryBase, IWorkHistoryRepository
    {
        public WorkHistoryRepository(string connectionString) : base(connectionString) { }

        /// <summary>
        /// Get all WorkHistory
        /// </summary>
        /// <returns></returns>
        public async Task<IEnumerable<WorkHistoryInfo>> GetAllAsync()
        {
            using (var conn = OpenDBConnection())
            {
                var sql = "SELECT * FROM dbo.WorkHistory ";
                return await conn.QueryAsync<WorkHistoryInfo>(sql);
            }
        }

        /// <summary>
        /// Get WorkHistory By Id
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<WorkHistoryInfo> FindAsync(int id)
        {
            using (var conn = OpenDBConnection())
            {
                var sql = "SELECT * FROM dbo.WorkHistory WHERE Id = @_id";
                var param = new { _id = id };
                return await conn.QueryFirstOrDefaultAsync<WorkHistoryInfo>(sql, param);
            }
        }

        /// <summary>
        /// Get WorkHistory By Id
        /// </summary>
        /// <param name="personId"></param>
        /// <returns></returns>
        public async Task<IEnumerable<WorkHistoryInfo>> FindAsyncByPersonId(int id)
        {
            using (var conn = OpenDBConnection())
            {
                var sql = "SELECT * FROM dbo.WorkHistory WHERE PersonId = @personId";
                var param = new { personId = id };
                var results = await conn.QueryAsync<WorkHistoryInfo>(sql, param);
                return results;
            }
        }

        /// <summary>
        /// Create WorkHistory
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        public async Task<int> InsertAsync(WorkHistoryInfo entity)
        {
            using (var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;

                sql.AppendLine(" INSERT INTO                             ");
                sql.AppendLine("      dbo.WorkHistory (Position,         ");
                sql.AppendLine("                   CompanyName,          ");
                sql.AppendLine("                   OrderIndex,           ");
                sql.AppendLine("                   StartDate,            ");
                sql.AppendLine("                   EndDate,              ");
                sql.AppendLine("                   PersonId,             ");
                sql.AppendLine("                   CreatedBy)            ");
                sql.AppendLine(" VALUES (@Position,                      ");
                sql.AppendLine("         @CompanyName,                   ");
                sql.AppendLine("         @OrderIndex,                    ");
                sql.AppendLine("         @StartDate,                     ");
                sql.AppendLine("         @EndDate,                       ");
                sql.AppendLine("         @PersonId,                      ");
                sql.AppendLine("         @CreatedBy)                     ");



                var param = new
                {
                    position = entity.Position,
                    companyName = entity.CompanyName,
                    orderIndex = entity.OrderIndex,
                    status = entity.Status,
                    startDate = entity.StartDate,
                    endDate = entity.EndDate,
                    personId = entity.PersonId,
                    createdBy = entity.CreatedBy
                };
                return await conn.ExecuteAsync(sql.ToString(), param);
                
            }
        }

        /// <summary>
        /// Update WorkHistory
        /// </summary>
        /// <param name="entityToUpdate"></param>
        /// <returns></returns>
        public async Task<int> UpdateAsync(WorkHistoryInfo entity)
        {
            using (var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;

                sql.AppendLine(" UPDATE dbo.WorkHistory                           ");
                sql.AppendLine("    SET Position   = @Position,                   ");
                sql.AppendLine("        CompanyName  = @CompanyName,              ");
                sql.AppendLine("        OrderIndex      = @OrderIndex,            ");
                sql.AppendLine("        StartDate      = @StartDate,              ");
                sql.AppendLine("        EndDate    = @EndDate,                    ");
                sql.AppendLine("        PersonId    = @PersonId,                  ");
                sql.AppendLine("        Status    = @Status,                      ");
                sql.AppendLine("        UpdatedBy    = @UpdatedBy                 ");
                sql.AppendLine(" WHERE Id = @Id                                   ");

                entity.UpdatedAt = DateTime.Now;

                var param = new
                {
                    id = entity.Id,
                    position = entity.Position,
                    companyName = entity.CompanyName,
                    orderIndex = entity.OrderIndex,
                    startDate = entity.StartDate,
                    endDate = entity.EndDate,
                    personId = entity.PersonId,
                    status = entity.Status,
                    updatedBy = entity.UpdatedBy
                };
                return await conn.ExecuteAsync(sql.ToString(), param);
            }
        }

        /// <summary>
        /// Delete WorkHistory
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<int> DeleteAsync(int id)
        {
            using (var conn = OpenDBConnection())
            {
                var sql = "DELETE FROM dbo.WorkHistory WHERE Id = @Id";
                return await conn.ExecuteAsync(sql.ToString(), new { Id = id });
            }
        }
    }
}
