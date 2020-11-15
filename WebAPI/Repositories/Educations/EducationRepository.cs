using Dapper;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using WebAPI.Models;

namespace WebAPI.Repositories.Educations
{
    public class EducationRepository : RepositoryBase, IEducationRepository
    {
        public EducationRepository(string connectionString) : base(connectionString) { }

        /// <summary>
        /// Get all Education
        /// </summary>
        /// <returns></returns>
        public async Task<IEnumerable<EducationInfo>> GetAllAsync()
        {
            using (var conn = OpenDBConnection())
            {
                var sql = "SELECT * FROM dbo.Education";
                return await conn.QueryAsync<EducationInfo>(sql);
            }
        }

        ///// <summary>
        ///// Check Education
        ///// </summary>
        ///// <param name="conn"></param>
        ///// <param name="userNo"></param>
        ///// <returns></returns>
        //public bool CheckEducation(IDbConnection conn, int userNo)
        //{
        //    StringBuilder sql = new StringBuilder();
        //    sql.Length = 0;
        //    sql.AppendLine("SELECT SUM(CNT)                           ");
        //    sql.AppendLine("FROM (                                    ");
        //    sql.AppendLine("        SELECT                            ");
        //    sql.AppendLine("           COUNT(*) AS CNT                ");
        //    sql.AppendLine("        FROM dbo.Educations                 ");
        //    sql.AppendLine("        WHERE USER_NO = @USER_NO          ");
        //    sql.AppendLine("     )                                    ");

        //    int count = conn.ExecuteScalar<int>(sql.ToString(), new { user_no = userNo });
        //    return count > 0;
        //}

        /// <summary>
        /// Find Education
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<EducationInfo> FindAsync(int id)
        {
            using (var conn = OpenDBConnection())
            {
                var sql = "SELECT * FROM dbo.Education WHERE Id = @_id";
                var param = new { _id = id };
                return await conn.QueryFirstOrDefaultAsync<EducationInfo>(sql, param);
            }
        }

        /// <summary>
        /// Create Education
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        public async Task<int> InsertAsync(EducationInfo entity)
        {
            using (var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;
                
                sql.AppendLine(" INSERT INTO                            ");
                sql.AppendLine("      dbo.Education (CollegeName,          ");
                sql.AppendLine("                    Major,          ");
                sql.AppendLine("                    OrderIndex,           ");
                sql.AppendLine("                    Status,          ");
                sql.AppendLine("                    StartDate,              ");
                sql.AppendLine("                    EndDate,              ");
                sql.AppendLine("                    CreatedBy,            ");
                sql.AppendLine("                    PersonId)            ");
                sql.AppendLine(" VALUES (@CollegeName,                      ");
                sql.AppendLine("         @Major,                    ");
                sql.AppendLine("         @OrderIndex,                     ");
                sql.AppendLine("         @Status,                    ");
                sql.AppendLine("         @StartDate,                        ");
                sql.AppendLine("         @EndDate,                        ");
                sql.AppendLine("         @CreatedBy,                      ");
                sql.AppendLine("         @PersonId)                      ");
                

                var param = new { collegeName = entity.CollegeName,
                    major = entity.Major,
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
        /// Update Education
        /// </summary>
        /// <param name="entityToUpdate"></param>
        /// <returns></returns>
        public async Task<int> UpdateAsync(EducationInfo entity)
        {
            using (var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;

                sql.AppendLine(" UPDATE dbo.Education                  ");
                sql.AppendLine("    SET CollegeName   = @CollegeName,        ");
                sql.AppendLine("        Major  = @Major,       ");
                sql.AppendLine("        OrderIndex      = @OrderIndex,           ");
                sql.AppendLine("        Status      = @Status,           ");
                sql.AppendLine("        StartDate    = @StartDate          ");
                sql.AppendLine("        EndDate    = @EndDate          ");
                sql.AppendLine("        UpdatedBy    = @UpdatedBy          ");
                sql.AppendLine("        PersonId    = @PersonId          ");
                sql.AppendLine(" WHERE Id = @Id              ");

                var param = new
                {
                    collegeName = entity.CollegeName,
                    major = entity.Major,
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
        /// Delete Education
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<int> DeleteAsync(int id)
        {
            using (var conn = OpenDBConnection())
            {
                var sql = "DELETE FROM dbo.Education WHERE Id = @Id";
                return await conn.ExecuteAsync(sql.ToString(), new { Id = id });
            }
        }
    }
}
