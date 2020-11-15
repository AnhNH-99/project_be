using Dapper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebAPI.Models;

namespace WebAPI.Repositories.Technologies
{
    public class TechnologyRepository : RepositoryBase, ITechnologyRepository
    {
        public TechnologyRepository(string connectionString):base(connectionString){}

        /// <summary>
        /// Get Technology
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<IEnumerable<Technology>> GetTechnologyAsync(int? id)
        {
            using (var conn = OpenDBConnection())
            {
                if(id == null)
                {
                    var sql = "SELECT TE.ID, TE.NAME, TE.CategoryId  FROM Technology AS TE WHERE TE.Status = 1";
                    return await conn.QueryAsync<Technology>(sql);
                }
                else
                {
                    StringBuilder sql = new StringBuilder();
                    sql.Length = 0;
                    sql.Append("SELECT TE.ID, TE.NAME, TE.CategoryId    ");
                    sql.Append("FROM Technology AS TE                   ");
                    sql.Append("INNER JOIN Category AS CA               ");
                    sql.Append("ON TE.CategoryID = CA.ID                ");
                    sql.Append("WHERE TE.Status = 1 AND CA.Id = @Id     ");
                    return await conn.QueryAsync<Technology>(sql.ToString(), new { Id = id});
                }
                
            }
        }

        /// <summary>
        /// Create Technology
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        public async Task<int> InsertAsync(Technology entity)
        {
            using (var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;
                sql.Append("INSERT INTO                          ");
                sql.Append("       Technology(CreatedBy,         ");
                sql.Append("                  UpdatedBy,         ");
                sql.Append("                  Name,              ");
                sql.Append("                  CategoryId)        ");
                sql.Append("VALUES (@CreatedBy,                  ");
                sql.Append("        @UpdatedBy,                  ");
                sql.Append("        @Name,                       ");
                sql.Append("        @CategoryId)                 ");
                var param = new
                {
                    CreatedBy = entity.CreatedBy,
                    UpdatedBy = entity.UpdatedBy,
                    Name = entity.Name,
                    CategoryId = entity.CategoryId
                };
                return await conn.ExecuteAsync(sql.ToString(), param);
            }
        }

        /// <summary>
        /// Update Technology
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        public async Task<int> UpdateAsync(Technology entity)
        {
            using (var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;
                sql.Append("UPDATE Technology                        ");
                sql.Append("      SET UpdatedAt   = @UpdatedAt,      ");
                sql.Append("          UpdatedBy    = @UpdatedBy,     ");
                sql.Append("          Name        = @Name,           ");
                sql.Append("          CategoryId  = @CategoryId      ");
                sql.Append("WHERE Id = @Id                           ");
                entity.UpdatedAt = DateTime.Now;
                var param = new
                {
                    UpdatedAt = entity.UpdatedAt,
                    UpdatedBy = entity.UpdatedBy,
                    Name = entity.Name,
                    CategoryId = entity.CategoryId,
                    Id = entity.Id
                };
                return await conn.ExecuteAsync(sql.ToString(), param);
            }
        }

        /// <summary>
        /// Delete Technology
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<int> DeleteAsync(int id)
        {
            using (var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;
                sql.Append("UPDATE Technology                        ");
                sql.Append("      SET UpdatedAt   = @UpdatedAt,      ");
                sql.Append("          UpdatedBy    = @UpdatedBy,     ");
                sql.Append("          Status      = 0                ");
                sql.Append("WHERE Id = @Id                           ");
                var param = new
                {
                    UpdatedAt = DateTime.Now,
                    UpdatedBy = WebAPI.Helpers.HttpContext.CurrentUser,
                    Id = id
                };
                return await conn.ExecuteAsync(sql.ToString(), param);
            }
        }

        public Task<IEnumerable<Technology>> GetAllAsync()
        {
            throw new NotImplementedException();
        }

        public Task<Technology> FindAsync(int id)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Check Technology
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public async Task<bool> CheckTechnologyAsync(Technology technology)
        {
            using(var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;
                sql.Append("SELECT COUNT(*)                           ");
                sql.Append("FROM Technology as TE                     ");
                sql.Append("WHERE TE.CategoryId = @CategoryId         ");
                sql.Append("AND TE.Name = @Name AND TE.Status = 1     ");
                var param = new
                {
                    CategoryId = technology.CategoryId,
                    Name = technology.Name
                };
                int count = await conn.ExecuteScalarAsync<int>(sql.ToString(), param);
                return count > 0;
            }
        }

        /// <summary>
        /// Check Technology on delete
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public async Task<Technology> CheckTechnologyOnDeleteAsync(Technology technology)
        {
            using (var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;
                sql.Append("SELECT *                                 ");
                sql.Append("FROM Technology as TE                    ");
                sql.Append("WHERE TE.CategoryId = @CategoryId        ");
                sql.Append("AND TE.Name = @Name AND TE.Status = 0    ");
                var param = new
                {
                    CategoryId = technology.CategoryId,
                    Name = technology.Name
                };
                Technology result = await conn.QueryFirstOrDefaultAsync<Technology>(sql.ToString(), param);
                return result;
            }
        }

        /// <summary>
        /// Update Technology to Insert
        /// </summary>
        /// <param name="categoryId"></param>
        /// <returns></returns>
        public async Task<int> UpdateTechnologyToInsert(Technology technology)
        {
            using (var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;
                sql.Append("UPDATE Technology                        ");
                sql.Append("      SET UpdatedAt   = @UpdatedAt,      ");
                sql.Append("          UpdatedBy    = @UpdatedBy,      ");
                sql.Append("          Status      = 1                ");
                sql.Append("WHERE Id = @Id                           ");
                var param = new
                {
                    UpdatedAt = DateTime.Now,
                    UpdatedBy = technology.UpdatedBy,
                    Id = technology.Id
                };
                return await conn.ExecuteAsync(sql.ToString(), param);
            }
        }
    }
}
