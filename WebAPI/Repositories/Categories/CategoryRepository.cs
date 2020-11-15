using Dapper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebAPI.Models;
using WebAPI.ViewModels;

namespace WebAPI.Repositories.Categories
{
    public class CategoryRepository : RepositoryBase, ICategoryRepository
    {

        public CategoryRepository(string connectionString) : base(connectionString) { }

        
        /// <summary>
        /// Get all Category
        /// </summary>
        /// <returns></returns>
        public async Task<IEnumerable<Category>> GetAllAsync()
        {
            using (var conn = OpenDBConnection())
            {
                var sql = "SELECT CA.Id, CA.Name FROM Category AS CA WHERE CA.Status = 1";
                return await conn.QueryAsync<Category>(sql);
            }
        }

        /// <summary>
        /// Create Category
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        public async Task<int> InsertAsync(Category entity)
        {
            using (var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;
                sql.AppendLine(" INSERT INTO                      ");
                sql.AppendLine("      Category (CreatedBy,        ");
                sql.AppendLine("                UpdatedBy,         ");
                sql.AppendLine("                Name )             ");
                sql.AppendLine(" VALUES (@CreatedBy,              ");
                sql.AppendLine("         @UpdatedBy,               ");
                sql.AppendLine("         @Name )                   ");
                var param = new
                {
                    CreatedBy = entity.CreatedBy,
                    UpdatedBy = entity.UpdatedBy,
                    Name = entity.Name
                };
                return await conn.ExecuteAsync(sql.ToString(), param);
            }
        }

        /// <summary>
        /// Update Category
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        public async Task<int> UpdateAsync(Category entity)
        {
            using (var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;
                sql.AppendLine(" UPDATE Category                    ");
                sql.AppendLine("    SET Name       = @Name,         ");
                sql.AppendLine("        UpdatedAt  = @UpdatedAt,    ");
                sql.AppendLine("        UpdatedBy   = @UpdatedBy    ");
                sql.AppendLine(" WHERE Id = @Id                     ");
                entity.UpdatedAt = DateTime.Now;
                var param = new
                {
                    Name = entity.Name,
                    UpdatedAt = entity.UpdatedAt,
                    UpdatedBy = entity.UpdatedBy,
                    Id = entity.Id
                };
                return await conn.ExecuteAsync(sql.ToString(), param);
            }
        }

        /// <summary>
        /// Delete Category
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<int> DeleteAsync(int id)
        {
            using (var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;
                sql.AppendLine(" UPDATE Category                  ");
                sql.AppendLine("    SET Status     = @Status,     ");
                sql.AppendLine("        UpdatedAt  = @UpdatedAt,  ");
                sql.AppendLine("        UpdatedBy  = @UpdatedBy   ");
                sql.AppendLine(" WHERE Id = @Id                   ");
                var param = new
                {
                    Status = 0,
                    UpdatedAt = DateTime.Now,
                    UpdatedBy = WebAPI.Helpers.HttpContext.CurrentUser,
                    Id = id
                };
                return await conn.ExecuteAsync(sql.ToString(), param);
            }
        }

        /// <summary>
        /// Check Category
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public async Task<bool> CheckCategoryAsync(string name)
        {
            using(var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;
                sql.Append("SELECT COUNT(*)            ");
                sql.Append("FROM Category              ");
                sql.Append("WHERE Name = @Name         ");
                sql.Append("AND Status = 1             ");
                int count = await conn.ExecuteScalarAsync<int>(sql.ToString(), new { Name = name });
                return count > 0;
            }
        }

        public Task<Category> FindAsync(int id)
        {
            throw new NotImplementedException();
        }

        public async Task<Category> CheckCategoryOnDeleteAsync(string name)
        {
            using (var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;
                sql.Append("SELECT *                  ");
                sql.Append("FROM Category             ");
                sql.Append("WHERE Name = @Name        ");
                sql.Append("AND Status = 0            ");
                Category result = await conn.QueryFirstOrDefaultAsync<Category>(sql.ToString(), new { Name = name });
                return result;
            }
        }

        public async Task<int> UpdateCategoryToInsert(Category category)
        {
            using (var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;
                sql.AppendLine(" UPDATE Category                  ");
                sql.AppendLine("    SET Status     = 1,           ");
                sql.AppendLine("        UpdatedAt  = @UpdatedAt,  ");
                sql.AppendLine("        UpdatedBy  = @UpdatedBy   ");
                sql.AppendLine(" WHERE Id = @Id                   ");
                var param = new
                {
                    Status = 1,
                    UpdatedAt = DateTime.Now,
                    UpdatedBy = category.UpdatedBy,
                    Id = category.Id
                };
                return await conn.ExecuteAsync(sql.ToString(), param);
            }
        }
    }
}
