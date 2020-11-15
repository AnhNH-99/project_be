using Dapper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebAPI.Models;

namespace WebAPI.Repositories.ProjectTechnologies
{
    public class ProjectTechnologyRepository : RepositoryBase, IProjectTechnologyRepository
    {
        public ProjectTechnologyRepository(string connectionString) : base(connectionString)
        {

        }
        public async Task<int> DeleteAsync(int id)
        {
            using (var conn = OpenDBConnection())
            {
                var sql = "Delete from dbo.ProjectTechnology WHERE ProjectId = @ProjectId";
                return await conn.ExecuteAsync(sql.ToString(), new { ProjectId = id });
            }
        }

        public Task<ProjectTechnology> FindAsync(int id)
        {
            throw new NotImplementedException();
        }

        public Task<IEnumerable<ProjectTechnology>> GetAllAsync()
        {
            throw new NotImplementedException();
        }

        public Task<int> InsertAsync(ProjectTechnology entity)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// get projectId has been inserted lastest
        /// </summary>
        /// <returns></returns>
        public int GetProjectId()
        {
            using(var conn = OpenDBConnection())
            {
                var sql = "select max(id) from dbo.Project";
                return conn.ExecuteScalar<int>(sql);
            }
        }

        /// <summary>
        /// insert value into ProjectTechnology table
        /// if insert value to table after insert to project table, projectId is max(id) of table
        /// if insert value to table for deleting case, projectId is value from object request
        /// </summary>
        /// <param name="entities"></param>
        /// <returns></returns>
        public async Task<int> InsertListTechnologyAsync(List<ProjectTechnology> entities)
        {
            var projectId = 0;
            if (entities.First().ProjectId <= 0)
                projectId = GetProjectId();
            else
            { 
                projectId = entities.First().ProjectId;
            }
            for(int i = 0; i < entities.Count; i++)
            {
                entities[i].CreatedAt = DateTime.Now;
                entities[i].UpdatedAt = DateTime.Now;
                entities[i].CreatedBy = "minh";
                entities[i].UpdatedBy = "minh";
                entities[i].Status = true;
                entities[i].ProjectId = projectId;
            }

            using(var conn = OpenDBConnection())
            {
                var transaction = conn.BeginTransaction();
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;
                sql.AppendLine("Insert into dbo.ProjectTechnology(CreatedAt, CreatedBy, UpdatedAt, UpdatedBy, Status, ProjectId, TechnologyId)");
                sql.AppendLine("values(@CreatedAt, @CreatedBy, @UpdatedAt, @UpdatedBy,@Status,@ProjectId,@TechnologyId)                       ");
                var result = await conn.ExecuteAsync(sql.ToString(), entities, transaction: transaction);
                transaction.Commit();
                return result;

            }
        }

        public Task<int> UpdateAsync(ProjectTechnology entity)
        {
            throw new NotImplementedException();
        }

        public async Task<int> UpdateListTechnologyAsync(List<ProjectTechnology> entities)
        {
            await DeleteAsync(entities.First().ProjectId);
            return await InsertListTechnologyAsync(entities);
        }
    }
}
