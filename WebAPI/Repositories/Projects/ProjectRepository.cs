using Dapper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebAPI.Models;

namespace WebAPI.Repositories.Projects
{
    public class ProjectRepository : RepositoryBase, IProjectRepository
    {
        public ProjectRepository(string connectionString) : base(connectionString)
        {

        }
        /// <summary>
        /// delete project by updating status = 0
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<int> DeleteAsync(int id)
        {
            using(var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;

                sql.AppendLine("Update dbo.Project                            ");
                sql.AppendLine("   set UpdatedAt = @UpdatedAt,                ");
                sql.AppendLine("       UpdatedBy = @UpdatedBy,                ");
                sql.AppendLine("       Status = 0                             ");
                sql.AppendLine("    where Id = @Id                            ");

                var param = new
                {
                    UpdatedAt = DateTime.Now,
                    UpdatedBy = "minh",
                    Id = id
                };
                return await conn.ExecuteAsync(sql.ToString(), param);
            }
        }

        public async Task<Project> FindAsync(int id)
        {
            using (var conn = OpenDBConnection())
            {
                var sql = "Select * from dbo.Project where Id = @Id";
                var param = new { Id = id };
                return await conn.QueryFirstOrDefaultAsync<Project>(sql, param);
            }
        }

        public async Task<IEnumerable<Project>> GetAllAsync()
        {
            using(var conn = OpenDBConnection())
            {
                var sql = "Select * from dbo.Project";
                return await conn.QueryAsync<Project>(sql);
            }
        }
        /// <summary>
        /// get project by personId
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<IEnumerable<Project>> GetProjectByPersonIdAsync(int id)
        {
            using(var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;

                sql.AppendLine("Select PJ.*, TLG.*                                       ");
                sql.AppendLine("from dbo.ProjectTechnology as PT                         ");
                sql.AppendLine("inner join dbo.Project as PJ                             ");
                sql.AppendLine("on PT.ProjectId = PJ.Id                                  ");
                sql.AppendLine("inner join dbo.Technology as TLG                         ");
                sql.AppendLine("on PT.TechnologyId = TLG.Id                              ");
                sql.AppendLine("where PJ.PersonId = @PersonId                            ");
                sql.AppendLine("and PJ.Status = 1 and PT.Status = 1 and TLG.Status = 1   ");

                var param = new { PersonId = id };
                var lookup = new Dictionary<int, Project>();

                await conn.QueryAsync<Project, Technology, Project>(sql.ToString(), (PJ, TLG) =>
                {
                    Project project;
                    if (!lookup.TryGetValue(PJ.Id, out project))
                    {
                        lookup.Add(PJ.Id, project = PJ);
                    }
                    if (project.Technologies == null)
                        project.Technologies = new List<Technology>();
                    project.Technologies.Add(TLG);
                    return project;
                }, param);

                var result = lookup.Values.ToList();
                return result;
            }
        }

        public async Task<int> InsertAsync(Project entity)
        {
            using(var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;

                sql.AppendLine("Insert into                          ");
                sql.AppendLine("       dbo.Project(CreatedAt,        ");
                sql.AppendLine("       CreatedBy, Status,            ");
                sql.AppendLine("       StartDate, EndDate,           ");
                sql.AppendLine("       Name, Description,            ");
                sql.AppendLine("       Position, Responsibilities,   ");
                sql.AppendLine("       TeamSize, OrderIndex,         ");
                sql.AppendLine("       PersonId)                     ");
                sql.AppendLine("values(@CreatedAt, @CreatedBy,       ");
                sql.AppendLine("       @Status, @StartDate,          ");
                sql.AppendLine("       @EndDate, @Name,              ");
                sql.AppendLine("       @Description, @Position,      ");
                sql.AppendLine("       @Responsibilities, @TeamSize, ");
                sql.AppendLine("       @OrderIndex, @PersonId)       ");

                var param = new
                {
                    createdAt = entity.CreatedAt,
                    createdBy = entity.CreatedBy,
                    status = entity.Status,
                    startDate = entity.StartDate,
                    endDate = entity.EndDate,
                    name = entity.Name,
                    description = entity.Description,
                    position = entity.Position,
                    responsibilities = entity.Responsibilities,
                    teamSize = entity.TeamSize,
                    orderIndex = entity.OrderIndex,
                    personId = entity.PersonId
                };
                return await conn.ExecuteAsync(sql.ToString(), param);
            }
        }

        public async Task<int> UpdateAsync(Project entity)
        {
            using(var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;

                sql.AppendLine("Update dbo.Project                            ");
                sql.AppendLine("   set UpdatedAt = @UpdatedAt,                ");
                sql.AppendLine("       UpdatedBy = @UpdatedBy,                  ");
                sql.AppendLine("       Status = @Status,                      ");
                sql.AppendLine("       StartDate = @StartDate,                ");
                sql.AppendLine("       EndDate = @EndDate,                    ");
                sql.AppendLine("       Name = @Name,                          ");
                sql.AppendLine("       Description = @Description,            ");
                sql.AppendLine("       Position = @Position,                  ");
                sql.AppendLine("       Responsibilities = @Responsibilities,  ");
                sql.AppendLine("       TeamSize = @TeamSize,                  ");
                sql.AppendLine("       OrderIndex = @OrderIndex               ");
                sql.AppendLine("    where Id = @Id                            ");

                var param = new
                {
                    updatedAt = entity.UpdatedAt,
                    updateBy = entity.UpdatedBy,
                    status = entity.Status,
                    startDate = entity.StartDate,
                    endDate = entity.EndDate,
                    name = entity.Name,
                    description = entity.Description,
                    position = entity.Position,
                    responsibilities = entity.Responsibilities,
                    teamSize = entity.TeamSize,
                    orderIndex = entity.OrderIndex,
                    personId = entity.PersonId,
                    id = entity.Id
                };
                return await conn.ExecuteAsync(sql.ToString(), param);
            }
        }
    }
}
