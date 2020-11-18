using Dapper;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WebAPI.Common;
using WebAPI.Models;
using WebAPI.ViewModels;

namespace WebAPI.Repositories.Persons
{
    public class PersonRepository : RepositoryBase, IPersonRepository
    {
        public PersonRepository(string connectionString) : base(connectionString) { }

        /// <summary>
        /// Get all Person
        /// </summary>
        /// <returns></returns>
        public async Task<IEnumerable<Person>> GetAllAsync()
        {
            using (var conn = OpenDBConnection())
            {
                var sql = "SELECT * FROM dbo.Person";
                return await conn.QueryAsync<Person>(sql);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="id"></param>
        /// <returns></returns>
        public bool CheckPerson(IDbConnection conn, int id)
        {
            StringBuilder sql = new StringBuilder();
            sql.Length = 0;
            sql.AppendLine("SELECT SUM(CNT)                           ");
            sql.AppendLine("FROM (                                    ");
            sql.AppendLine("        SELECT                            ");
            sql.AppendLine("           COUNT(*) AS CNT                ");
            sql.AppendLine("        FROM dbo.Person               ");
            sql.AppendLine("        WHERE Id = @Id                    ");
            sql.AppendLine("     )                                    ");

            int count = conn.ExecuteScalar<int>(sql.ToString(), new { Id = id });
            return count > 0;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<Person> FindAsync(int id)
        {
            using (var conn = OpenDBConnection())
            {
                var sql = "SELECT * FROM dbo.Person WHERE Id = @Id";
                var param = new { Id = id };
                return await conn.QueryFirstOrDefaultAsync<Person>(sql, param);
            }
        }

        #region Handling insert info Person
        public async Task<int> AmountOfPerson()
        {
            using(var conn =  OpenDBConnection())
            {
                var sql = "select count(*) from Person";
                return await conn.ExecuteScalarAsync<int>(sql);
            }
        }

        public async Task<int> GetMaxIdPerson()
        { 
            using (var conn = OpenDBConnection())
            {
                var sql = "select max(id) from Person";
                return await conn.ExecuteScalarAsync<int>(sql);
            }
        }
        #endregion
        /// <summary>
        /// Create Person
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        public async Task<int> InsertAsync(Person entity)
        {
            using (var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;

                sql.AppendLine(" INSERT INTO                            ");
                sql.AppendLine("      dbo.Person (StaffId,          ");
                sql.AppendLine("                    FullName,           ");
                sql.AppendLine("                    Email,              ");
                sql.AppendLine("                    Location,           ");
                sql.AppendLine("                    Avatar,             ");
                sql.AppendLine("                    Description,        ");
                sql.AppendLine("                    Phone,              ");
                sql.AppendLine("                    YearOfBirth,        ");
                sql.AppendLine("                    Gender,             ");
                sql.AppendLine("                    Status,             ");
                sql.AppendLine("                    CreatedAt,          ");
                sql.AppendLine("                    CreatedBy,          ");
                sql.AppendLine("                    UpdatedAt,          ");
                sql.AppendLine("                    UpdatedBy)          ");
                sql.AppendLine(" VALUES (@StaffId,                      ");
                sql.AppendLine("         @FullName,                     ");
                sql.AppendLine("         @Email,                        ");
                sql.AppendLine("         @Location,                     ");
                sql.AppendLine("         @Avatar,                       ");
                sql.AppendLine("         @Description,                  ");
                sql.AppendLine("         @Phone,                        ");
                sql.AppendLine("         @YearOfBirth,                  ");
                sql.AppendLine("         @Gender,                       ");
                sql.AppendLine("         @Status,                       ");
                sql.AppendLine("         getdate(),                     ");
                sql.AppendLine("         @CreatedBy,                    ");
                sql.AppendLine("         getdate(),                     ");
                sql.AppendLine("         @UpdatedBy)                    ");

                var param = new
                {
                    staffId = entity.StaffId,
                    fullName = entity.FullName,
                    email = entity.Email.ToString(),
                    location = entity.Location,
                    avatar = entity.Avatar,
                    description = entity.Description,
                    phone = entity.Phone,
                    yearOfBirth = entity.YearOfBirth.ToString(),
                    gender = entity.Gender,
                    status = entity.Status,
                    createdBy = entity.CreatedBy,
                    updatedBy = entity.UpdatedBy
                };
                return await conn.ExecuteAsync(sql.ToString(), param);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        public async Task<int> UpdateAsync(Person entity)
        {
            using (var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;

                sql.AppendLine(" UPDATE dbo.Person                         ");
                sql.AppendLine("    SET StaffId         = @StaffId,        ");
                sql.AppendLine("        FullName        = @FullName,       ");
                sql.AppendLine("        Email           = @Email,          ");
                sql.AppendLine("        Location        = @Location,       ");
                sql.AppendLine("        Avatar          = @Avatar,         ");
                sql.AppendLine("        Description     = @Description,    ");
                sql.AppendLine("        Phone           = @Phone,          ");
                sql.AppendLine("        YearOfBirth     = @YearOfBirth,    ");
                sql.AppendLine("        Gender          = @Gender,         ");
                sql.AppendLine("        Status          = @Status,         ");
                sql.AppendLine("        CreatedAt       = getdate(),       ");
                sql.AppendLine("        CreatedBy       = @CreatedBy,      ");
                sql.AppendLine("        UpdatedAt       = getdate(),       ");
                sql.AppendLine("        UpdatedBy       = @UpdatedBy       ");
                sql.AppendLine(" WHERE Id = @Id                            ");

                var param = new
                {
                    staffId = entity.StaffId,
                    fullName = entity.FullName.ToString(),
                    email = entity.Email.ToString(),
                    location = entity.Location,
                    avatar = entity.Avatar,
                    description = entity.Description,
                    phone = Convert.ToInt32(entity.Phone),
                    yearOfBirth = Convert.ToDateTime(entity.YearOfBirth),
                    gender = entity.Gender,
                    status=entity.Status,
                    createdAt = Convert.ToDateTime(entity.CreatedAt),
                    createdBy=entity.CreatedBy.ToString(),
                    updatedAt = Convert.ToDateTime(entity.UpdatedAt),
                    updatedBy=entity.UpdatedBy.ToString(),
                    id = entity.Id
                };
                return await conn.ExecuteAsync(sql.ToString(), param);
            }
        }

        /// <summary>
        /// Delete Person
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<int> DeleteAsync(int id)
        {
            using (var conn = OpenDBConnection())
            {
                var sql = "DELETE FROM dbo.Person WHERE Id = @Id";
                return await conn.ExecuteAsync(sql.ToString(), new { Id = id });
            }
        }

        public async Task<IEnumerable<Person>> SearchPersonAndSkillAsync(string FullName, int Location, List<int> TechnologyId)
        {
            using (var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;

                sql.AppendLine("SELECT pe.Id, pe.FullName, pe.Location, ca.Id, ca.Name, te.Id, te.Name, te.CategoryId                                                 ");
                sql.AppendLine("FROM (Select * from Person where FullName like '%"+ FullName + "%' and Location = @location) AS pe                    ");
                sql.AppendLine("INNER JOIN persontechnology AS pt                                                                                                                                     ");
                sql.AppendLine("ON pe.id = pt.PersonId                                                                                                                                                                 ");
                sql.AppendLine("INNER JOIN (Select * from technology where Id in @technologyIds) AS te                                                                ");
                sql.AppendLine("ON te.id = pt.TechnologyId                                                                                                                                               ");
                sql.AppendLine("INNER JOIN category AS ca                                                                                                                                              ");
                sql.AppendLine("ON ca.Id = te.CategoryId                                                                                                                                        ");
                sql.AppendLine("ORDER BY ca.Id ASC                                                                                                                                                ");

                var lookup = new Dictionary<int, Person>();
                var param = new
                {
                    location = Location,
                    technologyIds = TechnologyId
                };

                await conn.QueryAsync<Person, Category, Technology, Person>(sql.ToString(), map: (pe, ca, te) =>
                {
                    Person person;
                    if (!lookup.TryGetValue(pe.Id, out person))
                        lookup.Add(pe.Id, person = pe);

                    if (person.Categories == null)
                        person.Categories = new List<Category>();
                    int i = 0;
                    foreach (var item in person.Categories)
                    {
                        if (item.Id == ca.Id)
                            i++;
                    }
                    if (i <= 0)
                        person.Categories.Add(ca);

                    foreach (var item in person.Categories)
                    {
                        if (item.Technologies == null)
                            item.Technologies = new List<Technology>();
                        if (item.Id == te.CategoryId)
                        {
                            item.Technologies.Add(te);
                        }
                    }
                    return person;
                },param);
                var list = lookup.Values.ToList();
                return list;
            }
        }

        public async Task<IEnumerable<Person>> GetAllPersonAndSkillAsync()
        {
            using (var conn = OpenDBConnection())
            {
                StringBuilder sql = new StringBuilder();
                sql.Length = 0;

                sql.AppendLine("    SELECT pe.Id, pe.FullName, ca.Id, ca.Name, te.Id, te.Name, te.CategoryId        ");
                sql.AppendLine("    FROM person AS pe                                                               ");
                sql.AppendLine("    INNER JOIN persontechnology AS pt                                               ");
                sql.AppendLine("    ON pe.id = pt.PersonId                                                          ");
                sql.AppendLine("    INNER JOIN technology AS te                                                     ");
                sql.AppendLine("    ON te.id = pt.TechnologyId                                                      ");
                sql.AppendLine("    INNER JOIN category AS ca                                                       ");
                sql.AppendLine("    ON ca.Id = te.CategoryId                                                        ");
                sql.AppendLine("    ORDER BY ca.Id ASC                                                              ");
                var lookup = new Dictionary<int, Person>();
                await conn.QueryAsync<Person, Category, Technology, Person>(sql.ToString(), map: (pe, ca, te) =>
                 {
                     Person person;
                     if (!lookup.TryGetValue(pe.Id, out person))
                         lookup.Add(pe.Id, person = pe);

                     if (person.Categories == null)
                         person.Categories = new List<Category>();
                     int i = 0;
                     foreach(var item in person.Categories)
                     {
                         if (item.Id == ca.Id)
                             i++;
                     }
                     if(i<= 0)
                        person.Categories.Add(ca);

                     foreach (var item in person.Categories)
                     {
                         if (item.Technologies == null)
                             item.Technologies = new List<Technology>();
                         if (item.Id == te.CategoryId)
                         {
                             item.Technologies.Add(te);
                         }
                     }
                     return person;
                 });
                var list = lookup.Values.ToList();
                return list;
            }
        }
    }
}

