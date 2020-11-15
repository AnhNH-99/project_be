using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebAPI.Models;
using WebAPI.ViewModels;

namespace WebAPI.Services.Persons
{
    public interface IPersonService
    {
        Task<IEnumerable<Person>> GetAllPersonAndSkill();

        Task<IEnumerable<Person>> SearchPersonAndSkillAsync(string FullName, int Location, List<int> TechnologyId);

        /// <summary>
        /// Get all Person
        /// </summary>
        /// <returns></returns>
        Task<IEnumerable<Person>> GetAllPerson();

        /// <summary>
        /// Get Person by Id
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        Task<Person> GetPersonById(int id);

        /// <summary>
        /// Create Person
        /// </summary>
        /// <param name="person"></param>
        /// <returns></returns>
        public Task<PersonViewModel> InsertPerson(Person person);

        /// <summary>
        /// Update Person
        /// </summary>
        /// <param name="person"></param>
        /// <returns></returns>
        public Task<PersonViewModel> UpdatePerson(Person person);

        /// <summary>
        /// Delete Person
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public Task<PersonViewModel> DeletePerson(int id);
    }
}
