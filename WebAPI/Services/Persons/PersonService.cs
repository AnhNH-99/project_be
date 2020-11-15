using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebAPI.Models;
using WebAPI.Repositories.Persons;
using WebAPI.Services.Persons;
using WebAPI.ViewModels;

namespace WebAPI.Services.Persons
{
    public class PersonService : BaseService<PersonViewModel>, IPersonService
    {
        IPersonRepository _personRepository;
        public PersonService(IPersonRepository personRepository)
        {
            this._personRepository = personRepository;
        }



        public async Task<IEnumerable<Person>> GetAllPersonAndSkill()
        {
            return await _personRepository.GetAllPersonAndSkillAsync();
        }
        /// <summary>
        /// Get all Person
        /// </summary>
        /// <returns></returns>
        public async Task<IEnumerable<Person>> GetAllPerson()
        {
            return await _personRepository.GetAllAsync();
        }

        /// <summary>
        /// Get Person by Id
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<Person> GetPersonById(int id)
        {
            return await _personRepository.FindAsync(id);
        }

        /// <summary>
        /// Create Person
        /// </summary>
        /// <param name="person"></param>
        /// <returns></returns>
        public async Task<PersonViewModel> InsertPerson (Person person)
        {
            model.PersonInfo = person;
            var getdate = DateTime.Now.ToString("yyyyMMdd");
            int newestIdPerson = 1;
            if (await _personRepository.AmountOfPerson() > 0)
            {
                newestIdPerson = await _personRepository.GetMaxIdPerson() + 1;
                person.StaffId = string.Format($"{getdate + newestIdPerson}");
            }
            else
            {
                person.StaffId = string.Format($"{getdate + newestIdPerson}");
            }
            await _personRepository.InsertAsync(model.PersonInfo);
            return model;
        }

        /// <summary>
        /// Update Person
        /// </summary>
        /// <param name="person"></param>
        /// <returns></returns>
        public async Task<PersonViewModel> UpdatePerson(Person person)
        {
            model.PersonInfo = person;
            await _personRepository.UpdateAsync(model.PersonInfo);
            return model;
        }

        /// <summary>
        /// Delete Person
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<PersonViewModel> DeletePerson(int id)
        {
            await _personRepository.DeleteAsync(id);
            return model;
        }


        public async Task<IEnumerable<Person>> SearchPersonAndSkillAsync(string FullName, int Location, List<int> TechnologyId)
        {
            return await _personRepository.SearchPersonAndSkillAsync(FullName, Location, TechnologyId);
        }
    }
}
