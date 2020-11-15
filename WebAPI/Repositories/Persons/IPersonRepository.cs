using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebAPI.Repositories.Interfaces;
using WebAPI.Models;
using WebAPI.ViewModels;

namespace WebAPI.Repositories.Persons
{
    public interface IPersonRepository: IRepositoryBase<Person>
    {
        public Task<int> AmountOfPerson();
        public Task<int> GetMaxIdPerson();

        public Task<IEnumerable<Person>> GetAllPersonAndSkillAsync();

        public Task<IEnumerable<Person>> SearchPersonAndSkillAsync(string FullName, int Location, List<int> TechnologyId);
    }
}
