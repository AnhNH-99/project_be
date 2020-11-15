using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebAPI.Models;
using WebAPI.Repositories.Interfaces;

namespace WebAPI.Repositories.ProjectTechnologies
{
    public interface IProjectTechnologyRepository : IRepositoryBase<ProjectTechnology>
    {
        /// <summary>
        /// insert list technology
        /// </summary>
        /// <param name="entities"></param>
        /// <returns></returns>
        public Task<int> InsertListTechnologyAsync(List<ProjectTechnology> entities);

        public Task<int> UpdateListTechnologyAsync(List<ProjectTechnology> entities);
    }
}
