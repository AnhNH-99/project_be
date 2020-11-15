using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebAPI.Models;
using WebAPI.Repositories.ProjectTechnologies;
using WebAPI.ViewModels;

namespace WebAPI.Services.ProjectTechnologies
{
    public class ProjectTechnologyService : BaseService<ProjectTechnologyViewModel>, IProjectTechnologyService
    {
        IProjectTechnologyRepository _projectTechnologyRepository;

        public ProjectTechnologyService(IProjectTechnologyRepository projectTechnologyRepository)
        {
            this._projectTechnologyRepository = projectTechnologyRepository;
        }

        public async Task<ProjectTechnologyViewModel> DeleteProjectTechnology(int id)
        {
            await _projectTechnologyRepository.DeleteAsync(id);
            return model;
        }

        public async Task<IEnumerable<ProjectTechnology>> InsertListTechnologyAsync(List<ProjectTechnology> entities)
        {
            List<object> lst = new List<object>();

            foreach (var item in entities)
            {
                model.ProjectTechnology = item;
                lst.Add(model.ProjectTechnology);
            }

            await _projectTechnologyRepository.InsertListTechnologyAsync(entities);
            return lst as IEnumerable<ProjectTechnology>;
        }


        public async Task<IEnumerable<ProjectTechnology>> UpdateListTechnologyAsync(List<ProjectTechnology> entities)
        {
            List<object> lst = new List<object>();

            foreach (var item in entities)
            {
                model.ProjectTechnology = item;
                lst.Add(model.ProjectTechnology);
            }

            await _projectTechnologyRepository.UpdateListTechnologyAsync(entities);
            return lst as IEnumerable<ProjectTechnology>;
        }
    }
}
