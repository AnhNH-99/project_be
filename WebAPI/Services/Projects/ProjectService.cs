using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebAPI.Models;
using WebAPI.Repositories.Projects;
using WebAPI.ViewModels;

namespace WebAPI.Services.Projects
{
    public class ProjectService : BaseService<ProjectViewModel>, IProjectService
    {
        IProjectRepository _projectRepository;

        public ProjectService(IProjectRepository projectRepository)
        {
            this._projectRepository = projectRepository;
        }
        public async Task<ProjectViewModel> DeleteProject(int id)
        {
            await _projectRepository.DeleteAsync(id);
            return model;
        }

        public async Task<IEnumerable<Project>> GetAllProject()
        {
            return await _projectRepository.GetAllAsync();
        }

        public async Task<Project> GetProjectById(int id)
        {
            return await _projectRepository.FindAsync(id);
        }

        public async Task<IEnumerable<Project>> GetProjectByPersonId(int id)
        {
            return await _projectRepository.GetProjectByPersonIdAsync(id);
        }

        public async Task<ProjectViewModel> InsertProject(Project project)
        {
            model.Project = project;
            await _projectRepository.InsertAsync(model.Project);
            return model;

        }

        public async Task<ProjectViewModel> UpdateProject(Project project)
        {
            model.Project = project;
            await _projectRepository.UpdateAsync(model.Project);
            return model;
        }
    }
}
