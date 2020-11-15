using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using WebAPI.Models;
using WebAPI.Services.Projects;
using WebAPI.ViewModels;

namespace WebAPI.Controllers
{
    [Route("api/[controller]/[action]/{id?}")]
    [ApiController]
    public class ProjectController : BaseApiController<ProjectViewModel>
    {
        private IProjectService _projectService;

        public ProjectController(IProjectService projectService)
        {
            this._projectService = projectService;
        }

        [HttpGet]
        public async Task<IActionResult> GetAllProject()
        {
            var projects = await _projectService.GetAllProject();
            apiResult.AppResult.DataResult = projects.AsEnumerable<Project>();
            return Ok(apiResult.AppResult);
        }

        [HttpGet]
        public async Task<IActionResult> GetProjectById(int id)
        {
            var app = await _projectService.GetProjectById(id);
            apiResult.AppResult.DataResult = app;
            return Ok(apiResult.AppResult);
        }

        [HttpGet]

        public async Task<IActionResult> GetProjectByPersonId(int id)
        {
            var app = await _projectService.GetProjectByPersonId(id);
            apiResult.AppResult.DataResult = app;
            return Ok(apiResult.AppResult);
        }
        [HttpPost]
        public async Task<IActionResult> InsertProject([FromBody] Project project)
        {
            var app = await _projectService.InsertProject(project);
            return Ok(app.AppResult);
        }

        [HttpPut]
        public async Task<IActionResult> UpdateProject([FromBody] Project project)
        {
            var app = await _projectService.UpdateProject(project);
            return Ok(app.AppResult);
        }

        [HttpDelete]
        public async Task<IActionResult> DeleteProject(int id)
        {
            var app = await _projectService.DeleteProject(id);
            return Ok(app.AppResult);
        }
    }
}
