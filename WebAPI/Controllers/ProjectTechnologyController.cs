using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using WebAPI.Models;
using WebAPI.Services.ProjectTechnologies;
using WebAPI.ViewModels;

namespace WebAPI.Controllers
{
    [Route("api/[controller]/[action]/{id?}")]
    [ApiController]
    public class ProjectTechnologyController : BaseApiController<ProjectTechnologyViewModel>
    {
        public IProjectTechnologyService _projectTechnologyService;

        public ProjectTechnologyController(IProjectTechnologyService projectTechnologyService)
        {
            this._projectTechnologyService = projectTechnologyService;
        }

        [HttpPost]

        public async Task<IActionResult> InsertListTechnology(List<ProjectTechnology> entities)
        {
            var tech = await _projectTechnologyService.InsertListTechnologyAsync(entities);
            apiResult.AppResult.DataResult = tech.AsEnumerable<ProjectTechnology>();
            return Ok(apiResult.AppResult);
        }

        [HttpPut]

        public async Task<IActionResult> UpdateListTechnology(int id, List<ProjectTechnology> entities)
        {
            var tech = await _projectTechnologyService.UpdateListTechnologyAsync(entities);
            apiResult.AppResult.DataResult = tech.AsEnumerable<ProjectTechnology>();
            return Ok(apiResult.AppResult);
        }
    }
}
