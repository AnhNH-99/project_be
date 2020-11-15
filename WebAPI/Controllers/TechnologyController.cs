using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using WebAPI.Models;
using WebAPI.Services.Technologies;
using WebAPI.ViewModels;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace WebAPI.Controllers
{
    [Route("api/[controller]/[action]/{id?}")]
    [ApiController]
    public class TechnologyController : BaseApiController<TechnologyViewModel>
    {
        private ITechnologyService _technologyService;
        public TechnologyController(ITechnologyService technologyService)
        {
            this._technologyService = technologyService;
        }

        /// <summary>
        /// Get all Technology
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [HttpGet]
        public async Task<IActionResult> GetTechnology(int? id)
        {
            var technologies = await _technologyService.GetTechnology(id);
            apiResult.AppResult.DataResult = technologies.AsEnumerable<Technology>();
            return Ok(apiResult.AppResult);
        }

        /// <summary>
        /// Create Technology
        /// </summary>
        /// <param name="technology"></param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> InsertTechnology(Technology technology)
        {
            var app = await _technologyService.InsertTechnology(technology);
            return Ok(app.AppResult);
        }

        /// <summary>
        /// Update Technology
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        [HttpPut]
        public async Task<IActionResult> UpdateTechnology(Technology technology)
        {
            var app = await _technologyService.UpdateTechnology(technology);
            return Ok(app.AppResult);
        }

        /// <summary>
        /// Delete Technology
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [HttpDelete]
        public async Task<IActionResult> DeleteTechnology(int id)
        {
            var app = await _technologyService.DeleteTechnology(id);
            return Ok(app.AppResult);
        }
    }
}
