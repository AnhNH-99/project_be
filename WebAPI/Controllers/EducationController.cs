using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using WebAPI.Models;
using WebAPI.Services.Educations;
using WebAPI.ViewModels;

namespace WebAPI.Controllers
{
    [Route("api/[controller]/[action]/{id?}")]
    [ApiController]
    public class EducationController : BaseApiController<EducationViewModel>
    {
        private IEducationService _educationService;
        public EducationController(IEducationService educationService)
        {
            this._educationService = educationService;
        }

        /// <summary>
        /// Get all Education
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public async Task<IActionResult> GetAllEducation()
        {
            var employees = await _educationService.GetAllEducation();
            apiResult.AppResult.DataResult = employees.AsEnumerable<EducationInfo>();
            return Ok(apiResult.AppResult);
        }

        /// <summary>
        /// Get Education by Id
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [HttpGet]
        public async Task<IActionResult> GetEducationById(int id)
        {
            var app = await _educationService.GetEducationById(id);
            apiResult.AppResult.DataResult = app;
            return Ok(apiResult.AppResult);
        }

        /// <summary>
        /// Create Education
        /// </summary>
        /// <param name="educationInfo"></param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> InsertEducation([FromBody] EducationInfo educationInfo)
        {
            var app = await _educationService.InsertEducation(educationInfo);
            return Ok(app.AppResult);
        }

        /// <summary>
        /// Update Education
        /// </summary>
        /// <param name="educationInfo"></param>
        /// <returns></returns>
        [HttpPut]
        public async Task<IActionResult> UpdatetEducation([FromBody] EducationInfo educationInfo)
        {
            var app = await _educationService.UpdatetEducation(educationInfo);
            return Ok(app.AppResult);
        }

        /// <summary>
        /// Delete Education
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [HttpDelete]
        public async Task<IActionResult> DeleteEducation(int id)
        {
            var app = await _educationService.DeleteEducation(id);
            return Ok(app.AppResult);
        }
    }
}
