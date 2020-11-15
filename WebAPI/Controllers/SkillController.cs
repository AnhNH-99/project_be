using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using WebAPI.Models;
using WebAPI.RequestModel;
using WebAPI.Services.Categories;
using WebAPI.Services.Skills;
using WebAPI.ViewModels;

namespace WebAPI.Controllers
{
    [Route("api/[controller]/[action]/{id?}")]
    [ApiController]
    public class SkillController : BaseApiController<SkillViewModel>
    {
        private ISkillService _skillService;
        public SkillController(ISkillService skillService)
        {
            this._skillService = skillService;
        }

        /// <summary>
        /// Get Skill
        /// </summary>
        /// <param name="personId"></param>
        /// <returns></returns>
        [HttpGet]
        public async Task<IActionResult> GetSkill(int id)
        {
            var skills = await _skillService.GetSkillByPerson(id);
            apiResult.AppResult.DataResult = skills;
            return Ok(apiResult.AppResult);
        }


        /// <summary>
        /// Get Skill By Category
        /// </summary>
        /// <param name="personCategory"></param>
        /// <returns></returns>
        [HttpGet]
        public async Task<IActionResult> GetSkilByCategory(SkillRequestModel skillRequestModel)
        {
            var skills = await _skillService.GetSkilByCategoryy(skillRequestModel);
            apiResult.AppResult.DataResult = skills;
            return Ok(apiResult.AppResult);
        }

        /// <summary>
        /// Insert Skill
        /// </summary>
        /// <param name="skillRequestModel"></param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> InsertSkill(SkillRequestModel skillRequestModel)
        {
            var app = await _skillService.InserSkill(skillRequestModel);
            return Ok(app.AppResult);
        }

        /// <summary>
        /// Update Skill
        /// </summary>
        /// <param name="skillRequestModel"></param>
        /// <returns></returns>
        [HttpPut]
        public async Task<IActionResult> UpdateSkill(SkillRequestModel skillRequestModel)
        {
            var app = await _skillService.UpdateSkill(skillRequestModel);
            return Ok(app.AppResult);
        }

        /// <summary>
        /// Delete Skill By Technology
        /// </summary>
        /// <param name="personTechnology"></param>
        /// <returns></returns>
        [HttpDelete]
        public async Task<IActionResult> DeleteSkill(SkillRequestModel skillRequestModel)
        {
            var app = await _skillService.DeleteSkill(skillRequestModel);
            return Ok(app.AppResult);
        }

    }
}
