using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebAPI.Models;
using WebAPI.RequestModel;
using WebAPI.ViewModels;

namespace WebAPI.Services.Skills
{
    public interface ISkillService
    {
        /// <summary>
        /// Get Skill By PersonId
        /// </summary>
        /// <param name="personId"></param>
        /// <returns></returns>
        Task<IEnumerable<Category>> GetSkillByPerson(int personId);

        /// <summary>
        /// Inser Skill
        /// </summary>
        /// <param name="skillRequestModel"></param>
        /// <returns></returns>
        Task<SkillViewModel> InserSkill(SkillRequestModel skillRequestModel);


        /// <summary>
        /// Update Skill
        /// </summary>
        /// <param name="skillRequestModel"></param>
        /// <returns></returns>
        Task<SkillViewModel> UpdateSkill(SkillRequestModel skillRequestModel);

        /// <summary>
        /// Delete Skill 
        /// </summary>
        /// <param name="skillRequestModel"></param>
        /// <returns></returns>
        Task<SkillViewModel> DeleteSkill(SkillRequestModel skillRequestModel);

        /// <summary>
        /// Get Skill By Category
        /// </summary>
        /// <param name="skillRequestModel"></param>
        /// <returns></returns>
        Task<IEnumerable<Technology>> GetSkilByCategoryy(SkillRequestModel skillRequestModel);
    }
}
