using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebAPI.Models;
using WebAPI.Repositories.Interfaces;
using WebAPI.ViewModels;

namespace WebAPI.Repositories.Skills
{
    public interface ISkillRepository/*:IRepositoryBase<SkillViewModel>*/
    {
        /// <summary>
        /// Get Skill By PersonId
        /// </summary>
        /// <param name="personId"></param>
        /// <returns></returns>
        Task<IEnumerable<Category>> GetSkillByPersonAsync(int personId);

        /// <summary>
        /// Get Skill By Category
        /// </summary>
        /// <param name="personCategory"></param>
        /// <returns></returns>
        Task<IEnumerable<Technology>> GetSkilByCategoryAsync(PersonCategory personCategory);

        /// <summary>
        /// Check PersonCategory
        /// </summary>
        /// <param name="personCategory"></param>
        /// <returns></returns>
        Task<bool> CheckPersonCategoryAsync(PersonCategory personCategory);

        /// <summary>
        /// Check PersonCategory On Delete
        /// </summary>
        /// <param name="personCategory"></param>
        /// <returns></returns>
        Task<bool> CheckPersonCategoryOnDeleteAsync(PersonCategory personCategory);

        /// <summary>
        /// Update PersonCategory to Insert
        /// </summary>
        /// <param name="personCategory"></param>
        /// <returns></returns>
        Task<int> UpdatePersonCategoryToInsertAsync(PersonCategory personCategory);

        /// <summary>
        /// Insert PersonCategory
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        Task<int> InsertPersonCategoryAsync(PersonCategory entity);

        /// <summary>
        /// Check PersonTechnology
        /// </summary>
        /// <param name="personTechnology"></param>
        /// <returns></returns>
        Task<bool> CheckPersonTechnologyAsync(PersonTechnology personTechnology);

        /// <summary>
        /// Check PersonTechnology
        /// </summary>
        /// <param name="personTechnology"></param>
        /// <returns></returns>
        Task<bool> CheckPersonTechnologyOnDeleteAsync(PersonTechnology personTechnology);

        /// <summary>
        /// Update PersonTechnology to Insert
        /// </summary>
        /// <param name="personTechnologies"></param>
        /// <returns></returns>
        Task<int> UpdatePersonTechnologyToInsertAsync(List<PersonTechnology> personTechnologies);

        /// <summary>
        /// Create List PersonTechnology
        /// </summary>
        /// <param name="personTechnologies"></param>
        /// <returns></returns>
        Task<int> InsertListPersonTechnologyAsync(List<PersonTechnology> personTechnologies);

        /// <summary>
        /// Delete PersonTechnology to Update
        /// </summary>
        /// <param name="personCategory"></param>
        /// <returns></returns>
        Task<int> DeletePersonTechnologyToUpdateAsync(PersonCategory personCategory);

        /// <summary>
        /// Delete PersonCategory
        /// </summary>
        /// <param name="category"></param>
        /// <returns></returns>
        Task<int> DeletePersonCategoryAsync(PersonCategory category);

        /// <summary>
        /// Delete PersonTechnology
        /// </summary>
        /// <param name="personCategory"></param>
        /// <returns></returns>
        Task<int> DeletePersonTechnologyAsync(PersonCategory personCategory);

        
    }
}
