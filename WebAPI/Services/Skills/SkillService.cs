using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebAPI.Models;
using WebAPI.Repositories.Skills;
using WebAPI.RequestModel;
using WebAPI.ViewModels;

namespace WebAPI.Services.Skills
{
    public class SkillService : BaseService<SkillViewModel>, ISkillService
    {
        ISkillRepository _skillRepository;
        public SkillService(ISkillRepository skillRepository)
        {
            this._skillRepository = skillRepository;
        }

        /// <summary>
        /// Get Skill By PersonId
        /// </summary>
        /// <param name="personId"></param>
        /// <returns></returns>
        public async Task<IEnumerable<Category>> GetSkillByPerson(int personId)
        {
            return await _skillRepository.GetSkillByPersonAsync(personId);
        }

        /// <summary>
        /// Inser Skill
        /// </summary>
        /// <param name="skillRequestModel"></param>
        /// <returns></returns>
        public async Task<SkillViewModel> InserSkill(SkillRequestModel skillRequestModel)
        {
            model.SkillRequestModel = skillRequestModel;
            List<PersonTechnology> personTechnologies = new List<PersonTechnology>();
            PersonCategory personCategory = new PersonCategory
            {
                PersonId = skillRequestModel.PersonId,
                CategoryId = skillRequestModel.CategoryId,
                OrderIndex = skillRequestModel.OrderIndex,
                CreatedBy = WebAPI.Helpers.HttpContext.CurrentUser,
                UpdatedBy = WebAPI.Helpers.HttpContext.CurrentUser
            };
            foreach (var item in skillRequestModel.TechnologyId)
            {
                PersonTechnology personTechnology = new PersonTechnology
                {
                    PersonId = skillRequestModel.PersonId,
                    TechnologyId = item,
                    CreatedBy = WebAPI.Helpers.HttpContext.CurrentUser,
                    UpdatedBy = WebAPI.Helpers.HttpContext.CurrentUser
                };
                personTechnologies.Add(personTechnology);
            }
            if (!await _skillRepository.CheckPersonCategoryAsync(personCategory))
            {
                if (await _skillRepository.CheckPersonCategoryOnDeleteAsync(personCategory))
                {
                    personCategory.UpdatedAt = DateTime.Now;
                    await _skillRepository.UpdatePersonCategoryToInsertAsync(personCategory);
                }
                else
                {
                    await _skillRepository.InsertPersonCategoryAsync(personCategory);
                }

            }
            List<PersonTechnology> personTechnologiesUpdate = new List<PersonTechnology>();
            List<PersonTechnology> personTechnologiesInsert = new List<PersonTechnology>();
            foreach (var item in personTechnologies)
            {
                if (await _skillRepository.CheckPersonTechnologyOnDeleteAsync(item))
                {
                    item.UpdatedAt = DateTime.Now;
                    personTechnologiesUpdate.Add(item);
                }
                else
                {
                    if(!await _skillRepository.CheckPersonTechnologyAsync(item))
                        personTechnologiesInsert.Add(item);
                }
            }
            if (personTechnologiesUpdate.Count > 0)
            {
                await _skillRepository.UpdatePersonTechnologyToInsertAsync(personTechnologiesUpdate);
            }
            await _skillRepository.InsertListPersonTechnologyAsync(personTechnologiesInsert);
            return model;
        }

        /// <summary>
        /// Update Skill
        /// </summary>
        /// <param name="skillRequestModel"></param>
        /// <returns></returns>
        public async Task<SkillViewModel> UpdateSkill(SkillRequestModel skillRequestModel)
        {
            model.SkillRequestModel = skillRequestModel;
            List<PersonTechnology> personTechnologies = new List<PersonTechnology>();
            PersonCategory personCategory = new PersonCategory
            {
                PersonId = skillRequestModel.PersonId,
                CategoryId = skillRequestModel.CategoryId,
                OrderIndex = skillRequestModel.OrderIndex,
                CreatedBy = WebAPI.Helpers.HttpContext.CurrentUser,
                UpdatedBy = WebAPI.Helpers.HttpContext.CurrentUser
            };
            foreach (var item in skillRequestModel.TechnologyId)
            {
                PersonTechnology personTechnology = new PersonTechnology
                {
                    PersonId = skillRequestModel.PersonId,
                    TechnologyId = item,
                    CreatedBy = WebAPI.Helpers.HttpContext.CurrentUser,
                    UpdatedBy = WebAPI.Helpers.HttpContext.CurrentUser
                };
                personTechnologies.Add(personTechnology);
            }
            await _skillRepository.DeletePersonTechnologyToUpdateAsync(personCategory);
            await _skillRepository.InsertListPersonTechnologyAsync(personTechnologies);
            return model;
        }

        /// <summary>
        /// Delete Skill
        /// </summary>
        /// <param name="skillRequestModel"></param>
        /// <returns></returns>
        public async Task<SkillViewModel> DeleteSkill(SkillRequestModel skillRequestModel)
        {
            model.SkillRequestModel = skillRequestModel;
            PersonCategory personCategory = new PersonCategory
            {
                PersonId = skillRequestModel.PersonId,
                CategoryId = skillRequestModel.CategoryId,
                UpdatedAt = DateTime.Now,
                UpdatedBy = WebAPI.Helpers.HttpContext.CurrentUser
            };
            await _skillRepository.DeletePersonCategoryAsync(personCategory);
            await _skillRepository.DeletePersonTechnologyAsync(personCategory);
            return model;
        }


        /// <summary>
        /// Get Skill By Category
        /// </summary>
        /// <param name="personCategory"></param>
        /// <returns></returns>
        public async Task<IEnumerable<Technology>> GetSkilByCategoryy(SkillRequestModel skillRequestModel)
        {
            PersonCategory personCategory = new PersonCategory
            {
                PersonId = skillRequestModel.PersonId,
                CategoryId = skillRequestModel.CategoryId
            };
            return await _skillRepository.GetSkilByCategoryAsync(personCategory);
        }
    }
}
