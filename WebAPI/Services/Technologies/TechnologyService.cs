using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using WebAPI.Models;
using WebAPI.Repositories.Technologies;
using WebAPI.ViewModels;

namespace WebAPI.Services.Technologies
{
    public class TechnologyService : BaseService<TechnologyViewModel>, ITechnologyService
    {
        ITechnologyRepository _technologyRepository;
        public TechnologyService(ITechnologyRepository technologyRepository)
        {
            this._technologyRepository = technologyRepository;
        }

        /// <summary>
        /// Get Technology
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<IEnumerable<Technology>> GetTechnology(int? id)
        {
            return await _technologyRepository.GetTechnologyAsync(id);
        }

        /// <summary>
        /// Create Technology
        /// </summary>
        /// <param name="technology"></param>
        /// <returns></returns>
        public async Task<TechnologyViewModel> InsertTechnology(Technology technology)
        {
            //model.Technology = technology;
            technology.CreatedBy = WebAPI.Helpers.HttpContext.CurrentUser;
            technology.UpdatedBy = WebAPI.Helpers.HttpContext.CurrentUser;
            if(await _technologyRepository.CheckTechnologyAsync(technology))
            {
                model.AppResult = new AppResult { Result = false, StatusCd = "-1", Message = "Fail" };
            }
            else
            {
                Technology modelTechnology = await _technologyRepository.CheckTechnologyOnDeleteAsync(technology);
                if(modelTechnology == null)
                {
                   var result = await _technologyRepository.InsertAsync(technology);
                    if (result > 0)
                        model.AppResult = new AppResult { Result = true, StatusCd = "0", Message = "Success" };
                    else
                        model.AppResult = new AppResult { Result = false, StatusCd = "-1", Message = "Fail" };
                }
                else
                {
                    technology.Id = modelTechnology.Id;
                    technology.UpdatedAt = DateTime.Now;
                    int result = await _technologyRepository.UpdateTechnologyToInsert(technology);
                    if (result > 0)
                        model.AppResult = new AppResult { Result = true, StatusCd = "0", Message = "Success" };
                    else
                        model.AppResult = new AppResult { Result = false, StatusCd = "-1", Message = "Fail" };
                }
            }
            return model;
        }

        /// <summary>
        /// Update Technology
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        public async Task<TechnologyViewModel> UpdateTechnology(Technology technology)
        {
            technology.UpdatedAt = DateTime.Now;
            technology.UpdatedBy = WebAPI.Helpers.HttpContext.CurrentUser;
            if (await _technologyRepository.CheckTechnologyAsync(technology))
            {
                model.AppResult = new AppResult { Result = false, StatusCd = "-1", Message = "Fail" };
            }
            else
            {
                Technology modelTechnology = await _technologyRepository.CheckTechnologyOnDeleteAsync(technology);
                if (modelTechnology == null)
                {
                    var result = await _technologyRepository.UpdateAsync(technology);
                    if (result > 0)
                        model.AppResult = new AppResult { Result = true, StatusCd = "0", Message = "Success" };
                    else
                        model.AppResult = new AppResult { Result = false, StatusCd = "-1", Message = "Fail" };
                }
                else
                {
                    technology.UpdatedAt = DateTime.Now;
                    int result = await _technologyRepository.UpdateTechnologyToInsert(technology);
                    if (result > 0)
                        model.AppResult = new AppResult { Result = true, StatusCd = "0", Message = "Success" };
                    else
                        model.AppResult = new AppResult { Result = false, StatusCd = "-1", Message = "Fail" };
                }
            }
            await _technologyRepository.UpdateAsync(technology);
            return model;
        }

        /// <summary>
        /// Delete Technology
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<TechnologyViewModel> DeleteTechnology(int id)
        {
            await _technologyRepository.DeleteAsync(id);
            return model;
        }
    }
}
