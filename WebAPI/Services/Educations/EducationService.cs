using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebAPI.Models;
using WebAPI.Repositories.Educations;
using WebAPI.ViewModels;

namespace WebAPI.Services.Educations
{
    public class EducationService : BaseService<EducationViewModel>, IEducationService
    {
        private IEducationRepository _educationRepository;
        public EducationService(IEducationRepository educationRepository)
        {
            this._educationRepository = educationRepository;
        }

        /// <summary>
        /// Get all Education
        /// </summary>
        /// <returns></returns>
        public async Task<IEnumerable<EducationInfo>> GetAllEducation()
        {
            return await _educationRepository.GetAllAsync();
        }

        /// <summary>
        /// Get Education by Id
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<EducationInfo> GetEducationById(int id)
        {
            return await _educationRepository.FindAsync(id);
        }

        /// <summary>
        /// Create Education
        /// </summary>
        /// <param name="Education"></param>
        /// <returns></returns>
        public async Task<EducationViewModel> InsertEducation(EducationInfo education)
        {
            model.EducationInfo = education;
            await _educationRepository.InsertAsync(model.EducationInfo);
            return model;
        }

        /// <summary>
        /// Update Education
        /// </summary>
        /// <param name="Education"></param>
        /// <returns></returns>
        public async Task<EducationViewModel> UpdatetEducation(EducationInfo Education)
        {
            model.EducationInfo = Education;
            await _educationRepository.UpdateAsync(model.EducationInfo);
            return model;
        }

        /// <summary>
        /// Delete Education
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<EducationViewModel> DeleteEducation(int id)
        {
            await _educationRepository.DeleteAsync(id);
            return model;
        }
    }
}
