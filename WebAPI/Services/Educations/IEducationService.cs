using System.Collections.Generic;
using System.Threading.Tasks;
using WebAPI.Models;
using WebAPI.ViewModels;

namespace WebAPI.Services.Educations
{
    public interface IEducationService
    {
        /// <summary>
        /// Get all Education
        /// </summary>
        /// <returns></returns>
        Task<IEnumerable<EducationInfo>> GetAllEducation();

        /// <summary>
        /// Get Education by Id
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        Task<EducationInfo> GetEducationById(int id);

        /// <summary>
        /// Create Education
        /// </summary>
        /// <param name="Education"></param>
        /// <returns></returns>
        public Task<EducationViewModel> InsertEducation(EducationInfo Education);

        /// <summary>
        /// Update Education
        /// </summary>
        /// <param name="Education"></param>
        /// <returns></returns>
        public Task<EducationViewModel> UpdatetEducation(EducationInfo Education);

        /// <summary>
        /// Delete Education
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public Task<EducationViewModel> DeleteEducation(int id);
    }
}
