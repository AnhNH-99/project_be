using System.Collections.Generic;
using System.Threading.Tasks;
using WebAPI.Models;
using WebAPI.ViewModels;

namespace WebAPI.Services.WorkHistories
{
    public interface IWorkHistoryService 
    {
        /// <summary>
        /// Get all WorkHistory
        /// </summary>
        /// <returns></returns>
        Task<IEnumerable<WorkHistoryInfo>> GetAllWorkHistory();

        /// <summary>
        /// Get WorkHistory by Id
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        Task<WorkHistoryInfo> GetWorkHistoryById(int id);

        /// <summary>
        /// Get WorkHistory by PersonId
        /// </summary>
        /// <param name="personId"></param>
        /// <returns></returns>
        Task<IEnumerable<WorkHistoryInfo>> GetWorkHistoryByPersonId(int id);

        /// <summary>
        /// Create WorkHistory
        /// </summary>
        /// <param name="workHistory"></param>
        /// <returns></returns>
        public Task<WorkHistoryViewModel> InsertWorkHistory(WorkHistoryInfo workHistory);

        /// <summary>
        /// Update WorkHistory
        /// </summary>
        /// <param name="workHistory"></param>
        /// <returns></returns>
        public Task<WorkHistoryViewModel> UpdateWorkHistory(WorkHistoryInfo workHistory);

        /// <summary>
        /// Delete WorkHistory
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public Task<WorkHistoryViewModel> DeleteWorkHistory(int id);
    }
}
