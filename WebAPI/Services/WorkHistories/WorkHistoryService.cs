using System.Collections.Generic;
using System.Threading.Tasks;
using WebAPI.Models;
using WebAPI.Repositories.WorkHistories;
using WebAPI.ViewModels;

namespace WebAPI.Services.WorkHistories
{
    public class WorkHistoryService : BaseService<WorkHistoryViewModel>, IWorkHistoryService
    {
        IWorkHistoryRepository _workHistoryRepository;
        public WorkHistoryService(IWorkHistoryRepository workHistoryRepository)
        {
            this._workHistoryRepository = workHistoryRepository;
        }

        /// <summary>
        /// Get all WorkHistory
        /// </summary>
        /// <returns></returns>
        public async Task<IEnumerable<WorkHistoryInfo>> GetAllWorkHistory()
        {
            return await _workHistoryRepository.GetAllAsync();
        }

        /// <summary>
        /// Get WorkHistory by Id
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<WorkHistoryInfo> GetWorkHistoryById(int id)
        {
            return await _workHistoryRepository.FindAsync(id);
        }

        /// <summary>
        /// Get WorkHistory by PersonId
        /// </summary>
        /// <param name="personId"></param>
        /// <returns></returns>
        public async Task<IEnumerable<WorkHistoryInfo>> GetWorkHistoryByPersonId(int id)
        {
            return await _workHistoryRepository.FindAsyncByPersonId(id);
        }

        /// <summary>
        /// Create WorkHistory
        /// </summary>
        /// <param name="workHistory"></param>
        /// <returns></returns>
        public async Task<WorkHistoryViewModel> InsertWorkHistory(WorkHistoryInfo workHistory)
        {
            model.WorkHistory = workHistory;
            await _workHistoryRepository.InsertAsync(model.WorkHistory);
            return model;
        }

        /// <summary>
        /// Update WorkHistory
        /// </summary>
        /// <param name="workHistory"></param>
        /// <returns></returns>
        public async Task<WorkHistoryViewModel> UpdateWorkHistory(WorkHistoryInfo workHistory)
        {
            model.WorkHistory = workHistory;
            await _workHistoryRepository.UpdateAsync(model.WorkHistory);
            return model;
        }

        /// <summary>
        /// Delete WorkHistory
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<WorkHistoryViewModel> DeleteWorkHistory(int id)
        {
            await _workHistoryRepository.DeleteAsync(id);
            return model;
        }
    }
}
