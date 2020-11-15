using Microsoft.AspNetCore.Mvc;
using System.Linq;
using System.Threading.Tasks;
using WebAPI.Models;
using WebAPI.Services.WorkHistories;
using WebAPI.ViewModels;

namespace WebAPI.Controllers
{
    [Route("api/[controller]/[action]/{id?}")]
    [ApiController]
    public class WorkHistoryController : BaseApiController<WorkHistoryViewModel>
    {
        private IWorkHistoryService _workHistoryService;
        public WorkHistoryController(IWorkHistoryService workHistoryService)
        {
            this._workHistoryService = workHistoryService;
        }

        /// <summary>
        /// Get all WorkHistory
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public async Task<IActionResult> GetAllWorkHistory()
        {
            var workHistories = await _workHistoryService.GetAllWorkHistory();
            apiResult.AppResult.DataResult = workHistories.AsEnumerable<WorkHistoryInfo>();
            return Ok(apiResult.AppResult);
        }

        /// <summary>
        /// Get WorkHistory by Id
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [HttpGet]
        public async Task<IActionResult> GetworkHistoryById(int id)
        {
            var app = await _workHistoryService.GetWorkHistoryById(id);
            apiResult.AppResult.DataResult = app;
            return Ok(apiResult.AppResult);
        }
        /// <summary>
        /// Get WorkHistory by PersonId
        /// </summary>
        /// <param name="personId"></param>
        /// <returns></returns>
        [HttpGet]
        public async Task<IActionResult> GetWorkHistoryByPersonId(int id)
        {
            var app = await _workHistoryService.GetWorkHistoryByPersonId(id);
            apiResult.AppResult.DataResult = app;
            return Ok(apiResult.AppResult);
        }

        /// <summary>
        /// Create WorkHistory
        /// </summary>
        /// <param name="workHistory"></param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> InsertworkHistory([FromBody] WorkHistoryInfo workHistory)
        {
            var app = await _workHistoryService.InsertWorkHistory(workHistory);
            return Ok(app.AppResult);
        }

        /// <summary>
        /// Update WorkHistory
        /// </summary>
        /// <param name="workHistory"></param>
        /// <returns></returns>
        [HttpPut]
        public async Task<IActionResult> UpdateWorkHistory([FromBody] WorkHistoryInfo workHistory)
        {
            var app = await _workHistoryService.UpdateWorkHistory(workHistory);
            return Ok(app.AppResult);
        }

        /// <summary>
        /// Delete WorkHistory
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [HttpDelete]
        public async Task<IActionResult> DeleteworkHistory(int id)
        {
            var app = await _workHistoryService.DeleteWorkHistory(id);
            return Ok(app.AppResult);
        }
    }
}
