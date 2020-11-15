using WebAPI.Services.Persons;
using WebAPI.Models;
using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;
using WebAPI.ViewModels;
using System.Linq;
using WebAPI.RequestModel;

namespace WebAPI.Controllers
{
    [Route("api/[controller]/[action]/{id?}")]
    [ApiController]
    public class PersonController: BaseApiController<PersonViewModel>
    {
        private IPersonService _personService;
        public PersonController(IPersonService personSevice)
        {
            this._personService = personSevice;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public async Task<IActionResult> GetPersonAndSkill()
        {
            var person = await _personService.GetAllPersonAndSkill();
            apiResult.AppResult.DataResult = person;
            return Ok(apiResult.AppResult);
        }



/*        [HttpGet]
        public async Task<IActionResult> Pagination(int indexPagination)
        {
            var employees = await _personService.Pagination(indexPagination);
            apiResult.AppResult.DataResult = employees.AsEnumerable<Person>();

            return Ok(apiResult.AppResult);
        }*/


        /// <summary>
        /// Get all Person
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public async Task<IActionResult> GetAllPerson()
        {
            var employees = await _personService.GetAllPerson();
            apiResult.AppResult.DataResult = employees.AsEnumerable<Person>();
            
            return Ok(apiResult.AppResult);
        }

        /// <summary>
        /// Get Person by Id
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [HttpGet]
        public async Task<IActionResult> GetPersonById(int id)
        {
            var app = await _personService.GetPersonById(id);
            apiResult.AppResult.DataResult = app;
            return Ok(apiResult.AppResult);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="testRequestModel"></param>
        /// <returns></returns>
        [HttpGet]
        public async Task<IActionResult> SearchPerson(TestRequestModel testRequestModel)
        {
            
            var app = await _personService.SearchPersonAndSkillAsync(testRequestModel.FullName,testRequestModel.Location, testRequestModel.TechnologyId);
            apiResult.AppResult.DataResult = app;
            return Ok(apiResult.AppResult);
        }



        /// <summary>
        /// Create Person
        /// </summary>
        /// <param name="personInfo"></param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> InsertPerson([FromBody] Person personInfo)
        {
            var app = await _personService.InsertPerson(personInfo);
            return Ok(app.AppResult);
        }

        /// <summary>
        /// Update Person
        /// </summary>
        /// <param name="personInfo"></param>
        /// <returns></returns>
        [HttpPut]
        public async Task<IActionResult> UpdatePerson([FromBody] Person personInfo)
        {
            var app = await _personService.UpdatePerson(personInfo);
            return Ok(app.AppResult);
        }

        /// <summary>
        /// Delete Person
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [HttpDelete]
        public async Task<IActionResult> DeletePerson(int id)
        {
            var app = await _personService.DeletePerson(id);
            return Ok(app.AppResult);
        }
    }
}

