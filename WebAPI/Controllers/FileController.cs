using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using WebAPI.HandlingFiles;
using WebAPI.Models;
using WebAPI.Services.Persons;
using WebAPI.Services.WorkHistories;
using static WebAPI.Constants.Constant;


namespace WebAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FileController : ControllerBase
    {
        private IPersonService _personService;
        private IWorkHistoryService _workHistoryService;
        public FileController(IPersonService personSevice, IWorkHistoryService workHistoryService)
        {
            this._personService = personSevice;
            this._workHistoryService = workHistoryService;
        }
        [HttpPost]
        public async Task<IActionResult> ImportFile(IFormFile file)
        {
            var list = await ImportFileHandling.GetInformationCV(file);
            if (list.Count > 0)
            {
                Person person = ImportFileHandling.GetPerson(list);
                if (person != null)
                {
                    await _personService.InsertPerson(person);
                }
            }
            /*int indexSkill = list.IndexOf("SKILLS");
            int indexWorkHistory = list.IndexOf("WORKING HISTORY");
            int indexEducation = list.IndexOf("EDUCATION");
            int indexCertificate = list.IndexOf("CERTIFICATION");
            int indexProject = list.IndexOf("PROJECTS");
            int indexEnd = list.Count();
            if (indexSkill > 0 && indexWorkHistory > indexSkill)
            {

            }
            if(indexWorkHistory > 0 && indexEducation > indexWorkHistory)
            {
                List<WorkHistoryInfo> workHistoryInfos = ImportFileHandling.GetListWorkHistory(list);
                if (workHistoryInfos.Any())
                {
                    foreach(var item in workHistoryInfos)
                    {
                       await _workHistoryService.InsertWorkHistory(item);
                    }
                }
            }
            if(indexEducation > 0 && indexCertificate > indexEducation)
            {

            }
            if(indexCertificate > 0 && indexProject > 0)
            {

            }
            if(indexProject > 0 && indexEnd> indexProject)
            {

            }*/
            return Ok(list);
        }
    }
}
