using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using WebAPI.HandlingFiles;
using WebAPI.Models;
using WebAPI.Repositories.Persons;
using WebAPI.ViewModels;

namespace WebAPI.Services.ImprotFiles
{
    public class ImprotFile : IImportFile
    {
        IPersonRepository _personRepository;
        public ImprotFile(IPersonRepository personRepository)
        {
            this._personRepository = personRepository;
        }
        /*/// <summary>
        /// Save Person
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public async Task<PersonViewModel> SavePerson(List<string> list)
        {
            Person person = ImportFileHandling.GetPerson(list);
            if(person != null)
            {
                person.CreatedBy = WebAPI.Helpers.HttpContext.CurrentUser;
                person.UpdatedBy = WebAPI.Helpers.HttpContext.CurrentUser;
                await _personRepository.InsertAsync(person);
            }
        }*/
    }
}
