using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using WebAPI.Models;
using WebAPI.Services.Certificates;
using WebAPI.ViewModels;

namespace WebAPI.Controllers
{
    [Route("api/[controller]/[action]/{id?}")]
    [ApiController]
    public class CertificateController : BaseApiController<CertificateViewModel>
    {
        private ICertificateService _certificateService;
        public CertificateController(ICertificateService certificateService)
        {
            this._certificateService = certificateService;
        }

        /// <summary>
        /// Get all Certificate
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public async Task<IActionResult> GetAllCertificate()
        {
            var employees = await _certificateService.GetAllCertificate();
            apiResult.AppResult.DataResult = employees.AsEnumerable<CertificateInfo>();
            return Ok(apiResult.AppResult);
        }

        /// <summary>
        /// Get Certificate by Id
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [HttpGet]
        public async Task<IActionResult> GetCertificateById(int id)
        {
            var app = await _certificateService.GetCertificateById(id);
            apiResult.AppResult.DataResult = app;
            return Ok(apiResult.AppResult);
        }

        /// <summary>
        /// Create Certificate
        /// </summary>
        /// <param name="certificateInfo"></param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> InsertCertificate([FromBody] CertificateInfo certificateInfo)
        {
            var app = await _certificateService.InsertCertificate(certificateInfo);
            return Ok(app.AppResult);
        }

        /// <summary>
        /// Update Certificate
        /// </summary>
        /// <param name="certificateInfo"></param>
        /// <returns></returns>
        [HttpPut]
        public async Task<IActionResult> UpdatetCertificate([FromBody] CertificateInfo certificateInfo)
        {
            var app = await _certificateService.UpdatetCertificate(certificateInfo);
            return Ok(app.AppResult);
        }

        /// <summary>
        /// Delete Certificate
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [HttpDelete]
        public async Task<IActionResult> DeleteCertificate(int id)
        {
            var app = await _certificateService.DeleteCertificate(id);
            return Ok(app.AppResult);
        }
    }
}
