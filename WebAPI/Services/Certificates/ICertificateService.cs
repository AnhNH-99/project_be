using System.Collections.Generic;
using System.Threading.Tasks;
using WebAPI.Models;
using WebAPI.ViewModels;

namespace WebAPI.Services.Certificates
{
    public interface ICertificateService
    {
        /// <summary>
        /// Get all Certificate
        /// </summary>
        /// <returns></returns>
        Task<IEnumerable<CertificateInfo>> GetAllCertificate();

        /// <summary>
        /// Get Certificate by Id
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        Task<CertificateInfo> GetCertificateById(int id);

        /// <summary>
        /// Create Certificate
        /// </summary>
        /// <param name="Certificate"></param>
        /// <returns></returns>
        public Task<CertificateViewModel> InsertCertificate(CertificateInfo Certificate);

        /// <summary>
        /// Update Certificate
        /// </summary>
        /// <param name="Certificate"></param>
        /// <returns></returns>
        public Task<CertificateViewModel> UpdatetCertificate(CertificateInfo Certificate);

        /// <summary>
        /// Delete Certificate
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public Task<CertificateViewModel> DeleteCertificate(int id);
    }
}
