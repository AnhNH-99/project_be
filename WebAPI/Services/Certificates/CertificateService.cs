using System.Collections.Generic;
using System.Threading.Tasks;
using WebAPI.Models;
using WebAPI.Repositories.Certificates;
using WebAPI.ViewModels;

namespace WebAPI.Services.Certificates
{
    public class CertificateService : BaseService<CertificateViewModel>, ICertificateService
    {
        private ICertificateRepository _certificateRepository;
        public CertificateService(ICertificateRepository certificateRepository)
        {
            this._certificateRepository = certificateRepository;
        }

        /// <summary>
        /// Get all Certificate
        /// </summary>
        /// <returns></returns>
        public async Task<IEnumerable<CertificateInfo>> GetAllCertificate()
        {
            return await _certificateRepository.GetAllAsync();
        }

        /// <summary>
        /// Get Certificate by Id
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<CertificateInfo> GetCertificateById(int id)
        {
            return await _certificateRepository.FindAsync(id);
        }

        /// <summary>
        /// Create Certificate
        /// </summary>
        /// <param name="Certificate"></param>
        /// <returns></returns>
        public async Task<CertificateViewModel> InsertCertificate(CertificateInfo certificate)
        {
            model.CertificateInfo = certificate;
            await _certificateRepository.InsertAsync(model.CertificateInfo);
            return model;
        }

        /// <summary>
        /// Update Certificate
        /// </summary>
        /// <param name="Certificate"></param>
        /// <returns></returns>
        public async Task<CertificateViewModel> UpdatetCertificate(CertificateInfo Certificate)
        {
            model.CertificateInfo = Certificate;
            await _certificateRepository.UpdateAsync(model.CertificateInfo);
            return model;
        }

        /// <summary>
        /// Delete Certificate
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<CertificateViewModel> DeleteCertificate(int id)
        {
            await _certificateRepository.DeleteAsync(id);
            return model;
        }
    }
}
