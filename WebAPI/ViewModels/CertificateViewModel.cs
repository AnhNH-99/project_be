using WebAPI.Models;

namespace WebAPI.ViewModels
{
    public class CertificateViewModel : BaseViewModel
    {
        private CertificateInfo _certificateInfo;
        public CertificateViewModel()
        {
            if (_certificateInfo == null)
                _certificateInfo = new CertificateInfo();
        }
        public CertificateInfo CertificateInfo
        {
            get { return _certificateInfo; }
            set { _certificateInfo = value; }
        }
    }
}
