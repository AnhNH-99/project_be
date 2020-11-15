using WebAPI.Models;

namespace WebAPI.ViewModels
{
    public class AccountViewModel : BaseViewModel
    {
        private AccountInfo _accountInfo;
        public AccountViewModel() {
            if (_accountInfo == null)
                _accountInfo = new AccountInfo();
        }
        public AccountInfo AccountInfo
        {
            get { return _accountInfo; }
            set { _accountInfo = value; }
        }
    }
}
