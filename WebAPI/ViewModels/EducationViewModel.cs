using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebAPI.Models;

namespace WebAPI.ViewModels
{
    public class EducationViewModel : BaseViewModel
    {
        private EducationInfo _educationInfo;
        public EducationViewModel()
        {
            if (_educationInfo == null)
                _educationInfo = new EducationInfo();
        }
        public EducationInfo EducationInfo
        {
            get { return _educationInfo; }
            set { _educationInfo = value; }
        }
    }
}
