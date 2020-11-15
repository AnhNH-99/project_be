using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebAPI.Models;

namespace WebAPI.ViewModels
{
    public class TechnologyViewModel:BaseViewModel
    {
       /* public int Id { get; set; }
        public string Name { get; set; }
        public int CategoryId { get; set; }*/
        private Technology _technology;
        public TechnologyViewModel()
        {
            if (_technology == null)
                this._technology = new Technology();
        }
        public Technology Technology
        {
            get { return _technology; }
            set { _technology = value; }
        }
    }
}
