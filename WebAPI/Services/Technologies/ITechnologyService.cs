using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebAPI.Models;
using WebAPI.ViewModels;

namespace WebAPI.Services.Technologies
{
    public interface ITechnologyService
    {
        /// <summary>
        /// Get Technology
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        Task<IEnumerable<Technology>> GetTechnology(int? id);

        /// <summary>
        /// Create Technology
        /// </summary>
        /// <param name="technology"></param>
        /// <returns></returns>
        public Task<TechnologyViewModel> InsertTechnology(Technology technology);

        /// <summary>
        /// Update Technology
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        public Task<TechnologyViewModel> UpdateTechnology(Technology technology);

        /// <summary>
        /// Delete Technology
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public Task<TechnologyViewModel> DeleteTechnology(int id);
    }
}
