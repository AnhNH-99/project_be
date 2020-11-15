using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebAPI.Models;
using WebAPI.Repositories.Interfaces;

namespace WebAPI.Repositories.Technologies
{
    public interface ITechnologyRepository:IRepositoryBase<Technology>
    {
        /// <summary>
        /// Get Technology
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        Task<IEnumerable<Technology>> GetTechnologyAsync(int? id);

        /// <summary>
        /// Check Technology
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        Task<bool> CheckTechnologyAsync(Technology technology);

        /// <summary>
        /// Check Technology on delete
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        Task<Technology> CheckTechnologyOnDeleteAsync(Technology technology);

        /// <summary>
        /// Update Technology to Insert
        /// </summary>
        /// <param name="categoryId"></param>
        /// <returns></returns>
        Task<int> UpdateTechnologyToInsert(Technology technology);
    }
}
