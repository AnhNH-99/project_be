using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebAPI.Models;
using WebAPI.Repositories.Interfaces;
using WebAPI.ViewModels;

namespace WebAPI.Repositories.Categories
{
    public interface ICategoryRepository : IRepositoryBase<Category>
    {
        /// <summary>
        /// Check Category
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        Task<bool> CheckCategoryAsync(string name);

        /// <summary>
        /// Check Category on delete
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        Task<Category> CheckCategoryOnDeleteAsync(string name);

        /// <summary>
        /// Update Caegory to Insert
        /// </summary>
        /// <param name="categoryId"></param>
        /// <returns></returns>
        Task<int> UpdateCategoryToInsert(Category category);
    }
}
