using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using WebAPI.Models;
using WebAPI.Services.Categories;
using WebAPI.ViewModels;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace WebAPI.Controllers
{
    [Route("api/[controller]/[action]/{id?}")]
    [ApiController]
    public class CategoryController : BaseApiController<CategoryViewModel>
    {
        private ICategoryService _categoryService;
        public CategoryController(ICategoryService categoryService)
        {
            this._categoryService = categoryService;
        }
        /// <summary>
        /// Get all Category
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public async Task<IActionResult> GetAllCategory()
        {
            var app = await _categoryService.GetAllCategory();
            apiResult.AppResult.DataResult = app;
            return Ok(apiResult.AppResult);
        }

        /// <summary>
        /// Create Category
        /// </summary>
        /// <param name="category"></param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> InsertCategory([FromBody] Category category)
        {
            var app = await _categoryService.InsertCategory(category);
            return Ok(app.AppResult);
        }

        /// <summary>
        /// Update Category
        /// </summary>
        /// <param name="category"></param>
        /// <returns></returns>
        [HttpPut]
        public async Task<IActionResult> UpdateCategory([FromBody] Category category)
        {
            var app = await _categoryService.UpdateCategory(category);
            return Ok(app.AppResult);
        }

        /// <summary>
        /// Delete Category
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [HttpDelete]
        public async Task<IActionResult> DeleteCategory(int id)
        {
            var app = await _categoryService.DeleteCategory(id);
            return Ok(app.AppResult);
        }
    }
}
