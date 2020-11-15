using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.IdentityModel.Tokens;
using System;
using System.Text;
using WebAPI.Helpers;
using WebAPI.Models;
using WebAPI.Repositories.Accounts;
using WebAPI.Repositories.Categories;
using WebAPI.Repositories.Skills;
using WebAPI.Repositories.Technologies;
using WebAPI.Services.Accounts;
using WebAPI.Services.Categories;
using WebAPI.Services.Skills;
using WebAPI.Services.Technologies;

namespace WebAPI
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddControllers();
            //services.AddControllers(x => x.AllowEmptyInputInBodyModelBinding = true);
            services.AddMvc()
                    .AddJsonOptions(opts => opts.JsonSerializerOptions.PropertyNamingPolicy = null);

            var appSettingsSection = Configuration.GetSection("AppSettings");
            services.Configure<AppSettings>(appSettingsSection);

            // configure jwt authentication
            var appSettings = appSettingsSection.Get<AppSettings>();
            var key = Encoding.ASCII.GetBytes(appSettings.Secret);
            services.AddAuthentication(x =>
            {
                x.DefaultAuthenticateScheme = JwtBearerDefaults.AuthenticationScheme;
                x.DefaultChallengeScheme = JwtBearerDefaults.AuthenticationScheme;
            })
            .AddJwtBearer(x =>
            {
                x.RequireHttpsMetadata = false;
                x.SaveToken = true;
                x.TokenValidationParameters = new TokenValidationParameters
                {
                    ValidateIssuerSigningKey = true,
                    IssuerSigningKey = new SymmetricSecurityKey(key),
                    ValidateIssuer = false,
                    ValidateAudience = false,
                    // set clockskew to zero so tokens expire exactly at token expiration time (instead of 5 minutes later)
                    ClockSkew = TimeSpan.Zero
                };
            });
            services.AddHttpContextAccessor();
            services.AddScoped<IAccountRepository>(x => new AccountRepository(Configuration.GetConnectionString("DefaultConnection")));
            services.AddScoped(typeof(IAccountService), typeof(AccountService));
            services.AddScoped<ICategoryRepository>(x => new CategoryRepository(Configuration.GetConnectionString("DefaultConnection")));
            services.AddScoped(typeof(ICategoryService), typeof(CategoryService));
            services.AddScoped<ITechnologyRepository>(x => new TechnologyRepository(Configuration.GetConnectionString("DefaultConnection")));
            services.AddScoped(typeof(ITechnologyService), typeof(TechnologyService));
            services.AddScoped<ISkillRepository>(x => new SkillRepository(Configuration.GetConnectionString("DefaultConnection")));
            services.AddScoped(typeof(ISkillService), typeof(SkillService));

        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            HttpContext.Configure(app.ApplicationServices.
                GetRequiredService<Microsoft.AspNetCore.Http.IHttpContextAccessor>()
            );
            if (env.IsDevelopment())
            {
                app.UseExceptionHandler("/api/Error/CreateLog");
            }
            else
            {
                app.UseExceptionHandler("/api/Error/CreateLog");
                app.UseHsts();
            }

            //app.UseHttpsRedirection();

            app.UseRouting();

            // global cors policy
            app.UseCors(x => x
                .SetIsOriginAllowed(origin => true)
                .AllowAnyMethod()
                .AllowAnyHeader()
                .AllowCredentials());

            app.UseAuthentication();
            app.UseAuthorization();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllerRoute(
                    name: "DefaultApi",
                    pattern: "{controller}/{action}/{id?}");

                endpoints.MapControllerRoute(
                    name: "ActionApi",
                    pattern: "api/{controller}/{action}/{id?}");
            });
        }
    }
}
