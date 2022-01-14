    using Microsoft.AspNetCore.Builder;
    using Microsoft.AspNetCore.Hosting;
    using Microsoft.EntityFrameworkCore;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.DependencyInjection;
    using Repository.Models;
    using Microsoft.AspNetCore.Rewrite;
    using Microsoft.OpenApi.Models;

    namespace CoreAPI
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
                //Add CORS Support
                //services.AddCors();

                //We can even create the policy and later we can use these policies' runtime:
                // Add service and create Policy with options

                services.AddCors(options =>
                {
                    options.AddPolicy("AllowAll",
                        builder =>
                        {
                            builder
                            .AllowAnyOrigin()
                            .AllowAnyMethod()
                            .AllowAnyHeader()
                            .AllowCredentials();
                        });
                });


                services.AddDbContext<Context>(options =>
                            options.UseSqlServer(Configuration.GetConnectionString("DefaultConnection")));

           
                //Register DI Extension Method
                services.RegisterServices();
                services.RegisterSingleTone(Configuration);
          
            
                services.AddMvc();
                services.AddSwaggerGen(c => {
                    c.SwaggerDoc("v1", new OpenApiInfo() { Title = "Shop-Link API", Description = "Shop-Link API" });
                });

            }

            // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
            public void Configure(IApplicationBuilder app, IHostingEnvironment env)
            {
                if (env.IsDevelopment())
                {
                    app.UseDeveloperExceptionPage();
                }

           
                app.UseCors(x => x.AllowAnyOrigin().AllowAnyMethod().AllowAnyHeader().AllowCredentials());
            
                var option = new RewriteOptions();
                option.AddRedirect("^$", "swagger");
                app.UseRewriter(option);
                app.UseMvc();
                app.UseSwagger();
                app.UseSwaggerUI(c => {
                    c.SwaggerEndpoint(Configuration["SwaggerConfiguration"], "Shop-Link API");
                });
                app.UseMvc();
            }
        }
    }
