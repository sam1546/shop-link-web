using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using Repository;
using Repository.Interface;
using Repository.Repositories;
using Microsoft.Extensions.Configuration;

namespace CoreAPI
{
    public static class ServiceExtensions
    {
        public static IServiceCollection RegisterServices(
            this IServiceCollection services)
        {
            // Add all other services here.
            services.AddTransient<ILogin, LoginRepo>();
            services.AddTransient<IOperation, OperationRepo>();
            return services;
        }

        public static void RegisterSingleTone(this IServiceCollection services, IConfiguration config)
        {
            services.AddSingleton<IConfiguration>(config);
        }
    }
}
