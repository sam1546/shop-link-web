    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Threading.Tasks;
    using Repository.CustomViewModels;
    using Repository.Interface;
    using Repository.Models;
    using Microsoft.EntityFrameworkCore;
    using System.Linq;
    using System.Net; 

    namespace Repository.Repositories
    {
        public class LoginRepo : ILogin
        {
            private readonly Context _context;
            public LoginRepo(Context context)
            {
                _context = context;
            }

            public async Task<LoginVM> Login(string userid, string password)
            { 
                var result = await _context.tblUserDetails.Where(um => um.UserName == userid && um.Password == password).SingleOrDefaultAsync();
                if (result == null)
                {
                    return null;
                }
                else
                {
                LoginVM userLoginDetailsVM = new LoginVM()
                {
                    UserName = result.UserName,
                    //Password = result.Password, 
                    //Repassword = result.Repassword,
                    Department = result.Department,
                    PlantAddress = result.PlantAddress,
                    Plantcode = result.Plantcode,
                    Role = result.Role
                };
                    return userLoginDetailsVM;
                }
            }
        }
    }