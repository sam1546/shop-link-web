using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Repository.CustomViewModels; 

namespace Repository.Interface
{
    public interface ILogin
    {   
        Task<LoginVM> Login(string userid, string password);  
    }
}
