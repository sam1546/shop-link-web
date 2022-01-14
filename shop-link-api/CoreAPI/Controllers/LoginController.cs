using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Repository.Interface;
using Repository.CustomViewModels;
using Microsoft.AspNetCore.Cors; 
using Repository.Models;
using Repository.Repositories;

namespace CoreAPI.Controllers
{ 
    [Produces("application/json")]
    [Route("api/Login")]
    public class LoginController : Controller
    {
        public ILogin LoginRepo { get; set; }
        public LoginController(ILogin _loginrepo)
        {
            LoginRepo = _loginrepo;
        }
        
        [HttpGet]
        [Route("User")]
        public async Task<IActionResult> UserLogin(string UserId,string UserPassword)
        { 
            var result = await  LoginRepo.Login(UserId, UserPassword);
            return Ok(result);
        }
    }
}