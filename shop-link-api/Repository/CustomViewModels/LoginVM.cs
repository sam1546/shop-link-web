using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace Repository.CustomViewModels
{
    public class LoginVM
    { 
        [Required(ErrorMessage = "User Name required.", AllowEmptyStrings = false)] 
        public string UserName { get; set; }
        [Required(ErrorMessage = "Password required.", AllowEmptyStrings = false)]
        [DataType(DataType.Password)]
        public string Password { get; set; }
        public string Repassword { get; set; }
        public string Plantcode { get; set; }
        public string Role { get; set; }
        public string PlantAddress { get; set; }
        public string Department { get; set; }
    }
}
