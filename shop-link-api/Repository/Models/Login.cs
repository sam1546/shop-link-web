using System;
using System.ComponentModel.DataAnnotations;

namespace Repository.Models
{
    public partial class Login
    {
        [Key]
        public string UserName { get; set; }
        public string Password { get; set; }
        public string Repassword { get; set; }
        public string Plantcode { get; set; }
        public string Role { get; set; }
        public string PlantAddress { get; set; }
        public string Department { get; set; }
    }
}
