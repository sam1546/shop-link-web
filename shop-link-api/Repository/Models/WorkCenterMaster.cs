using System;
using System.ComponentModel.DataAnnotations;

namespace Repository.Models
{
    public partial class WorkCenterMaster
    {
        [Key]
        public string WorkCenterCode { get; set; }
        public string GroupName { get; set; }
        public string GroupId { get; set; }
        public string Operation { get; set; }
        public string PlantCode { get; set; }
        public string WorkCenterActive { get; set; }
        public string RailOperation { get; set; }
        public string MachineType { get; set; }
    }
}
