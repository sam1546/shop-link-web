using System;
using System.ComponentModel.DataAnnotations;

namespace Repository.Models
{
    public partial class Operations
    {
        [Key]
        public int index { get; set; }
        public string RSNo { get; set; }
        public string FGItem { get; set; }
        public int? Operation { get; set; }
        public double? Pices { get; set; }
        public string ActualPiece { get; set; }
        public double? Wheight { get; set; }
        public int Setups { get; set; }
        public string SctDinemtion { get; set; }
        public DateTime? PlanningDate { get; set; }
        public string PlanningShift { get; set; }
        public string Operator { get; set; }
        public int? Status { get; set; }
        public string Mirno { get; set; } 
        public string LotCode { get; set; }
        public string BP { get; set; }
        public string CountStatus { get; set; }
        public double? TotalWt { get; set; }
        public double? Tot_OPS { get; set; }
        public double? Length { get; set; }
        public string MachineName { get; set; }
        public DateTime? JDDate { get; set; }
        public int? CompletedQty { get; set; }
        public string OPStatus { get; set; }
        public double? RunTime { get; set; }
        public string MachineGroup { get; set; }
        public string RackDetails { get; set; }
        public int? Bal_Fab { get; set; }
        public int? Bal_Notch { get; set; }
        public int? Bal_Weld { get; set; }
        public int? Bal_HM { get; set; }
        public int? Bal_Galva { get; set; }
        public int? Bal_Bend { get; set; }
        public int? Bal_Paint { get; set; }
        public int? Bal_TPI { get; set; }
        public string Flag_Notch { get; set; }
        public string Flag_Bend { get; set; }
        public string Flag_HM { get; set; }
        public string Flag_Weld { get; set; }
        public string Flag_Galva { get; set; }
        public string Flag_Paint { get; set; }
        public string Flag_Fab { get; set; }
        public string Flag_Ack { get; set; }
 
    }
}
