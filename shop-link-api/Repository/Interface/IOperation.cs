using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Repository.CustomViewModels;
using Repository.Models;

namespace Repository.Interface
{
    public interface IOperation
    {
        Task<List<WorkCenterMaster>> GetWorkCeters(string group);
        Task<List<WorkCenterMaster>> GetGroups(); 
    }
}
