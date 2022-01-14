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
using System.Data.SqlClient; 
using System.Xml;
using Microsoft.Extensions.Configuration;
using System.Configuration;
namespace Repository.Repositories
{
    public class OperationRepo:IOperation
    {
        private readonly Context _context;
        private IConfiguration Configuration;
        XmlDocument xml1 = null;

        public OperationRepo(Context context, IConfiguration _configuration)
        {
            _context = context;
            Configuration = _configuration;
        }

        public async Task<List<WorkCenterMaster>> GetGroups()
        {
            List<WorkCenterMaster> workCenters = new List<WorkCenterMaster>();
            try
            {
                var query = (from pro in _context.WorkCenterMaster
                            select new { pro.GroupId }).Distinct().ToList();
                if (query == null)
                {
                    return null;
                }
                else
                {
                    foreach (var item in query)
                    {
                        WorkCenterMaster wc = new WorkCenterMaster();
                        wc.GroupId = item.GroupId;
                        workCenters.Add(wc);
                    }
                }
            }
            catch (Exception ex)
            {
            }
            return workCenters;
        }
 
        public async Task<List<WorkCenterMaster>> GetWorkCeters(string group)
        {
            List<WorkCenterMaster> workCenters = new List<WorkCenterMaster>();
            try
            {
                var query = (from pro in _context.WorkCenterMaster
                             where pro.GroupId == @group
                            select new { pro.WorkCenterCode}).Distinct().ToList();
                if (query == null)
                {
                    return null;
                }
                else
                {
                    foreach (var item in query)
                    {
                        WorkCenterMaster wc = new WorkCenterMaster();
                        wc.WorkCenterCode = item.WorkCenterCode; 
                        workCenters.Add(wc); 
                    }
                }
            }
            catch (Exception ex)
            {
            }
            return workCenters;
        }
    }
}