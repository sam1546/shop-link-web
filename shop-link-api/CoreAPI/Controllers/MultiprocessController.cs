using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Repository.Interface;
using Repository.CustomViewModels;
using Microsoft.AspNetCore.Cors;
using Repository.Models;
using Repository.Repositories;
using Microsoft.Extensions.Configuration;
using Microsoft.AspNetCore.Hosting;
using System.Net;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System;
using System.Xml;

namespace CoreAPI.Controllers
{
    [Produces("application/json")]
    [Route("api/Multiprocess")]
    public class MultiprocessController : Controller
    {
        WebRequest req = null;
        WebResponse rsp = null;
        XmlDocument xml1 = null;
        XmlNodeList nodes = null;
        CredentialCache cc = null;
        NetworkCredential credentials = null;
        StreamWriter writer = null;
        StreamReader sr = null;
        DataSet ds = null;

        string result, uri, WCAssign, POConfirm, SPOConfirm, SecondaryWc = "";
        string UserName = "PP1_SHOPLINK", Password = "Sapprd@1";

        public TimeSpan shift1start = new TimeSpan(0, 0, 0, 0, 0);
        public TimeSpan shift2start = new TimeSpan(0, 0, 0, 0, 0);
        public TimeSpan shift3start = new TimeSpan(0, 0, 0, 0, 0);

        //public IOperation operationRepo { get; set; }
        public IConfiguration Configuration;
        private readonly IHostingEnvironment _env;

        SqlConnection Conn;
        SqlConnection cn;
        SqlCommand com = new SqlCommand();
        public MultiprocessController(IConfiguration _configuration, IHostingEnvironment env)
        {
            //operationRepo = _dashboardRepo;
            Configuration = _configuration;
            _env = env;

            WCAssign = _env.ContentRootPath + "//files//xml//Shoplink_WorkCenterAssignment.xml";
            POConfirm = _env.ContentRootPath + "//files//xml//Shoplink_ProductionOrderConfirmation.xml";
            SPOConfirm = _env.ContentRootPath + "//files//xml//Shoplink_SECProductionOrderConfirmation.xml";
            SecondaryWc = _env.ContentRootPath + "//files//xml//Secondary_WorkCenterAssignment.xml";

            Conn = new SqlConnection(Configuration.GetConnectionString("DefaultConnection").ToString());
            cn = new SqlConnection(Configuration.GetConnectionString("DefaultConnection").ToString());
        }

        public class jsonResult
        {
            public string incorrectPOs { get; set; }
            public int totalupdates { get; set; }
            public string error { get; set; }
            public int errorCode { get; set; }
            public string message { get; set; }
        }

        private string GetTextFromXMLFile(string file)
        {
            StreamReader reader = new StreamReader(file);
            string ret = reader.ReadToEnd();
            reader.Close();
            return ret;
        }

        [HttpGet]
        [Route("reversePO")]
        // GET: /<controller>/
        public async Task<IActionResult> ReversePO(string[] fileSplit)
        {
            jsonResult jsonResult = new jsonResult();

            int totalupdates = 0;
            string incorrectPOs = string.Empty;
            for (int j = 0; j < fileSplit.Length; j++)
            {
                fileSplit[j] = fileSplit[j].Trim();
                if (fileSplit[j].Contains("\t"))
                    fileSplit[j] = fileSplit[j].Replace("\t", "");
                fileSplit[j] = fileSplit[j].PadLeft(12, '0');

                long num = 0;
                if (long.TryParse(fileSplit[j], out num))
                { }
                else
                {
                    incorrectPOs += incorrectPOs + ":(Incorrect PO) ,"; 
                    continue; ;
                }
                SqlCommand check;
                SqlDataReader drcheck;
                string MachineName = "";
                check = new SqlCommand("select * from Operations where RSNo='" + fileSplit[j] + "'   ", cn);
                drcheck = check.ExecuteReader();
                if (drcheck.Read())
                {
                    string ack = drcheck["Flag_Ack"].ToString();
                    MachineName = drcheck["PrimaryWC"].ToString();
                    drcheck.Close();
                    if (ack == "")
                    {
                        incorrectPOs += incorrectPOs + ":(NO ACK.) ,";
                        continue;
                    }
                }
                else
                {
                    drcheck.Close();
                    incorrectPOs += incorrectPOs + ":(No Records) ,";
                    continue;
                }
                if (MachineName == "")
                {
                    incorrectPOs += incorrectPOs + ":(NO. WC.) ,";
                    continue;
                }
                SqlCommand cmdc = new SqlCommand("delete from Production_Details where RSNo='" + fileSplit[j] + "' ", cn);
                cmdc.ExecuteNonQuery();
                {
                }
                SqlCommand c1 = new SqlCommand();
                XmlDocument xml1 = new XmlDocument();
                xml1.Load(WCAssign);
                nodes = xml1.SelectNodes("//WADOCO");
                foreach (XmlElement element1 in nodes)
                {
                    element1.InnerText = fileSplit[j].Trim();
                    xml1.Save(WCAssign);
                }
                nodes = xml1.SelectNodes("//Machine");
                foreach (XmlElement element1 in nodes)
                {
                    element1.InnerText = MachineName;
                    xml1.Save(WCAssign);
                }

                int i = 0;
                nodes = xml1.SelectNodes("//IROPNO");
                foreach (XmlElement element1 in nodes)
                {
                    int status = 0;
                    if (i == 0)
                        status = 10;
                    else if (i == 1)
                        status = 40;
                    else
                        status = 50;
                    element1.InnerText = status.ToString();
                    xml1.Save(WCAssign);
                    i++;
                }
                try
                {
                    sr = new StreamReader(WCAssign);
                    uri = sr.ReadLine();
                    sr.Close();
                    // uri = "http://kecpodpp1app.hec.kecrpg.com:50000/XISOAPAdapter/MessageServlet?senderParty=&senderService=BC_SHOPLINK&receiverParty=&receiverService=&interface=SI_S_WorkCenterAssignment&interfaceNamespace=urn:kecrpg.com:HANA:PlanToProduce/ShopLinkWorkCenterAssignment";

                    credentials = new NetworkCredential(UserName, Password);
                    cc = new CredentialCache();
                    cc.Add(new Uri(uri), "Basic", credentials);
                    req = WebRequest.Create(uri);
                    req.Method = "POST";
                    req.ContentType = "text/xml";
                    writer = new StreamWriter(req.GetRequestStream());
                    writer.WriteLine(this.GetTextFromXMLFile(WCAssign));
                    writer.Close();
                    req.Credentials = cc;
                    rsp = req.GetResponse();
                    sr = new StreamReader(rsp.GetResponseStream());
                    result = sr.ReadToEnd();
                    sr.Close();
                }
                catch (Exception ex)
                {
                    jsonResult.error = "Exception: " + ex.Message;
                    jsonResult.errorCode = 404;
                    return Ok(jsonResult);
                }

                c1 = new SqlCommand("update Operations set MachineName='" + MachineName + "' where RSNo='" + fileSplit[j] + "' and POType='Primary' ", cn);
                c1.ExecuteNonQuery();

                c1 = new SqlCommand("delete from  Operations  where RSNo='" + fileSplit[j] + "' and POType='Duplicate' ", cn);
                c1.ExecuteNonQuery();

                totalupdates = totalupdates + 1;

            }
            jsonResult.totalupdates = totalupdates;
            jsonResult.incorrectPOs = incorrectPOs;
            jsonResult.message = "Total " + totalupdates + " WorkOrders reversed from shoplink";
            return Ok(jsonResult);
        }




    }
}
