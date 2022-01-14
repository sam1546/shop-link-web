using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Repository.Interface;
using Microsoft.Extensions.Configuration;
using System.Net;
using System.Xml;
using System.IO;
using Microsoft.AspNetCore.Hosting;
using System.Data;
using System.Configuration;
using Newtonsoft.Json;
// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace CoreAPI.Controllers
{
    [Produces("application/json")]
    [Route("api/Dashboard/")]
    public class DashboardController : Controller
    {
        public IOperation operationRepo { get; set; }
        public IConfiguration Configuration;
        WebRequest req = null;
        WebResponse rsp = null;
        XmlDocument xml1 = null;
        XmlNodeList nodes = null;
        string uri = "";
        CredentialCache cc = null;
        NetworkCredential credentials = null;
        string UserName = "";
        string Password = "";
        public static string loginuser = "";

        //PD1_SMARTCAM Welcome1234
        StreamWriter writer = null;
        StreamReader sr = null;
        string result = "";

        private readonly IHostingEnvironment _env;
        string Request, Response, ShopAck, WCAssign, GWCAssign, PaintWCAssign, WeldWCAssign = "";
        string ShoplinkMirUrl = "";


        DataSet ds = null;
        public static string[] machineGroup = { "Group1" };
        public TimeSpan shift1start = new TimeSpan(0, 0, 0, 0, 0);
        public TimeSpan shift2start = new TimeSpan(0, 0, 0, 0, 0);
        public TimeSpan shift3start = new TimeSpan(0, 0, 0, 0, 0);

        //MachinePlan login1 = new MachinePlan();
        string UserRole = "";
        SqlConnection Conn;
        SqlCommand com = new SqlCommand();
        public DashboardController(IOperation _dashboardRepo, IConfiguration _configuration, IHostingEnvironment env)
        {
            operationRepo = _dashboardRepo;
            Configuration = _configuration;
            _env = env;

            Request = _env.ContentRootPath + "//files//xml//Request.xml";
            Response = _env.ContentRootPath + "//files//xml//Response.xml";
            ShopAck = _env.ContentRootPath + "//files//xml//Shoplink_Acknowledgement.xml";
            WCAssign = _env.ContentRootPath + "//files//xml//Shoplink_WorkCenterAssignment.xml";
            GWCAssign = _env.ContentRootPath + "//files//xml//Galvanization_WorkCenterAssignment.xml";
            PaintWCAssign = _env.ContentRootPath + "//files//xml//PAINT_WorkCenterAssignment.xml";
            WeldWCAssign = _env.ContentRootPath + "//files//Weld_WorkCenterAssignment.xml";

            ShoplinkMirUrl = _env.ContentRootPath + "//files//ini//ShoplinkMirUrl.ini";

            Conn = new SqlConnection(Configuration.GetConnectionString("DefaultConnection").ToString());
        }

        public string GetTextFromXMLFile(string file)
        {
            StreamReader reader = new StreamReader(file);
            string ret = reader.ReadToEnd();
            reader.Close();
            return ret;
        }

        [HttpGet]
        [Route("getOperationByMirno")]
        public async Task<IActionResult> getOperationByMirno(string mirno)
        {
            Conn.Open();
            SqlCommand cmd2 = new SqlCommand("select *  from Operations where Mirno='" + mirno + "' ", Conn);
            SqlDataReader dr1 = cmd2.ExecuteReader();
            if (dr1.Read())
            {
                var res = Ok(dr1["BP"].ToString().Trim());
                dr1.Close();
                Conn.Close();
                return res;
            }
            else
            {
                dr1.Close();
                Conn.Close();
                return Ok("");
            }
        }

        [HttpGet]
        [Route("findDB")]
        public async Task<IActionResult> findDB(string Mirno)
        {
            int totalupdates = 0;
            string wammcu = "", waitm = "", BranchCode = "", item = "", lot = "", date = "", element = "", BP = "", billable = "", profile = "", RSNo = "", ixitm = "", ang = "", Records = "Fail", Diameter = "";
            float pices = 0, perWheight = 0, len = 0, totalWt = 0, TotalQty = 0;
            int operation = 0, RSNumber = 0, status = 0, TotalWO = 0, ReleasedWO = 0;

            SqlCommand cmd2 = new SqlCommand("select *  from Operations where Mirno='" + Mirno + "' ", Conn);
            SqlDataReader dr1 = cmd2.ExecuteReader();
            if (dr1.Read())
            {
                xml1 = new XmlDocument();
                xml1.Load(Request);
                nodes = xml1.SelectNodes("//WADOCO");
                foreach (XmlElement element1 in nodes)
                {
                    element1.InnerText = Mirno;
                    xml1.Save(Request);
                }
                sr = new StreamReader(ShoplinkMirUrl);
                uri = sr.ReadLine();
                sr.Close();
                try
                {
                    credentials = new NetworkCredential(UserName, Password);
                    cc = new CredentialCache();
                    cc.Add(new Uri(uri), "Basic", credentials);
                    req = WebRequest.Create(uri);
                    req.Method = "POST";        // Post method
                    req.ContentType = "text/xml";     // content type
                    writer = new StreamWriter(req.GetRequestStream());
                    writer.WriteLine(this.GetTextFromXMLFile(Request));
                    writer.Close();
                    req.Credentials = cc;
                    rsp = req.GetResponse();
                    sr = new StreamReader(rsp.GetResponseStream());
                    result = sr.ReadToEnd();
                    sr.Close();
                }
                catch (Exception e)
                {
                    return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                }

                System.IO.File.WriteAllText(Response, result);

                string opsq = "";
                string OPStatus = "";
                float wauorg = 0, wheight = 0;

                #region countTotalPO

                using (XmlReader reader = XmlReader.Create(Response))
                {
                    while (reader.Read())
                    {
                        if (reader.IsStartElement())
                        {
                            switch (reader.Name.ToString())
                            {
                                case "WADOCO":
                                    if (!reader.IsEmptyElement)
                                    {
                                        RSNo = reader.ReadString();

                                        TotalWO += 1;
                                    }
                                    break;
                            }
                        }
                    }
                }
                #endregion
                using (XmlReader reader = XmlReader.Create(Response))
                {
                    while (reader.Read())
                    {
                        if (reader.IsStartElement())
                        {
                            switch (reader.Name.ToString())
                            {
                                case "WADOCO":
                                    if (!reader.IsEmptyElement)
                                    {
                                        RSNo = reader.ReadString();
                                    }
                                    break;

                                case "WALITM":
                                    if (!reader.IsEmptyElement)
                                        item = reader.ReadString();
                                    break;

                                case "WAUORG":
                                    if (!reader.IsEmptyElement)
                                    {
                                        wauorg = float.Parse(reader.ReadString());
                                        pices = wauorg;
                                    }
                                    break;

                                case "IXY55THW":
                                    if (!reader.IsEmptyElement)
                                    {
                                        wheight = float.Parse(reader.ReadString());
                                        perWheight = wheight;
                                    }
                                    break;

                                case "WASRST":
                                    if (!reader.IsEmptyElement)
                                        status = int.Parse(reader.ReadString());
                                    break;

                                case "WAMMCU":
                                    if (!reader.IsEmptyElement)
                                    {
                                        wammcu = reader.ReadString();
                                        BP = wammcu;
                                    }
                                    break;

                                case "WARORN":
                                    if (!reader.IsEmptyElement)
                                        Mirno = reader.ReadString();
                                    break;

                                case "IXY55RML":
                                    if (!reader.IsEmptyElement)
                                        len = float.Parse(reader.ReadString());
                                    break;

                                case "IROPSQ":
                                    if (!reader.IsEmptyElement)
                                        opsq = reader.ReadString();
                                    break;

                                case "IXY55BM9":
                                    if (!reader.IsEmptyElement)
                                    {
                                        operation = int.Parse(reader.ReadString());
                                        operation = operation + 2;
                                    }
                                    break;
                                case "IXLITM":
                                    if (!reader.IsEmptyElement)
                                        profile = reader.ReadString();
                                    break;

                                case "WALOTN":
                                    if (!reader.IsEmptyElement)
                                        lot = reader.ReadString();
                                    break;

                                case "DIAMTR":
                                    if (!reader.IsEmptyElement)
                                        Diameter = reader.ReadString();
                                    break;
                            }

                            if (reader.Name.ToString() == "DIAMTR")
                            {
                                ReleasedWO += 1;
                                Records = "Success";
                                totalWt = pices * perWheight;
                                TotalQty = pices * operation;
                                DateTime AckDate = DateTime.Now;
                                string AckJdDate = AckDate.Month + "/" + AckDate.Day + "/" + AckDate.Year;

                                com = new SqlCommand("select * from operations where RSNo='" + RSNo + "'", Conn);
                                SqlDataReader dr = com.ExecuteReader();
                                if (!dr.Read())
                                {
                                    dr.Close();
                                    com = new SqlCommand("insert into Operations(RSNo,FGItem,Operation,Pices,Wheight,Setups,SctDinemtion,Status,Mirno,LotCode,BP,TotalWt,Tot_OPS,Length,OPStatus,Bal_Fab,Bal_Notch,Bal_Weld,Bal_Bend,Bal_HM,Bal_Galva,POType,Diameter,SAPPulledDate) values('" + RSNo + "','" + item + "'," + operation + "," + pices + ",'" + perWheight + "',1,'" + profile + "','" + status + "','" + Mirno + "','" + lot + "','" + BP + "'," + totalWt + "," + TotalQty + "," + len + ",'" + opsq + "'," + pices + "," + pices + "," + pices + "," + pices + "," + pices + "," + pices + ",'Primary','" + Diameter + "','" + AckJdDate + "')", Conn);
                                    com.ExecuteNonQuery();
                                    SqlCommand Bal = new SqlCommand();
                                    if (!opsq.Contains("N"))
                                    {
                                        Bal = new SqlCommand("update Operations set Bal_Notch=0 where RSNo='" + RSNo + "'", Conn);
                                        Bal.ExecuteNonQuery();
                                    }
                                    else
                                    {
                                        Bal = new SqlCommand("update Operations set Flag_Notch='FALSE' where RSNo='" + RSNo + "'", Conn);
                                        Bal.ExecuteNonQuery();
                                    }
                                    if (!opsq.Contains("M"))
                                    {
                                        Bal = new SqlCommand("update Operations set Bal_HM=0 where RSNo='" + RSNo + "'", Conn);
                                        Bal.ExecuteNonQuery();
                                    }
                                    else
                                    {
                                        Bal = new SqlCommand("update Operations set Flag_HM='FALSE' where RSNo='" + RSNo + "'", Conn);
                                        Bal.ExecuteNonQuery();
                                    }
                                    if (!opsq.Contains("B"))
                                    {
                                        Bal = new SqlCommand("update Operations set Bal_Bend=0 where RSNo='" + RSNo + "'", Conn);
                                        Bal.ExecuteNonQuery();
                                    }
                                    else
                                    {
                                        Bal = new SqlCommand("update Operations set Flag_Bend='FALSE' where RSNo='" + RSNo + "'", Conn);
                                        Bal.ExecuteNonQuery();
                                    }
                                    if (!opsq.Contains("W"))
                                    {
                                        Bal = new SqlCommand("update Operations set Bal_Weld=0 where RSNo='" + RSNo + "'", Conn);
                                        Bal.ExecuteNonQuery();
                                    }
                                    else
                                    {
                                        Bal = new SqlCommand("update Operations set Flag_Weld='FALSE' where RSNo='" + RSNo + "'", Conn);
                                        Bal.ExecuteNonQuery();
                                    }
                                }
                                dr.Close();
                                opsq = "";
                                operation = 0;
                                lot = "";
                                profile = "";
                                len = 0;
                                BP = "";
                                status = 0;
                                perWheight = 0;
                                pices = 0;
                                wauorg = 0;
                                item = "";
                                RSNo = "";
                                Diameter = "";
                            }
                        }
                    }
                }
                if (Records == "Fail")
                {
                    return Ok("No Records found");
                }
            }
            else
            {
                if (Mirno != "")
                {
                    try
                    {
                        xml1 = new XmlDocument();
                        xml1.Load(Request);
                        nodes = xml1.SelectNodes("//WADOCO");
                        foreach (XmlElement element1 in nodes)
                        {
                            element1.InnerText = Mirno;
                            xml1.Save(Request);
                        }
                        sr = new StreamReader(ShoplinkMirUrl);
                        uri = sr.ReadLine();
                        sr.Close();
                        credentials = new NetworkCredential(UserName, Password);
                        cc = new CredentialCache();
                        cc.Add(new Uri(uri), "Basic", credentials);
                        req = WebRequest.Create(uri);
                        req.Method = "POST";        // Post method
                        req.ContentType = "text/xml";     // content type
                        writer = new StreamWriter(req.GetRequestStream());
                        // Write the XML text into the stream
                        writer.WriteLine(this.GetTextFromXMLFile(Request));
                        // string xml=GetTextFromXMLFile(fileName);
                        writer.Close();
                        // Send the data to the webserver
                        req.Credentials = cc;
                        //req.Timeout = 1000;
                        rsp = req.GetResponse();
                        sr = new StreamReader(rsp.GetResponseStream());
                        result = sr.ReadToEnd();
                        sr.Close();
                    }
                    catch (Exception e)
                    {
                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                    }
                    System.IO.File.WriteAllText(Response, result);
                    string opsq = "";
                    string OPStatus = "";
                    float wauorg = 0, wheight = 0;

                    #region count Total PO

                    using (XmlReader reader = XmlReader.Create(Response))
                    {
                        while (reader.Read())
                        {
                            if (reader.IsStartElement())
                            {
                                switch (reader.Name.ToString())
                                {
                                    case "WADOCO":
                                        if (!reader.IsEmptyElement)
                                        {
                                            RSNo = reader.ReadString();

                                            TotalWO += 1;
                                        }
                                        break;
                                }
                            }
                        }
                    }

                    #endregion

                    using (XmlReader reader = XmlReader.Create(Response))
                    {
                        while (reader.Read())
                        {
                            if (reader.IsStartElement())
                            {
                                switch (reader.Name.ToString())
                                {
                                    case "WADOCO":
                                        if (!reader.IsEmptyElement)
                                        {
                                            RSNo = reader.ReadString();
                                        }
                                        break;
                                    case "WALITM":
                                        if (!reader.IsEmptyElement)
                                            item = reader.ReadString();
                                        break;
                                    case "WAUORG":
                                        if (!reader.IsEmptyElement)
                                        {
                                            wauorg = float.Parse(reader.ReadString());
                                            pices = wauorg;
                                        }
                                        break;
                                    case "IXY55THW":

                                        if (!reader.IsEmptyElement)
                                        {
                                            wheight = float.Parse(reader.ReadString());
                                            perWheight = wheight;
                                        }
                                        break;
                                    case "WASRST":
                                        if (!reader.IsEmptyElement)
                                            status = int.Parse(reader.ReadString());
                                        break;
                                    case "WAMMCU":
                                        if (!reader.IsEmptyElement)
                                        {
                                            wammcu = reader.ReadString();
                                            BP = wammcu;
                                        }
                                        break;
                                    case "WARORN":
                                        if (!reader.IsEmptyElement)
                                            Mirno = reader.ReadString();
                                        break;
                                    case "IXY55RML":
                                        if (!reader.IsEmptyElement)
                                            len = float.Parse(reader.ReadString());
                                        break;
                                    case "IROPSQ":
                                        if (!reader.IsEmptyElement)
                                            opsq = reader.ReadString();
                                        break;

                                    case "IXY55BM9":
                                        if (!reader.IsEmptyElement)
                                        {
                                            operation = int.Parse(reader.ReadString());
                                            operation = operation + 2;
                                        }
                                        break;
                                    case "IXLITM":
                                        if (!reader.IsEmptyElement)
                                            profile = reader.ReadString();
                                        break;
                                    case "WALOTN":
                                        if (!reader.IsEmptyElement)
                                            lot = reader.ReadString();
                                        break;
                                    case "DIAMTR":
                                        if (!reader.IsEmptyElement)
                                            Diameter = reader.ReadString();
                                        break;
                                }
                                //IROPSQ
                                if (reader.Name.ToString() == "DIAMTR")
                                {
                                    ReleasedWO += 1;
                                    Records = "Success";
                                    //testing
                                    totalWt = pices * perWheight;
                                    TotalQty = pices * operation;
                                    DateTime AckDate = DateTime.Now;
                                    string AckJdDate = AckDate.Month + "/" + AckDate.Day + "/" + AckDate.Year;
                                    //SqlCommand cmd1 = new SqlCommand();    
                                    SqlCommand com = new SqlCommand("insert into Operations(RSNo,FGItem,Operation,Pices,Wheight,Setups,SctDinemtion,Status,Mirno,LotCode,BP,TotalWt,Tot_OPS,Length,OPStatus,Bal_Fab,Bal_Notch,Bal_Weld,Bal_Bend,Bal_HM,Bal_Galva,POType,Diameter,SAPPulledDate) values('" + RSNo + "','" + item + "'," + operation + "," + pices + ",'" + perWheight + "',1,'" + profile + "','" + status + "','" + Mirno + "','" + lot + "','" + BP + "'," + totalWt + "," + TotalQty + "," + len + ",'" + opsq + "'," + pices + "," + pices + "," + pices + "," + pices + "," + pices + "," + pices + ",'Primary','" + Diameter + "','" + AckJdDate + "')", Conn);
                                    com.ExecuteNonQuery();

                                    SqlCommand Bal = new SqlCommand();
                                    if (!opsq.Contains("N"))
                                    {
                                        Bal = new SqlCommand("update Operations set Bal_Notch=0 where RSNo='" + RSNo + "'", Conn);
                                        Bal.ExecuteNonQuery();
                                    }
                                    else
                                    {
                                        Bal = new SqlCommand("update Operations set Flag_Notch='FALSE' where RSNo='" + RSNo + "'", Conn);
                                        Bal.ExecuteNonQuery();
                                    }
                                    if (!opsq.Contains("M"))
                                    {
                                        Bal = new SqlCommand("update Operations set Bal_HM=0 where RSNo='" + RSNo + "'", Conn);
                                        Bal.ExecuteNonQuery();
                                    }
                                    else
                                    {
                                        Bal = new SqlCommand("update Operations set Flag_HM='FALSE' where RSNo='" + RSNo + "'", Conn);
                                        Bal.ExecuteNonQuery();
                                    }
                                    if (!opsq.Contains("B"))
                                    {
                                        Bal = new SqlCommand("update Operations set Bal_Bend=0 where RSNo='" + RSNo + "'", Conn);
                                        Bal.ExecuteNonQuery();
                                    }
                                    else
                                    {
                                        Bal = new SqlCommand("update Operations set Flag_Bend='FALSE' where RSNo='" + RSNo + "'", Conn);
                                        Bal.ExecuteNonQuery();
                                    }
                                    if (!opsq.Contains("W"))
                                    {
                                        Bal = new SqlCommand("update Operations set Bal_Weld=0 where RSNo='" + RSNo + "'", Conn);
                                        Bal.ExecuteNonQuery();
                                    }
                                    else
                                    {
                                        Bal = new SqlCommand("update Operations set Flag_Weld='FALSE' where RSNo='" + RSNo + "'", Conn);
                                        Bal.ExecuteNonQuery();
                                    }

                                    opsq = "";
                                    operation = 0;
                                    lot = "";
                                    profile = "";
                                    len = 0;
                                    BP = "";
                                    status = 0;
                                    perWheight = 0;
                                    pices = 0;
                                    wauorg = 0;
                                    item = "";
                                    RSNo = "";
                                    Diameter = "";
                                }
                            }
                        }
                    }
                    if (Records == "Fail")
                    {
                        return Ok("No Records found");
                    }
                }
            }

            //double wheight1 = 0.0, Operations = 0.0, Runtime1 = 0.0;
            //int totalWorkorders = 0;
            //SqlCommand c1 = new SqlCommand("select sum(TotalWt)/1000 as TotalWheight,count(RSNo) as RSno,sum(Tot_OPS) as TotalOpns,sum(RunTime) as RunTime from Operations where Mirno='" + Mirno + "' and BP='" + BP + "' and POType='Primary' ", Conn);
            //SqlDataReader drmir = c1.ExecuteReader();

            SqlCommand myCommand = new SqlCommand("select sum(TotalWt)/1000 as TotalWheight,count(RSNo) as RSno,sum(Tot_OPS) as TotalOpns,sum(RunTime) as RunTime from Operations where Mirno='" + Mirno + "' and BP='" + BP + "' and POType='Primary' ", Conn);
            SqlDataAdapter myAdapter = new SqlDataAdapter();
            myAdapter.SelectCommand = myCommand;
            DataSet myDataSet = new DataSet();
            Conn.Open();
            myAdapter.Fill(myDataSet);
            Conn.Close();
            //string json = JsonConvert.SerializeObject(myDataSet.Tables[0], Newtonsoft.Json.Formatting.Indented);
            return Ok(myDataSet.Tables[0]);
              
            //if (drmir.Read())
            //{
            //    if (drmir["TotalWheight"].ToString() != "")
            //        wheight1 = Math.Round(double.Parse(drmir["TotalWheight"].ToString()), 3);
            //    if (drmir["RSno"].ToString() != "")
            //        totalWorkorders = int.Parse(drmir["RSno"].ToString());

            //    if (drmir["TotalOpns"].ToString() != "")
            //        Operations = double.Parse(drmir["TotalOpns"].ToString());
            //    if (drmir["RunTime"].ToString() != "")
            //        Runtime1 = Math.Round(double.Parse(drmir["RunTime"].ToString()), 2);
            //}
            //txtwheight.Text = wheight1.ToString();
            //txtOprns.Text = Operations.ToString();
            //txtRunTime.Text = Runtime1.ToString();
            //txtRs.Text = totalWorkorders.ToString();
            //drmir.Close();
            //SetLoading(false);
            //BindDataGridALLRecords(Mirno, BP);

            //MessageBox.Show(totalWorkorders + " Production Orders added, MIR No: " + txtAckMirno.Text);

            //lblUpdationBar.Visible = false;
            //progressBar2.Visible = false;

            //return Ok(result);
        }



        [HttpGet]
        [Route("loadWorkCenters")]
        public async Task<IActionResult> LoadWorkCenters(string group)
        {
            var res = await operationRepo.GetWorkCeters(group);
            return Ok(res);
        }

        [HttpGet]
        [Route("loadGroups")]
        public async Task<IActionResult> LoadGroups()
        {
            var res = await operationRepo.GetGroups();
            return Ok(res);
        }

        [HttpGet]
        [Route("bindDataGrid")]
        public async Task<IActionResult> BindDataGrid(string mirno, string plantCode)
        {
            string CommandText = "SELECT RSNo as RSNo,JDDate,FGItem as ElementNo,Mirno as MIRNO,Pices as QTY,Wheight,TotalWt,Operation,Tot_OPS,Length,SctDinemtion as SECTION,LotCode as Billable_Lot, Status FROM Operations where Mirno='" + mirno + "' and BP='" + plantCode + "' and Flag_Fab is null order by [index] desc ";
            SqlCommand myCommand = new SqlCommand(CommandText, Conn);
            SqlDataAdapter myAdapter = new SqlDataAdapter();
            myAdapter.SelectCommand = myCommand;
            DataSet myDataSet = new DataSet();
            Conn.Open();
            myAdapter.Fill(myDataSet);
            Conn.Close();
            //string json = JsonConvert.SerializeObject(myDataSet.Tables[0], Newtonsoft.Json.Formatting.Indented);
            return Ok(myDataSet.Tables[0]);
        }

        [HttpGet]
        [Route("bindDataGridALLRecords")]
        public async Task<IActionResult> BindDataGridALLRecords(string mirno, string plantCode) 
        { 
                string CommandText = "SELECT RSNo as rsNo,FGItem as Item_No,Mirno as mirNo,MachineName as WorkCenter,SctDinemtion as SECTION,OPStatus,LotCode as billable_Lot,Pices as QTY,Length,Wheight as Wt_Pcs, FORMAT (JDDate, 'dd/MM/yyyy ')  as ReleasedDate,TotalWt,RackDetails,Operation as OPRPCs,Tot_OPS,PlanningShift, FORMAT (PlanningDate, 'dd/MM/yyyy ') as PlanningDate,RunTime,Status FROM Operations  where BP='" + plantCode + "' and Mirno='"+mirno+"' order by [index] desc "; //Diameter,SAPPulledDate as PulledDate,

            SqlCommand myCommand = new SqlCommand(CommandText, Conn);
                SqlDataAdapter myAdapter = new SqlDataAdapter();
                myAdapter.SelectCommand = myCommand;
                DataSet myDataSet = new DataSet();
                if (Conn.State == ConnectionState.Closed)
                    Conn.Open();
                myAdapter.Fill(myDataSet);
                Conn.Close();
                return Ok(myDataSet.Tables[0]); 
        }

        // GET api/values/5
        [HttpGet("{id}")]
        public string Get(int id)
        {
            return "value sammm";
        }

        // POST api/values
        [HttpPost]
        public void Post([FromBody] string value)
        {
        }

        // PUT api/values/5
        [HttpPut("{id}")]
        public void Put(int id, [FromBody] string value)
        {
        }

        // DELETE api/values/5
        [HttpDelete("{id}")]
        public void Delete(int id)
        {
        }
    }
}
