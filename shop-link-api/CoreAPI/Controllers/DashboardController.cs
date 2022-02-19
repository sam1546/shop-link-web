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
using System.Threading;
using System.Text.RegularExpressions;
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

        string Request, Response, ShopAck, WCAssign, WCAssignini, GWCAssign, PaintWCAssign, WeldWCAssign = "";
        string ShoplinkMirUrl, ShoplinkAck, All, NB, NM, BM, Notch, HM, Bend = "";

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
            ShoplinkAck = _env.ContentRootPath + "//files//ini//ShoplinkAck.ini";
            All = _env.ContentRootPath + "//files//xml//All.ini";
            WCAssignini = _env.ContentRootPath + "//files//ini//WCAssign.ini";
            NB = _env.ContentRootPath + "//files//xml//NB.xml";
            NM = _env.ContentRootPath + "//files//xml//NM.xml";
            BM = _env.ContentRootPath + "//files//xml//BM.xml";
            Notch = _env.ContentRootPath + "//files//xml//Notch.xml";
            HM = _env.ContentRootPath + "//files//xml//HM.xml";
            Bend = _env.ContentRootPath + "//files//xml//Bend.xml";




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
        [Route("getBpByMirno")]
        public async Task<IActionResult> getBpByMirno(string mirno)
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

            SqlCommand myCommand = new SqlCommand("select sum(TotalWt)/1000 as TotalWheight,count(RSNo) as RSno,sum(Tot_OPS) as TotalOpns,sum(RunTime) as RunTime from Operations where Mirno='" + Mirno + "' and BP='" + BP + "' ", Conn); //and POType='Primary' 
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
        public async Task<IActionResult> BindDataGridALLRecords(string mirno, string plantCode, string poType)
        {
            string CommandText = "";
            if (poType == "Primary")
            {
                CommandText = "SELECT RSNo as rsNo,FGItem as Item_No,Mirno as mirNo,MachineName as WorkCenter,SctDinemtion as SECTION,OPStatus,LotCode as billable_Lot,Pices as QTY,Length,Wheight as Wt_Pcs, FORMAT (JDDate, 'dd/MM/yyyy ')  as ReleasedDate,TotalWt,RackDetails,Operation as OPRPCs,Tot_OPS,PlanningShift, FORMAT (PlanningDate, 'dd/MM/yyyy ') as PlanningDate,RunTime,Status FROM Operations where Mirno='" + mirno + "' and BP='" + plantCode + "'  order by [index] desc "; //and POType='Primary'
            }
            else
            {
                CommandText = "SELECT RSNo as rsNo,FGItem as Item_No,Mirno as mirNo,MachineName as WorkCenter,SctDinemtion as SECTION,OPStatus,LotCode as billable_Lot,Pices as QTY,Length,Wheight as Wt_Pcs, FORMAT (JDDate, 'dd/MM/yyyy ')  as ReleasedDate,TotalWt,RackDetails,Operation as OPRPCs,Tot_OPS,PlanningShift, FORMAT (PlanningDate, 'dd/MM/yyyy ') as PlanningDate,RunTime,Status FROM Operations  where BP='" + plantCode + "' and Mirno='" + mirno + "' order by [index] desc "; //Diameter,SAPPulledDate as PulledDate,
            }
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

        [HttpGet]
        [Route("getCalculations")]
        public async Task<IActionResult> GetCalculations(string mirno, string plantCode, string poType)
        {
            string CommandText = "";
            if (poType == "Primary")
            {
                CommandText = "select round(sum(TotalWt)/1000,3) as TotalWheight,count(RSNo) as RSno,sum(Tot_OPS) as TotalOpns,round(sum(RunTime),2) as RunTime from Operations where Mirno='" + mirno + "' and BP='" + plantCode + "'  ";  //and POType='Primary'
            }
            else
            {
                CommandText = "select round(sum(TotalWt)/1000,3) as TotalWheight,count(RSNo) as RSno,sum(Tot_OPS) as TotalOpns,round(sum(RunTime),2) as RunTime from Operations where Mirno='" + mirno + "' and BP = '" + plantCode + "' ";
            }

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

        [HttpGet]
        [Route("getOperationsByMirno")]
        public async Task<IActionResult> GetOperationsByMirno(string mirno)
        {
            string CommandText = "select *  from Operations where Mirno='" + mirno + "' ";
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

        [HttpGet]
        [Route("gettotalWO_Totalreleased")]
        public async Task<IActionResult> GettotalWO_Totalreleased(string mirno)
        {
            Boolean flag = false;
            SqlCommand check;
            int totalWO = 0, Totalreleased = 0;
            check = new SqlCommand("select count(RSNo) as TotalWo from Operations where Mirno='" + mirno + "'   ", Conn); //and POType='Primary'
            SqlDataReader drcount = check.ExecuteReader();
            if (drcount.Read())
            {
                totalWO = int.Parse(drcount["TotalWo"].ToString());
            }
            drcount.Close();

            check = new SqlCommand("select count(RSNo) as TotalWo from Operations where Mirno='" + mirno + "' and Flag_Ack='TRUE' ", Conn); //and POType='Primary' 
            SqlDataReader drack = check.ExecuteReader();
            if (drack.Read())
            {
                Totalreleased = int.Parse(drack["TotalWo"].ToString());
            }
            drack.Close();
            if (totalWO != Totalreleased)
            {
                flag = true;
            }
            return Ok(flag);

        }

        public class jsonResult {
            public int releasedPO { get; set; }
            public int totalPO { get; set; }
            public int Totalreleased { get; set; }
        }

        [HttpGet]
        [Route("releasePO")]
        public async Task<IActionResult> ReleasePO(string mirno, string plantCode)
        {
            jsonResult jsonResult = new jsonResult();
            int releasedPO = 0;
            try
            {
                int totalPO = 0, Totalreleased = 0;
                DateTime AckDate = DateTime.Now;
                string AckJdDate = AckDate.Month + "/" + AckDate.Day + "/" + AckDate.Year;
                SqlCommand cmd = new SqlCommand();
                if (Conn.State == ConnectionState.Closed)
                    Conn.Open();
                SqlCommand c1 = null;
                c1 = new SqlCommand("select count(RSNo) as RSno from Operations where Mirno='" + mirno + "' and BP='" + plantCode + "'  ", Conn); //and POType='Primary'
                SqlDataReader drmir2 = c1.ExecuteReader();
                if (drmir2.Read())
                {
                    totalPO = int.Parse(drmir2["RSno"].ToString());

                }
                drmir2.Close();

                c1 = new SqlCommand("select count(RSNo) as TotalWo from Operations where Mirno='" + mirno + "' and Flag_Ack='TRUE' ", Conn); //and POType='Primary' 
                SqlDataReader drack = c1.ExecuteReader();
                if (drack.Read())
                {
                    Totalreleased = int.Parse(drack["TotalWo"].ToString());
                }
                drack.Close();

                int totalPending = totalPO - Totalreleased;

                jsonResult.totalPO = totalPO;
                jsonResult.Totalreleased = Totalreleased;
                // lblUpdationBar.Visible = true;
                c1 = new SqlCommand("select RSNo from Operations where Mirno='" + mirno + "' and Flag_Ack is null", Conn);
                SqlDataReader dr = c1.ExecuteReader();
                while (dr.Read())
                {
                    try
                    {
                        string RSNo = dr["RSNo"].ToString();
                        xml1 = new XmlDocument();
                        xml1.Load(ShopAck);
                        nodes = xml1.SelectNodes("//WADOCO");
                        foreach (XmlElement element1 in nodes)
                        {
                            element1.InnerText = RSNo;
                            xml1.Save(ShopAck);
                        }
                        sr = new StreamReader(ShoplinkAck);
                        uri = sr.ReadLine();
                        sr.Close();

                        credentials = new NetworkCredential(UserName, Password);

                        cc = new CredentialCache();

                        cc.Add(new Uri(uri), "Basic", credentials);

                        req = WebRequest.Create(uri);
                        req.Method = "POST";
                        req.ContentType = "text/xml";
                        writer = new StreamWriter(req.GetRequestStream());
                        writer.WriteLine(this.GetTextFromXMLFile(ShopAck));


                        writer.Close();
                        req.Credentials = cc;
                        rsp = req.GetResponse();
                        sr = new StreamReader(rsp.GetResponseStream());
                        result = sr.ReadToEnd();
                        sr.Close();
                        Thread.Sleep(100);

                        cmd = new SqlCommand("update Operations set JDDate='" + AckJdDate + "', Flag_Ack='TRUE' where RSNo='" + RSNo + "'", Conn);
                        cmd.ExecuteNonQuery();

                        releasedPO += 1;
                    }
                    catch (Exception ex)
                    {
                        jsonResult.totalPO = 0;
                        jsonResult.Totalreleased = 0;
                        jsonResult.releasedPO = releasedPO;
                    }
                }
                jsonResult.releasedPO = releasedPO;
                dr.Close();
            }
            catch
            {
                jsonResult.totalPO = 0;
                jsonResult.Totalreleased = 0;
                jsonResult.releasedPO = releasedPO;
            }
            return Ok(jsonResult);
        }

        [HttpGet]
        [Route("ackPO")]
        public async Task<IActionResult> AckPO(string mirno)
        {
            string message = "";
            try
            {
                SqlCommand check;
                SqlDataReader drcheck;
                check = new SqlCommand("select * from Operations where RSNo='" + mirno + "'   ", Conn);
                drcheck = check.ExecuteReader();
                if (drcheck.Read())
                {
                    drcheck.Close();

                    xml1 = new XmlDocument();

                    xml1.Load(ShopAck);

                    nodes = xml1.SelectNodes("//WADOCO");
                    foreach (XmlElement element1 in nodes)
                    {
                        element1.InnerText = mirno.Trim();
                        xml1.Save(ShopAck);
                    }
                    sr = new StreamReader(ShoplinkAck);
                    uri = sr.ReadLine();
                    sr.Close();

                    credentials = new NetworkCredential(UserName, Password);

                    cc = new CredentialCache();

                    cc.Add(new Uri(uri), "Basic", credentials);

                    req = WebRequest.Create(uri);
                    req.Method = "POST";
                    req.ContentType = "text/xml";
                    writer = new StreamWriter(req.GetRequestStream());
                    writer.WriteLine(this.GetTextFromXMLFile(ShopAck));


                    writer.Close();
                    req.Credentials = cc;
                    rsp = req.GetResponse();
                    sr = new StreamReader(rsp.GetResponseStream());
                    result = sr.ReadToEnd();
                    sr.Close();

                    SqlCommand cmd = new SqlCommand("update Operations set Flag_Ack='TRUE' where RSNo='" + mirno + "'", Conn);
                    cmd.ExecuteNonQuery();
                    message = "Production Order Number " + mirno + " Acknowledged in SAP";
                }
                else
                {
                    message = "Data is not avaalable in Shoplink";
                }
            }
            catch (Exception ex)
            {
                message = ex.Message.ToString();
            }
            return Ok(message);
        }


        [HttpGet]
        [Route("allocate")]
        public async Task<IActionResult> Allocate(string mirno, string plantCode, string comboBox_MachineName, string cmb_Group, string cmbShift, string txtRack)
        {
            //enter
            try
            {
                bool showMessage = true;
                int totalPO = 0, PendingPO = 0, completedPO = 0, releasedWC = 0; ;
                double wheight = 0, Operations = 0, Runtime1 = 0, xfactore = 0, Thickness = 0, GripAllow = 0, CarriageSpeed = 0, YFactoreConst = 0, FactoreSpeed = 0, punchpara = 0, cutpara = 0, stampara = 0, yfactore = 0, punch = 0, cutting = 0, stamping = 0, carriage = 0, gripping = 0, length = 0, pices = 0, operations = 0, setup = 0, inspection = 0, Runtime = 0;
                string routeSheet = "", section = "", maachinetype = "", workcenterName = "";
                SqlCommand c1 = new SqlCommand();
                SqlCommand c2 = new SqlCommand();
                SqlDataReader drpara;

                if (Conn.State == ConnectionState.Closed)
                    Conn.Open();
                SqlCommand check;
                SqlDataReader drcheck;
                check = new SqlCommand("");
                check = new SqlCommand("select * from Operations where Mirno='" + mirno + "'   ", Conn); //and POType='Primary'
                drcheck = check.ExecuteReader();
                if (drcheck.Read())
                {
                    string ack = drcheck["Flag_Ack"].ToString();
                    string BP = drcheck["BP"].ToString();
                    workcenterName = drcheck["PrimaryWC"].ToString();
                    if (BP != plantCode)
                    {
                        return Ok("MIR " + mirno + " is from " + BP + " Plant. Cannot change WorkCenter");
                    }
                    drcheck.Close();
                    if (ack != "TRUE")
                    {
                        return Ok("Workcenter assignment is not allowed without Acknowledgment");
                    }
                }
                else
                {
                    drcheck.Close();
                    return Ok("Data is not Available in SHOPLink");
                }

                check = new SqlCommand("select count(RSNo) as TotalPO from Operations where Mirno='" + mirno + "'  ", Conn); //and POType='Primary' 
                drcheck = check.ExecuteReader();
                if (drcheck.Read())
                {
                    totalPO = int.Parse(drcheck["TotalPO"].ToString());
                }
                drcheck.Close();

                check = new SqlCommand("select count(RSNo) as TotalPO from Operations where Mirno='" + mirno + "' and MachineName is null  ", Conn);
                drcheck = check.ExecuteReader();
                if (drcheck.Read())
                {
                    PendingPO = int.Parse(drcheck["TotalPO"].ToString());
                }
                drcheck.Close();

                completedPO = totalPO - PendingPO;

                if (totalPO != completedPO)
                {
                    if (workcenterName != "")
                    {
                        if (comboBox_MachineName != workcenterName)
                        {
                            return Ok("Previously " + workcenterName + " is assigned to completed " + completedPO + " workorders. Please select " + workcenterName);
                            //return;
                        }
                        //DialogResult diaResult = MessageBox.Show(PendingPO + " WorkOrders are pending for Workcenter assignment. Previously " + workcenterName + " WorkCenter is assigned for " + completedPO + " Workorders. Do you want to assign WorkCenter for pending Workorders?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        //if (diaResult == DialogResult.Yes)
                        //{
                        //    //cmb_Group.Text = "";
                        //    //comboBox_MachineName.ResetText();
                        //    //comboBox_MachineName.Text= workcenterName.ToString();
                        //    //BeginInvoke(new Action(() => comboBox_MachineName.Text =workcenterName.ToString()));
                        //    //comboBox_MachineName.SelectedText = workcenterName;
                        //}
                        //else
                        //{
                        //    return Ok("");
                        //}
                    }
                }
                else
                {
                    //DialogResult diaResult = MessageBox.Show("Workcenter " + workcenterName + " is Already assigned to all " + totalPO + " Workorders. Do you want to change the Workcenter??", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    //if (diaResult == DialogResult.Yes)
                    //{

                    //}
                    //else
                    //{
                    //    return;
                    //}
                }
                /*check = new SqlCommand("select * from Production_Details where Mirno='" + txtMirno.Text + "'  and ActualPiece is not null ", cn);
                drcheck = check.ExecuteReader();
                if(drcheck.Read())
                {
                    string wc = drcheck["MachineName"].ToString();
                    drcheck.Close();
                    MessageBox.Show("MIR Number " + txtMirno.Text + " already assigned on WorkCenter and in proceess so WorkCenter cannot be changed");
                    //cn.Close();
                    return;
                }
                drcheck.Close();*/
                DateTime d1 = DateTime.Now;
                string scan_Date1 = d1.Month + "/" + d1.Day + "/" + d1.Year;
                xml1 = new XmlDocument();
                //progressBar2.Value = 0;
                //progressBar2.Minimum = 0;
                //progressBar2.Maximum = PendingPO;
                //progressBar2.Visible = true;
                //lblUpdationBar.Visible = true;
                #region ShopWCAssign
                #region JPR
                if (plantCode == "TM02")
                {
                    if (cmb_Group.Contains("Shop") && (cmb_Group != "Shop 1"))
                    {
                        int notch = 0, meel = 0, weld = 0, bend = 0, total = 0, i = 0; ;
                        int[] ar = new int[10];
                        string opstatus = "", PO = "";
                        SqlCommand cd = new SqlCommand("select * from Operations where Mirno= '" + mirno + "' and MachineName is null", Conn);
                        SqlDataReader dr1 = cd.ExecuteReader();
                        while (dr1.Read())
                        {
                            PO = dr1["RSNo"].ToString();
                            opstatus = dr1["OPStatus"].ToString();
                            i = 0;

                            #region All
                            if (opstatus.Contains("N") && opstatus.Contains("B") && opstatus.Contains("M"))
                            {
                                ar[0] = 10;
                                ar[1] = 40;
                                ar[2] = 50;
                                ar[3] = 70;
                                ar[4] = 90;
                                ar[5] = 100;

                                //string All = Application.StartupPath + "\\" + "All.xml";
                                xml1.Load(All);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(All);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName;
                                    xml1.Save(All);
                                }
                                nodes = xml1.SelectNodes("//IROPNO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = ar[i].ToString();
                                    xml1.Save(All);
                                    i++;
                                }
                                try
                                {
                                    sr = new StreamReader(WCAssignini);
                                    uri = sr.ReadLine();
                                    sr.Close();

                                    credentials = new NetworkCredential(UserName, Password);
                                    cc = new CredentialCache();
                                    cc.Add(new Uri(uri), "Basic", credentials);
                                    req = WebRequest.Create(uri);
                                    req.Method = "POST";
                                    req.ContentType = "text/xml";
                                    writer = new StreamWriter(req.GetRequestStream());
                                    writer.WriteLine(this.GetTextFromXMLFile(All));
                                    writer.Close();
                                    req.Credentials = cc;
                                    rsp = req.GetResponse();
                                    sr = new StreamReader(rsp.GetResponseStream());
                                    result = sr.ReadToEnd();
                                    sr.Close();
                                }
                                catch (Exception e)
                                {
                                    //progressBar2.Visible = false;
                                    //lblUpdationBar.Visible = false;
                                    return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                    //return;
                                }
                            }
                            #endregion All
                            #region N&B
                            else if (opstatus.Contains("N") && opstatus.Contains("B"))
                            {
                                ar[0] = 10;
                                ar[1] = 40;
                                ar[2] = 50;
                                ar[3] = 70;
                                ar[4] = 100;
                                //string NB = Application.StartupPath + "\\" + "NB.xml";
                                xml1.Load(NB);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(NB);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName;
                                    xml1.Save(NB);
                                }
                                nodes = xml1.SelectNodes("//IROPNO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = ar[i].ToString();
                                    xml1.Save(NB);
                                    i++;
                                }
                                try
                                {
                                    sr = new StreamReader(WCAssignini);
                                    uri = sr.ReadLine();
                                    sr.Close();

                                    credentials = new NetworkCredential(UserName, Password);
                                    cc = new CredentialCache();
                                    cc.Add(new Uri(uri), "Basic", credentials);
                                    req = WebRequest.Create(uri);
                                    req.Method = "POST";
                                    req.ContentType = "text/xml";
                                    writer = new StreamWriter(req.GetRequestStream());
                                    writer.WriteLine(this.GetTextFromXMLFile(NB));
                                    writer.Close();
                                    req.Credentials = cc;
                                    rsp = req.GetResponse();
                                    sr = new StreamReader(rsp.GetResponseStream());
                                    result = sr.ReadToEnd();
                                    sr.Close();
                                }
                                catch (Exception e)
                                {
                                    //progressBar2.Visible = false;
                                    //lblUpdationBar.Visible = false;
                                    return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                    //return;
                                }
                            }
                            #endregion N&B
                            #region N&M
                            else if (opstatus.Contains("N") && opstatus.Contains("M"))
                            {
                                ar[0] = 10;
                                ar[1] = 40;
                                ar[2] = 50;
                                ar[3] = 70;
                                ar[4] = 90;
                                //string NM = Application.StartupPath + "\\" + "NM.xml";
                                xml1.Load(NM);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(NM);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName;
                                    xml1.Save(NM);
                                }
                                nodes = xml1.SelectNodes("//IROPNO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = ar[i].ToString();
                                    xml1.Save(NM);
                                    i++;
                                }
                                try
                                {
                                    sr = new StreamReader(WCAssignini);
                                    uri = sr.ReadLine();
                                    sr.Close();

                                    credentials = new NetworkCredential(UserName, Password);
                                    cc = new CredentialCache();
                                    cc.Add(new Uri(uri), "Basic", credentials);
                                    req = WebRequest.Create(uri);
                                    req.Method = "POST";
                                    req.ContentType = "text/xml";
                                    writer = new StreamWriter(req.GetRequestStream());
                                    writer.WriteLine(this.GetTextFromXMLFile(NM));
                                    writer.Close();
                                    req.Credentials = cc;
                                    rsp = req.GetResponse();
                                    sr = new StreamReader(rsp.GetResponseStream());
                                    result = sr.ReadToEnd();
                                    sr.Close();
                                }
                                catch (Exception e)
                                {
                                    //progressBar2.Visible = false;
                                    //lblUpdationBar.Visible = false;
                                    return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                    //return;
                                }
                            }
                            #endregion N&M
                            #region B&M
                            else if (opstatus.Contains("B") && opstatus.Contains("M"))
                            {
                                ar[0] = 10;
                                ar[1] = 40;
                                ar[2] = 50;
                                ar[3] = 90;
                                ar[4] = 100;

                                //string BM = Application.StartupPath + "\\" + "BM.xml";
                                xml1.Load(BM);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(BM);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName;
                                    xml1.Save(BM);
                                }
                                nodes = xml1.SelectNodes("//IROPNO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = ar[i].ToString();
                                    xml1.Save(BM);
                                    i++;
                                }
                                try
                                {
                                    sr = new StreamReader(WCAssignini);
                                    uri = sr.ReadLine();
                                    sr.Close();

                                    credentials = new NetworkCredential(UserName, Password);
                                    cc = new CredentialCache();
                                    cc.Add(new Uri(uri), "Basic", credentials);
                                    req = WebRequest.Create(uri);
                                    req.Method = "POST";
                                    req.ContentType = "text/xml";
                                    writer = new StreamWriter(req.GetRequestStream());
                                    writer.WriteLine(this.GetTextFromXMLFile(BM));
                                    writer.Close();
                                    req.Credentials = cc;
                                    rsp = req.GetResponse();
                                    sr = new StreamReader(rsp.GetResponseStream());
                                    result = sr.ReadToEnd();
                                    sr.Close();
                                }
                                catch (Exception e)
                                {
                                    //progressBar2.Visible = false;
                                    //lblUpdationBar.Visible = false;
                                    return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                    //return;
                                }
                            }
                            #endregion N&M
                            #region N
                            else if (opstatus.Contains("N"))
                            {
                                ar[0] = 10;
                                ar[1] = 40;
                                ar[2] = 50;
                                ar[3] = 70;

                                //string Notch = Application.StartupPath + "\\" + "Notch.xml";
                                xml1.Load(Notch);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(Notch);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName;
                                    xml1.Save(Notch);
                                }
                                nodes = xml1.SelectNodes("//IROPNO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = ar[i].ToString();
                                    xml1.Save(Notch);
                                    i++;
                                }
                                try
                                {
                                    sr = new StreamReader(WCAssignini);
                                    uri = sr.ReadLine();
                                    sr.Close();

                                    credentials = new NetworkCredential(UserName, Password);
                                    cc = new CredentialCache();
                                    cc.Add(new Uri(uri), "Basic", credentials);
                                    req = WebRequest.Create(uri);
                                    req.Method = "POST";
                                    req.ContentType = "text/xml";
                                    writer = new StreamWriter(req.GetRequestStream());
                                    writer.WriteLine(this.GetTextFromXMLFile(Notch));
                                    writer.Close();
                                    req.Credentials = cc;
                                    rsp = req.GetResponse();
                                    sr = new StreamReader(rsp.GetResponseStream());
                                    result = sr.ReadToEnd();
                                    sr.Close();
                                }
                                catch (Exception e)
                                {
                                    //progressBar2.Visible = false;
                                    //lblUpdationBar.Visible = false;
                                    return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                    //return;
                                }
                            }
                            #endregion N
                            #region M
                            else if (opstatus.Contains("M"))
                            {
                                ar[0] = 10;
                                ar[1] = 40;
                                ar[2] = 50;
                                ar[3] = 90;

                                //string HM = Application.StartupPath + "\\" + "HM.xml";
                                xml1.Load(HM);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(HM);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName;
                                    xml1.Save(HM);
                                }
                                nodes = xml1.SelectNodes("//IROPNO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = ar[i].ToString();
                                    xml1.Save(HM);
                                    i++;
                                }
                                try
                                {
                                    sr = new StreamReader(WCAssignini);
                                    uri = sr.ReadLine();
                                    sr.Close();

                                    credentials = new NetworkCredential(UserName, Password);
                                    cc = new CredentialCache();
                                    cc.Add(new Uri(uri), "Basic", credentials);
                                    req = WebRequest.Create(uri);
                                    req.Method = "POST";
                                    req.ContentType = "text/xml";
                                    writer = new StreamWriter(req.GetRequestStream());
                                    writer.WriteLine(this.GetTextFromXMLFile(HM));
                                    writer.Close();
                                    req.Credentials = cc;
                                    rsp = req.GetResponse();
                                    sr = new StreamReader(rsp.GetResponseStream());
                                    result = sr.ReadToEnd();
                                    sr.Close();
                                }
                                catch (Exception ex)
                                {
                                    //progressBar2.Visible = false;
                                    //lblUpdationBar.Visible = false;
                                    return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                    //return;
                                }
                            }
                            #endregion M
                            #region B
                            else if (opstatus.Contains("B"))
                            {
                                ar[0] = 10;
                                ar[1] = 40;
                                ar[2] = 50;
                                ar[3] = 100;

                                //string Bend = Application.StartupPath + "\\" + "Bend.xml";
                                xml1.Load(Bend);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(Bend);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName;
                                    xml1.Save(Bend);
                                }
                                nodes = xml1.SelectNodes("//IROPNO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = ar[i].ToString();
                                    xml1.Save(Bend);
                                    i++;
                                }
                                try
                                {
                                    sr = new StreamReader(WCAssignini);
                                    uri = sr.ReadLine();
                                    sr.Close();

                                    credentials = new NetworkCredential(UserName, Password);
                                    cc = new CredentialCache();
                                    cc.Add(new Uri(uri), "Basic", credentials);
                                    req = WebRequest.Create(uri);
                                    req.Method = "POST";
                                    req.ContentType = "text/xml";
                                    writer = new StreamWriter(req.GetRequestStream());
                                    writer.WriteLine(this.GetTextFromXMLFile(Bend));
                                    writer.Close();
                                    req.Credentials = cc;
                                    rsp = req.GetResponse();
                                    sr = new StreamReader(rsp.GetResponseStream());
                                    result = sr.ReadToEnd();
                                    sr.Close();
                                }
                                catch (Exception ex)
                                {
                                    //progressBar2.Visible = false;
                                    //lblUpdationBar.Visible = false;
                                    return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                    //return;
                                }
                            }
                            #endregion B

                            #region OnlyFab
                            else
                            {
                                xml1.Load(WCAssign);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(WCAssign);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName;
                                    xml1.Save(WCAssign);
                                }
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
                                    sr = new StreamReader(WCAssignini);
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
                                    //progressBar2.Visible = false;
                                    //lblUpdationBar.Visible = false;
                                    return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                    //return;
                                }
                            }

                            #endregion OnlyFab
                            SqlCommand c3 = new SqlCommand("update Operations set PlanningDate='" + scan_Date1 + "',  MachineName='" + comboBox_MachineName + "',PrimaryWC='" + comboBox_MachineName + "',PlanningShift='" + cmbShift + "',RackDetails='" + txtRack + "' where RSNo='" + PO.ToString() + "'", Conn);
                            c3.ExecuteNonQuery();
                            //progressBar2.Value += 1;
                            releasedWC += 1;
                        }
                        dr1.Close();

                        #region Workcenter Change
                        if (totalPO == completedPO)
                        {
                            completedPO = 0;
                            //progressBar2.Maximum = totalPO;
                            c1 = new SqlCommand("select * from Operations where Mirno= '" + mirno + "' ", Conn);  //and POType='Primary'
                            SqlDataReader drChange = c1.ExecuteReader();
                            while (drChange.Read())
                            {
                                PO = drChange["RSNo"].ToString();
                                opstatus = drChange["OPStatus"].ToString();
                                i = 0;

                                #region All
                                if (opstatus.Contains("N") && opstatus.Contains("B") && opstatus.Contains("M"))
                                {
                                    ar[0] = 10;
                                    ar[1] = 40;
                                    ar[2] = 50;
                                    ar[3] = 70;
                                    ar[4] = 90;
                                    ar[5] = 100;

                                    //string All = Application.StartupPath + "\\" + "All.xml";
                                    xml1.Load(All);
                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(All);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(All);
                                    }
                                    nodes = xml1.SelectNodes("//IROPNO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = ar[i].ToString();
                                        xml1.Save(All);
                                        i++;
                                    }
                                    try
                                    {
                                        sr = new StreamReader(WCAssignini);
                                        uri = sr.ReadLine();
                                        sr.Close();

                                        credentials = new NetworkCredential(UserName, Password);
                                        cc = new CredentialCache();
                                        cc.Add(new Uri(uri), "Basic", credentials);
                                        req = WebRequest.Create(uri);
                                        req.Method = "POST";
                                        req.ContentType = "text/xml";
                                        writer = new StreamWriter(req.GetRequestStream());
                                        writer.WriteLine(this.GetTextFromXMLFile(All));
                                        writer.Close();
                                        req.Credentials = cc;
                                        rsp = req.GetResponse();
                                        sr = new StreamReader(rsp.GetResponseStream());
                                        result = sr.ReadToEnd();
                                        sr.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                        //return;
                                    }
                                }
                                #endregion All
                                #region N&B
                                else if (opstatus.Contains("N") && opstatus.Contains("B"))
                                {
                                    ar[0] = 10;
                                    ar[1] = 40;
                                    ar[2] = 50;
                                    ar[3] = 70;
                                    ar[4] = 100;

                                    //string NB = Application.StartupPath + "\\" + "NB.xml";
                                    xml1.Load(NB);
                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(NB);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(NB);
                                    }
                                    nodes = xml1.SelectNodes("//IROPNO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = ar[i].ToString();
                                        xml1.Save(NB);
                                        i++;
                                    }
                                    try
                                    {
                                        sr = new StreamReader(WCAssignini);
                                        uri = sr.ReadLine();
                                        sr.Close();

                                        credentials = new NetworkCredential(UserName, Password);
                                        cc = new CredentialCache();
                                        cc.Add(new Uri(uri), "Basic", credentials);
                                        req = WebRequest.Create(uri);
                                        req.Method = "POST";
                                        req.ContentType = "text/xml";
                                        writer = new StreamWriter(req.GetRequestStream());
                                        writer.WriteLine(this.GetTextFromXMLFile(NB));
                                        writer.Close();
                                        req.Credentials = cc;
                                        rsp = req.GetResponse();
                                        sr = new StreamReader(rsp.GetResponseStream());
                                        result = sr.ReadToEnd();
                                        sr.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                        //return;
                                    }
                                }
                                #endregion N&B
                                #region N&M
                                else if (opstatus.Contains("N") && opstatus.Contains("M"))
                                {
                                    ar[0] = 10;
                                    ar[1] = 40;
                                    ar[2] = 50;
                                    ar[3] = 70;
                                    ar[4] = 90;

                                    //string NM = Application.StartupPath + "\\" + "NM.xml";
                                    xml1.Load(NM);
                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(NM);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(NM);
                                    }
                                    nodes = xml1.SelectNodes("//IROPNO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = ar[i].ToString();
                                        xml1.Save(NM);
                                        i++;
                                    }
                                    try
                                    {
                                        sr = new StreamReader(WCAssignini);
                                        uri = sr.ReadLine();
                                        sr.Close();

                                        credentials = new NetworkCredential(UserName, Password);
                                        cc = new CredentialCache();
                                        cc.Add(new Uri(uri), "Basic", credentials);
                                        req = WebRequest.Create(uri);
                                        req.Method = "POST";
                                        req.ContentType = "text/xml";
                                        writer = new StreamWriter(req.GetRequestStream());
                                        writer.WriteLine(this.GetTextFromXMLFile(NM));
                                        writer.Close();
                                        req.Credentials = cc;
                                        rsp = req.GetResponse();
                                        sr = new StreamReader(rsp.GetResponseStream());
                                        result = sr.ReadToEnd();
                                        sr.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                        //return;
                                    }
                                }
                                #endregion N&M
                                #region B&M
                                else if (opstatus.Contains("B") && opstatus.Contains("M"))
                                {
                                    ar[0] = 10;
                                    ar[1] = 40;
                                    ar[2] = 50;
                                    ar[3] = 90;
                                    ar[4] = 100;

                                    //string BM = Application.StartupPath + "\\" + "BM.xml";
                                    xml1.Load(BM);
                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(BM);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(BM);
                                    }
                                    nodes = xml1.SelectNodes("//IROPNO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = ar[i].ToString();
                                        xml1.Save(BM);
                                        i++;
                                    }
                                    try
                                    {
                                        sr = new StreamReader(WCAssignini);
                                        uri = sr.ReadLine();
                                        sr.Close();

                                        credentials = new NetworkCredential(UserName, Password);
                                        cc = new CredentialCache();
                                        cc.Add(new Uri(uri), "Basic", credentials);
                                        req = WebRequest.Create(uri);
                                        req.Method = "POST";
                                        req.ContentType = "text/xml";
                                        writer = new StreamWriter(req.GetRequestStream());
                                        writer.WriteLine(this.GetTextFromXMLFile(BM));
                                        writer.Close();
                                        req.Credentials = cc;
                                        rsp = req.GetResponse();
                                        sr = new StreamReader(rsp.GetResponseStream());
                                        result = sr.ReadToEnd();
                                        sr.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                        //return;
                                    }
                                }
                                #endregion N&M
                                #region N
                                else if (opstatus.Contains("N"))
                                {
                                    ar[0] = 10;
                                    ar[1] = 40;
                                    ar[2] = 50;
                                    ar[3] = 70;

                                    //string Notch = Application.StartupPath + "\\" + "Notch.xml";
                                    xml1.Load(Notch);
                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(Notch);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(Notch);
                                    }
                                    nodes = xml1.SelectNodes("//IROPNO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = ar[i].ToString();
                                        xml1.Save(Notch);
                                        i++;
                                    }
                                    try
                                    {
                                        sr = new StreamReader(WCAssignini);
                                        uri = sr.ReadLine();
                                        sr.Close();

                                        credentials = new NetworkCredential(UserName, Password);
                                        cc = new CredentialCache();
                                        cc.Add(new Uri(uri), "Basic", credentials);
                                        req = WebRequest.Create(uri);
                                        req.Method = "POST";
                                        req.ContentType = "text/xml";
                                        writer = new StreamWriter(req.GetRequestStream());
                                        writer.WriteLine(this.GetTextFromXMLFile(Notch));
                                        writer.Close();
                                        req.Credentials = cc;
                                        rsp = req.GetResponse();
                                        sr = new StreamReader(rsp.GetResponseStream());
                                        result = sr.ReadToEnd();
                                        sr.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                        //return;
                                    }
                                }
                                #endregion N
                                #region M
                                else if (opstatus.Contains("M"))
                                {
                                    ar[0] = 10;
                                    ar[1] = 40;
                                    ar[2] = 50;
                                    ar[3] = 90;

                                    //string HM = Application.StartupPath + "\\" + "HM.xml";
                                    xml1.Load(HM);
                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(HM);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(HM);
                                    }
                                    nodes = xml1.SelectNodes("//IROPNO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = ar[i].ToString();
                                        xml1.Save(HM);
                                        i++;
                                    }
                                    try
                                    {
                                        sr = new StreamReader(WCAssignini);
                                        uri = sr.ReadLine();
                                        sr.Close();

                                        credentials = new NetworkCredential(UserName, Password);
                                        cc = new CredentialCache();
                                        cc.Add(new Uri(uri), "Basic", credentials);
                                        req = WebRequest.Create(uri);
                                        req.Method = "POST";
                                        req.ContentType = "text/xml";
                                        writer = new StreamWriter(req.GetRequestStream());
                                        writer.WriteLine(this.GetTextFromXMLFile(HM));
                                        writer.Close();
                                        req.Credentials = cc;
                                        rsp = req.GetResponse();
                                        sr = new StreamReader(rsp.GetResponseStream());
                                        result = sr.ReadToEnd();
                                        sr.Close();
                                    }
                                    catch (Exception ex)
                                    {
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                        //return;
                                    }
                                }
                                #endregion M
                                #region B
                                else if (opstatus.Contains("B"))
                                {
                                    ar[0] = 10;
                                    ar[1] = 40;
                                    ar[2] = 50;
                                    ar[3] = 100;

                                    //string Bend = Application.StartupPath + "\\" + "Bend.xml";
                                    xml1.Load(Bend);
                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(Bend);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(Bend);
                                    }
                                    nodes = xml1.SelectNodes("//IROPNO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = ar[i].ToString();
                                        xml1.Save(Bend);
                                        i++;
                                    }
                                    try
                                    {
                                        sr = new StreamReader(WCAssignini);
                                        uri = sr.ReadLine();
                                        sr.Close();

                                        credentials = new NetworkCredential(UserName, Password);
                                        cc = new CredentialCache();
                                        cc.Add(new Uri(uri), "Basic", credentials);
                                        req = WebRequest.Create(uri);
                                        req.Method = "POST";
                                        req.ContentType = "text/xml";
                                        writer = new StreamWriter(req.GetRequestStream());
                                        writer.WriteLine(this.GetTextFromXMLFile(Bend));
                                        writer.Close();
                                        req.Credentials = cc;
                                        rsp = req.GetResponse();
                                        sr = new StreamReader(rsp.GetResponseStream());
                                        result = sr.ReadToEnd();
                                        sr.Close();
                                    }
                                    catch (Exception ex)
                                    {
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                        //return;
                                    }
                                }
                                #endregion B

                                #region OnlyFab
                                else
                                {
                                    xml1.Load(WCAssign);

                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(WCAssign);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(WCAssign);
                                    }
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
                                        sr = new StreamReader(WCAssignini);
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
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!"+ ex.Message);
                                        //return;
                                    }
                                }
                                #endregion OnlyFab
                                //progressBar2.Value += 1;
                                releasedWC += 1;
                            }
                            drChange.Close();
                            c1 = new SqlCommand("insert into Operations(RSNo,FGItem,Operation,Pices,ActualPiece,Wheight,Setups,SctDinemtion,PlanningDate,PlanningShift,Operator,Status,Mirno,LotCode,BP,CountStatus,TotalWt,Tot_OPS,Length,MachineName,JDDate,CompletedQty,OPStatus,RunTime,Flag_Ack,PrimaryWC,POType) select RSNo,FGItem,Operation,Pices,ActualPiece,Wheight,Setups,SctDinemtion,'" + scan_Date1 + "','" + cmbShift + "',Operator,Status,Mirno,LotCode,BP,CountStatus,TotalWt,Tot_OPS,Length,'" + comboBox_MachineName + "',JDDate,CompletedQty,OPStatus,RunTime,Flag_Ack,PrimaryWC,'Duplicate' from Operations where Mirno='" + mirno + "' and MachineName=(select TOP 1 MachineName from Operations where Mirno='" + mirno + "') and PlanningDate=(select TOP 1 PlanningDate from Operations where Mirno='" + mirno + "') and PlanningShift=(select TOP 1 PlanningShift from Operations where Mirno='" + mirno + "')", Conn);
                            c1.ExecuteNonQuery();
                        }
                        #endregion WorkCenter Change

                        //***Code moved in common function : bindDataGridAfterAllocate 
                        //string CommandText = "SELECT RSNo as [PO.NO.],FGItem as [Item No],Mirno as MIRNO,MachineName as WorkCenter,SctDinemtion as SECTION,OPStatus,LotCode as [Billable Lot],Pices as QTY,Length,Wheight as Weight,RackDetails,PlanningShift,SAPPulledDate as PulledDate,JDDate as ReleaseDate,PlanningDate,Operation,Tot_OPS,RunTime,Status,TotalWt FROM Operations where Mirno='" + mirno + "' and BP='" + plantCode + "' and Flag_Fab is null order by [index] desc ";
                        //SqlCommand myCommand = new SqlCommand(CommandText, Conn);
                        //SqlDataAdapter myAdapter = new SqlDataAdapter();
                        //myAdapter.SelectCommand = myCommand;
                        //DataSet myDataSet = new DataSet();
                        //myAdapter.Fill(myDataSet);
                        //dataGridView1.DataSource = myDataSet.Tables[0];

                        //Added on 22.9.2018 for showing Runtime and weight...

                        //***Code moved in common function : getCalculationsAfterAllocate
                        //c1 = new SqlCommand("select sum(TotalWt)/1000 as TotalWheight,count(RSNo) as RSno,sum(Tot_OPS) as TotalOpns,sum(RunTime) as RunTime from Operations where Mirno='" + txtMirno.Text + "' and BP='" + lblPlantCode.Text + "' and Flag_Fab is null and POType='Primary'", cn);
                        //SqlDataReader dr = c1.ExecuteReader();
                        //if (dr.Read())
                        //{
                        //    if (dr["TotalWheight"].ToString() != "")
                        //        wheight = Math.Round(double.Parse(dr["TotalWheight"].ToString()), 3);

                        //    txtRs.Text = dr["RSno"].ToString();
                        //    if (dr["TotalOpns"].ToString() != "")
                        //        Operations = double.Parse(dr["TotalOpns"].ToString());
                        //    if (dr["RunTime"].ToString() != "")
                        //        Runtime1 = Math.Round(double.Parse(dr["RunTime"].ToString()), 2);
                        //}
                        //txtwheight.Text = wheight.ToString();
                        //txtOprns.Text = Operations.ToString();
                        //txtRunTime.Text = Runtime1.ToString();
                        //dr.Close();

                        int totalcompleted = completedPO + releasedWC;
                        //MessageBox.Show("Workcenter " + comboBox_MachineName.Text + " is assigned to total " + totalcompleted + "Production Orders. out of" + totalPO + " MIR No: " + txtAckMirno.Text);
                        return Ok("Workcenter " + comboBox_MachineName + " is assigned to total " + totalcompleted + " Production Orders out of " + totalPO + " MIR No: " + mirno);

                        //progressBar2.Visible = false;
                        //lblUpdationBar.Visible = false;
                        //return;

                    }

                }

                #endregion JPR and JBP

                #region JBP
                if (plantCode == "TM01")
                {
                    if ((cmb_Group != "Shop 1") && (cmb_Group != "Shop 2"))
                    {
                        // cn.Open();
                        // Conn.Open();
                        int notch = 0, meel = 0, weld = 0, bend = 0, total = 0, i = 0; ;
                        int[] ar = new int[10];
                        string opstatus = "", PO = "";
                        SqlCommand cd = new SqlCommand("select * from Operations where Mirno= '" + mirno + "' and MachineName is null", Conn);
                        SqlDataReader dr1 = cd.ExecuteReader();
                        while (dr1.Read())
                        {
                            PO = dr1["RSNo"].ToString();
                            opstatus = dr1["OPStatus"].ToString();
                            i = 0;

                            #region All
                            if (opstatus.Contains("N") && opstatus.Contains("B") && opstatus.Contains("M"))
                            {
                                ar[0] = 10;
                                ar[1] = 40;
                                ar[2] = 50;
                                ar[3] = 70;
                                ar[4] = 90;
                                ar[5] = 100;

                                //string All = Application.StartupPath + "\\" + "All.xml";
                                xml1.Load(All);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(All);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName.Trim();
                                    xml1.Save(All);
                                }
                                nodes = xml1.SelectNodes("//IROPNO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = ar[i].ToString();
                                    xml1.Save(All);
                                    i++;
                                }
                                try
                                {
                                    sr = new StreamReader(WCAssignini);
                                    uri = sr.ReadLine();
                                    sr.Close();

                                    credentials = new NetworkCredential(UserName, Password);
                                    cc = new CredentialCache();
                                    cc.Add(new Uri(uri), "Basic", credentials);
                                    req = WebRequest.Create(uri);
                                    req.Method = "POST";
                                    req.ContentType = "text/xml";
                                    writer = new StreamWriter(req.GetRequestStream());
                                    writer.WriteLine(this.GetTextFromXMLFile(All));
                                    writer.Close();
                                    req.Credentials = cc;
                                    rsp = req.GetResponse();
                                    sr = new StreamReader(rsp.GetResponseStream());
                                    result = sr.ReadToEnd();
                                    sr.Close();
                                }
                                catch (Exception ex)
                                {
                                    return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                    //return;
                                }
                            }
                            #endregion All
                            #region N&B
                            else if (opstatus.Contains("N") && opstatus.Contains("B"))
                            {
                                ar[0] = 10;
                                ar[1] = 40;
                                ar[2] = 50;
                                ar[3] = 70;
                                ar[4] = 100;

                                //string NB = Application.StartupPath + "\\" + "NB.xml";
                                xml1.Load(NB);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(NB);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName.Trim();
                                    xml1.Save(NB);
                                }
                                nodes = xml1.SelectNodes("//IROPNO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = ar[i].ToString();
                                    xml1.Save(NB);
                                    i++;
                                }
                                try
                                {
                                    sr = new StreamReader(WCAssignini);
                                    uri = sr.ReadLine();
                                    sr.Close();

                                    credentials = new NetworkCredential(UserName, Password);
                                    cc = new CredentialCache();
                                    cc.Add(new Uri(uri), "Basic", credentials);
                                    req = WebRequest.Create(uri);
                                    req.Method = "POST";
                                    req.ContentType = "text/xml";
                                    writer = new StreamWriter(req.GetRequestStream());
                                    writer.WriteLine(this.GetTextFromXMLFile(NB));
                                    writer.Close();
                                    req.Credentials = cc;
                                    rsp = req.GetResponse();
                                    sr = new StreamReader(rsp.GetResponseStream());
                                    result = sr.ReadToEnd();
                                    sr.Close();
                                }
                                catch (Exception ex)
                                {
                                    return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                    //return;
                                }
                            }
                            #endregion N&B
                            #region N&M
                            else if (opstatus.Contains("N") && opstatus.Contains("M"))
                            {
                                ar[0] = 10;
                                ar[1] = 40;
                                ar[2] = 50;
                                ar[3] = 70;
                                ar[4] = 90;

                                //string NM = Application.StartupPath + "\\" + "NM.xml";
                                xml1.Load(NM);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(NM);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName.Trim();
                                    xml1.Save(NM);
                                }
                                nodes = xml1.SelectNodes("//IROPNO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = ar[i].ToString();
                                    xml1.Save(NM);
                                    i++;
                                }
                                try
                                {
                                    sr = new StreamReader(WCAssignini);
                                    uri = sr.ReadLine();
                                    sr.Close();

                                    credentials = new NetworkCredential(UserName, Password);
                                    cc = new CredentialCache();
                                    cc.Add(new Uri(uri), "Basic", credentials);
                                    req = WebRequest.Create(uri);
                                    req.Method = "POST";
                                    req.ContentType = "text/xml";
                                    writer = new StreamWriter(req.GetRequestStream());
                                    writer.WriteLine(this.GetTextFromXMLFile(NM));
                                    writer.Close();
                                    req.Credentials = cc;
                                    rsp = req.GetResponse();
                                    sr = new StreamReader(rsp.GetResponseStream());
                                    result = sr.ReadToEnd();
                                    sr.Close();
                                }
                                catch (Exception ex)
                                {
                                    return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                    //return;
                                }
                            }
                            #endregion N&M
                            #region B&M
                            else if (opstatus.Contains("B") && opstatus.Contains("M"))
                            {
                                ar[0] = 10;
                                ar[1] = 40;
                                ar[2] = 50;
                                ar[3] = 90;
                                ar[4] = 100;

                                //string BM = Application.StartupPath + "\\" + "BM.xml";
                                xml1.Load(BM);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(BM);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName.Trim();
                                    xml1.Save(BM);
                                }
                                nodes = xml1.SelectNodes("//IROPNO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = ar[i].ToString();
                                    xml1.Save(BM);
                                    i++;
                                }
                                try
                                {
                                    sr = new StreamReader(WCAssignini);
                                    uri = sr.ReadLine();
                                    sr.Close();

                                    credentials = new NetworkCredential(UserName, Password);
                                    cc = new CredentialCache();
                                    cc.Add(new Uri(uri), "Basic", credentials);
                                    req = WebRequest.Create(uri);
                                    req.Method = "POST";
                                    req.ContentType = "text/xml";
                                    writer = new StreamWriter(req.GetRequestStream());
                                    writer.WriteLine(this.GetTextFromXMLFile(BM));
                                    writer.Close();
                                    req.Credentials = cc;
                                    rsp = req.GetResponse();
                                    sr = new StreamReader(rsp.GetResponseStream());
                                    result = sr.ReadToEnd();
                                    sr.Close();
                                }
                                catch (Exception ex)
                                {
                                    return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                    //return;
                                }
                            }
                            #endregion N&M
                            #region N
                            else if (opstatus.Contains("N"))
                            {
                                ar[0] = 10;
                                ar[1] = 40;
                                ar[2] = 50;
                                ar[3] = 70;

                                //string Notch = Application.StartupPath + "\\" + "Notch.xml";
                                xml1.Load(Notch);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(Notch);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName.Trim();
                                    xml1.Save(Notch);
                                }
                                nodes = xml1.SelectNodes("//IROPNO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = ar[i].ToString();
                                    xml1.Save(Notch);
                                    i++;
                                }
                                try
                                {
                                    sr = new StreamReader(WCAssignini);
                                    uri = sr.ReadLine();
                                    sr.Close();

                                    credentials = new NetworkCredential(UserName, Password);
                                    cc = new CredentialCache();
                                    cc.Add(new Uri(uri), "Basic", credentials);
                                    req = WebRequest.Create(uri);
                                    req.Method = "POST";
                                    req.ContentType = "text/xml";
                                    writer = new StreamWriter(req.GetRequestStream());
                                    writer.WriteLine(this.GetTextFromXMLFile(Notch));
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
                                    //return;
                                }
                            }
                            #endregion N
                            #region M
                            else if (opstatus.Contains("M"))
                            {
                                ar[0] = 10;
                                ar[1] = 40;
                                ar[2] = 50;
                                ar[3] = 90;

                                //string HM = Application.StartupPath + "\\" + "HM.xml";
                                xml1.Load(HM);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(HM);
                                }

                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName.Trim();
                                    xml1.Save(HM);
                                }
                                nodes = xml1.SelectNodes("//IROPNO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = ar[i].ToString();
                                    xml1.Save(HM);
                                    i++;
                                }
                                try
                                {
                                    sr = new StreamReader(WCAssignini);
                                    uri = sr.ReadLine();
                                    sr.Close();

                                    credentials = new NetworkCredential(UserName, Password);
                                    cc = new CredentialCache();
                                    cc.Add(new Uri(uri), "Basic", credentials);
                                    req = WebRequest.Create(uri);
                                    req.Method = "POST";
                                    req.ContentType = "text/xml";
                                    writer = new StreamWriter(req.GetRequestStream());
                                    writer.WriteLine(this.GetTextFromXMLFile(HM));
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
                                    //return;
                                }
                            }
                            #endregion M
                            #region B
                            else if (opstatus.Contains("B"))
                            {
                                ar[0] = 10;
                                ar[1] = 40;
                                ar[2] = 50;
                                ar[3] = 100;

                                //string Bend = Application.StartupPath + "\\" + "Bend.xml";
                                xml1.Load(Bend);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(Bend);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName.Trim();
                                    xml1.Save(Bend);
                                }
                                nodes = xml1.SelectNodes("//IROPNO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = ar[i].ToString();
                                    xml1.Save(Bend);
                                    i++;
                                }
                                try
                                {
                                    sr = new StreamReader(WCAssignini);
                                    uri = sr.ReadLine();
                                    sr.Close();

                                    credentials = new NetworkCredential(UserName, Password);
                                    cc = new CredentialCache();
                                    cc.Add(new Uri(uri), "Basic", credentials);

                                    req = WebRequest.Create(uri);
                                    req.Method = "POST";
                                    req.ContentType = "text/xml";
                                    writer = new StreamWriter(req.GetRequestStream());
                                    writer.WriteLine(this.GetTextFromXMLFile(Bend));
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
                                    //return;
                                }
                            }
                            #endregion B

                            #region OnlyFab
                            else
                            {
                                xml1.Load(WCAssign);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(WCAssign);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName.Trim();
                                    xml1.Save(WCAssign);
                                }
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
                                    sr = new StreamReader(WCAssignini);
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
                                catch (Exception e)
                                {
                                    return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                    //return;
                                }
                            }
                            #endregion OnlyFab

                            SqlCommand c3 = new SqlCommand("update Operations set PlanningDate='" + scan_Date1 + "',  MachineName='" + comboBox_MachineName + "', PrimaryWC='" + comboBox_MachineName + "',PlanningShift='" + cmbShift + "',RackDetails='" + txtRack + "' where RSNo='" + PO.Trim() + "'", Conn);
                            c3.ExecuteNonQuery();
                            //progressBar2.Value += 1;
                            releasedWC += 1;
                        }
                        dr1.Close();

                        #region Workcenter Change
                        if (totalPO == completedPO)
                        {
                            completedPO = 0;
                            //progressBar2.Maximum = totalPO;
                            c1 = new SqlCommand("select * from Operations where Mirno= '" + mirno + "' ", Conn); //and POType='Primary'
                            SqlDataReader drChange = c1.ExecuteReader();
                            while (drChange.Read())
                            {
                                PO = drChange["RSNo"].ToString();
                                opstatus = drChange["OPStatus"].ToString();
                                i = 0;

                                #region All
                                if (opstatus.Contains("N") && opstatus.Contains("B") && opstatus.Contains("M"))
                                {
                                    ar[0] = 10;
                                    ar[1] = 40;
                                    ar[2] = 50;
                                    ar[3] = 70;
                                    ar[4] = 90;
                                    ar[5] = 100;

                                    //string All = Application.StartupPath + "\\" + "All.xml";
                                    xml1.Load(All);
                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(All);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(All);
                                    }
                                    nodes = xml1.SelectNodes("//IROPNO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = ar[i].ToString();
                                        xml1.Save(All);
                                        i++;

                                    }
                                    try
                                    {
                                        sr = new StreamReader(WCAssignini);
                                        uri = sr.ReadLine();
                                        sr.Close();

                                        credentials = new NetworkCredential(UserName, Password);
                                        cc = new CredentialCache();
                                        cc.Add(new Uri(uri), "Basic", credentials);
                                        req = WebRequest.Create(uri);
                                        req.Method = "POST";
                                        req.ContentType = "text/xml";
                                        writer = new StreamWriter(req.GetRequestStream());
                                        writer.WriteLine(this.GetTextFromXMLFile(All));
                                        writer.Close();
                                        req.Credentials = cc;
                                        rsp = req.GetResponse();
                                        sr = new StreamReader(rsp.GetResponseStream());
                                        result = sr.ReadToEnd();
                                        sr.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                        //return;
                                    }
                                }
                                #endregion All
                                #region N&B
                                else if (opstatus.Contains("N") && opstatus.Contains("B"))
                                {
                                    ar[0] = 10;
                                    ar[1] = 40;
                                    ar[2] = 50;
                                    ar[3] = 70;
                                    ar[4] = 100;

                                    //string NB = Application.StartupPath + "\\" + "NB.xml";
                                    xml1.Load(NB);
                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(NB);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(NB);
                                    }
                                    nodes = xml1.SelectNodes("//IROPNO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = ar[i].ToString();
                                        xml1.Save(NB);
                                        i++;
                                    }
                                    try
                                    {
                                        sr = new StreamReader(WCAssignini);
                                        uri = sr.ReadLine();
                                        sr.Close();

                                        credentials = new NetworkCredential(UserName, Password);
                                        cc = new CredentialCache();
                                        cc.Add(new Uri(uri), "Basic", credentials);
                                        req = WebRequest.Create(uri);
                                        req.Method = "POST";
                                        req.ContentType = "text/xml";
                                        writer = new StreamWriter(req.GetRequestStream());
                                        writer.WriteLine(this.GetTextFromXMLFile(NB));
                                        writer.Close();
                                        req.Credentials = cc;
                                        rsp = req.GetResponse();
                                        sr = new StreamReader(rsp.GetResponseStream());
                                        result = sr.ReadToEnd();
                                        sr.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                        //return;
                                    }
                                }
                                #endregion N&B
                                #region N&M
                                else if (opstatus.Contains("N") && opstatus.Contains("M"))
                                {
                                    ar[0] = 10;
                                    ar[1] = 40;
                                    ar[2] = 50;
                                    ar[3] = 70;
                                    ar[4] = 90;

                                    //string NM = Application.StartupPath + "\\" + "NM.xml";
                                    xml1.Load(NM);
                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(NM);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(NM);
                                    }
                                    nodes = xml1.SelectNodes("//IROPNO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = ar[i].ToString();
                                        xml1.Save(NM);
                                        i++;
                                    }
                                    try
                                    {
                                        sr = new StreamReader(WCAssignini);
                                        uri = sr.ReadLine();
                                        sr.Close();

                                        credentials = new NetworkCredential(UserName, Password);
                                        cc = new CredentialCache();
                                        cc.Add(new Uri(uri), "Basic", credentials);
                                        req = WebRequest.Create(uri);
                                        req.Method = "POST";
                                        req.ContentType = "text/xml";
                                        writer = new StreamWriter(req.GetRequestStream());
                                        writer.WriteLine(this.GetTextFromXMLFile(NM));
                                        writer.Close();
                                        req.Credentials = cc;
                                        rsp = req.GetResponse();
                                        sr = new StreamReader(rsp.GetResponseStream());
                                        result = sr.ReadToEnd();
                                        sr.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                        //return;
                                    }
                                }
                                #endregion N&M
                                #region B&M
                                else if (opstatus.Contains("B") && opstatus.Contains("M"))
                                {
                                    ar[0] = 10;
                                    ar[1] = 40;
                                    ar[2] = 50;
                                    ar[3] = 90;
                                    ar[4] = 100;

                                    //string BM = Application.StartupPath + "\\" + "BM.xml";
                                    xml1.Load(BM);
                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(BM);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(BM);
                                    }
                                    nodes = xml1.SelectNodes("//IROPNO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = ar[i].ToString();
                                        xml1.Save(BM);
                                        i++;
                                    }
                                    try
                                    {
                                        sr = new StreamReader(WCAssignini);
                                        uri = sr.ReadLine();
                                        sr.Close();

                                        credentials = new NetworkCredential(UserName, Password);
                                        cc = new CredentialCache();
                                        cc.Add(new Uri(uri), "Basic", credentials);
                                        req = WebRequest.Create(uri);
                                        req.Method = "POST";
                                        req.ContentType = "text/xml";
                                        writer = new StreamWriter(req.GetRequestStream());
                                        writer.WriteLine(this.GetTextFromXMLFile(BM));
                                        writer.Close();
                                        req.Credentials = cc;
                                        rsp = req.GetResponse();
                                        sr = new StreamReader(rsp.GetResponseStream());
                                        result = sr.ReadToEnd();
                                        sr.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                        //return;
                                    }
                                }
                                #endregion N&M
                                #region N
                                else if (opstatus.Contains("N"))
                                {
                                    ar[0] = 10;
                                    ar[1] = 40;
                                    ar[2] = 50;
                                    ar[3] = 70;

                                    //string Notch = Application.StartupPath + "\\" + "Notch.xml";
                                    xml1.Load(Notch);
                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(Notch);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(Notch);
                                    }
                                    nodes = xml1.SelectNodes("//IROPNO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = ar[i].ToString();
                                        xml1.Save(Notch);
                                        i++;
                                    }
                                    try
                                    {
                                        sr = new StreamReader(WCAssignini);
                                        uri = sr.ReadLine();
                                        sr.Close();

                                        credentials = new NetworkCredential(UserName, Password);
                                        cc = new CredentialCache();
                                        cc.Add(new Uri(uri), "Basic", credentials);
                                        req = WebRequest.Create(uri);
                                        req.Method = "POST";
                                        req.ContentType = "text/xml";
                                        writer = new StreamWriter(req.GetRequestStream());
                                        writer.WriteLine(this.GetTextFromXMLFile(Notch));
                                        writer.Close();
                                        req.Credentials = cc;
                                        rsp = req.GetResponse();
                                        sr = new StreamReader(rsp.GetResponseStream());
                                        result = sr.ReadToEnd();
                                        sr.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                        //return;
                                    }
                                }
                                #endregion N
                                #region M
                                else if (opstatus.Contains("M"))
                                {
                                    ar[0] = 10;
                                    ar[1] = 40;
                                    ar[2] = 50;
                                    ar[3] = 90;

                                    //string HM = Application.StartupPath + "\\" + "HM.xml";
                                    xml1.Load(HM);
                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(HM);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(HM);
                                    }
                                    nodes = xml1.SelectNodes("//IROPNO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = ar[i].ToString();
                                        xml1.Save(HM);
                                        i++;
                                    }
                                    try
                                    {
                                        sr = new StreamReader(WCAssignini);
                                        uri = sr.ReadLine();
                                        sr.Close();

                                        credentials = new NetworkCredential(UserName, Password);
                                        cc = new CredentialCache();
                                        cc.Add(new Uri(uri), "Basic", credentials);
                                        req = WebRequest.Create(uri);
                                        req.Method = "POST";
                                        req.ContentType = "text/xml";
                                        writer = new StreamWriter(req.GetRequestStream());
                                        writer.WriteLine(this.GetTextFromXMLFile(HM));
                                        writer.Close();
                                        req.Credentials = cc;
                                        rsp = req.GetResponse();
                                        sr = new StreamReader(rsp.GetResponseStream());
                                        result = sr.ReadToEnd();
                                        sr.Close();
                                    }
                                    catch (Exception ex)
                                    {
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                        //return;
                                    }
                                }
                                #endregion M
                                #region B
                                else if (opstatus.Contains("B"))
                                {
                                    ar[0] = 10;
                                    ar[1] = 40;
                                    ar[2] = 50;
                                    ar[3] = 100;

                                    //string Bend = Application.StartupPath + "\\" + "Bend.xml";
                                    xml1.Load(Bend);
                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(Bend);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(Bend);
                                    }
                                    nodes = xml1.SelectNodes("//IROPNO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = ar[i].ToString();
                                        xml1.Save(Bend);
                                        i++;
                                    }
                                    try
                                    {
                                        sr = new StreamReader(WCAssignini);
                                        uri = sr.ReadLine();
                                        sr.Close();

                                        credentials = new NetworkCredential(UserName, Password);
                                        cc = new CredentialCache();
                                        cc.Add(new Uri(uri), "Basic", credentials);
                                        req = WebRequest.Create(uri);
                                        req.Method = "POST";
                                        req.ContentType = "text/xml";
                                        writer = new StreamWriter(req.GetRequestStream());
                                        writer.WriteLine(this.GetTextFromXMLFile(Bend));
                                        writer.Close();
                                        req.Credentials = cc;
                                        rsp = req.GetResponse();
                                        sr = new StreamReader(rsp.GetResponseStream());
                                        result = sr.ReadToEnd();
                                        sr.Close();
                                    }
                                    catch (Exception ex)
                                    {
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                        //return;
                                    }
                                }
                                #endregion B

                                #region OnlyFab
                                else
                                {
                                    xml1.Load(WCAssign);
                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(WCAssign);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(WCAssign);
                                    }
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
                                        sr = new StreamReader(WCAssignini);
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
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                        //return;
                                    }
                                }
                                #endregion OnlyFab
                                //progressBar2.Value += 1;
                                releasedWC += 1;
                            }
                            drChange.Close();

                            c1 = new SqlCommand("insert into Operations(RSNo,FGItem,Operation,Pices,ActualPiece,Wheight,Setups,SctDinemtion,PlanningDate,PlanningShift,Operator,Status,Mirno,LotCode,BP,CountStatus,TotalWt,Tot_OPS,Length,MachineName,JDDate,CompletedQty,OPStatus,RunTime,Flag_Ack,PrimaryWC,POType) select RSNo,FGItem,Operation,Pices,ActualPiece,Wheight,Setups,SctDinemtion,'" + scan_Date1 + "','" + cmbShift + "',Operator,Status,Mirno,LotCode,BP,CountStatus,TotalWt,Tot_OPS,Length,'" + comboBox_MachineName + "',JDDate,CompletedQty,OPStatus,RunTime,Flag_Ack,PrimaryWC,'Duplicate' from Operations where Mirno='" + mirno + "' and MachineName=(select TOP 1 MachineName from Operations where Mirno='" + mirno + "') and PlanningDate=(select TOP 1 PlanningDate from Operations where Mirno='" + mirno + "') and PlanningShift=(select TOP 1 PlanningShift from Operations where Mirno='" + mirno + "')", Conn);
                            c1.ExecuteNonQuery();
                        }
                        #endregion WorkCenter Change

                        //***Code moved in common function : bindDataGridAfterAllocate 

                        //string CommandText = "SELECT RSNo as [PO.NO.],FGItem as [Item No],Mirno as MIRNO,MachineName as WorkCenter,SctDinemtion as SECTION,OPStatus,LotCode as [Billable Lot],Pices as QTY,Length,Wheight as Weight,RackDetails,PlanningShift,SAPPulledDate as PulledDate,JDDate as ReleaseDate,PlanningDate,Operation,Tot_OPS,RunTime,Status,TotalWt FROM Operations where Mirno='" + mirno + "' and BP='" + plantCode + "' and Flag_Fab is null order by [index] desc ";
                        //SqlCommand myCommand = new SqlCommand(CommandText, Conn);
                        //SqlDataAdapter myAdapter = new SqlDataAdapter();
                        //myAdapter.SelectCommand = myCommand;
                        //DataSet myDataSet = new DataSet();
                        //myAdapter.Fill(myDataSet);
                        //dataGridView1.DataSource = myDataSet.Tables[0];

                        int totalcompleted = completedPO + releasedWC;
                        //MessageBox.Show("Workcenter " + comboBox_MachineName.Text + " is assigned to " + totalcompleted + " Production Orders out of" + totalPO + ",  MIR No: " +txtAckMirno.Text);
                        return Ok("Workcenter " + comboBox_MachineName + " is assigned to total " + totalcompleted + " Production Orders out of " + totalPO + " MIR No: " + mirno);

                        // MessageBox.Show(" Workcenter assigned Workorders: " + releasedWC + " Total Workorders: " + totalPO + " MIR No: " + txtAckMirno.Text + " Workcenter Name " + comboBox_MachineName.Text);
                        //progressBar2.Visible = false;
                        //lblUpdationBar.Visible = false;
                        //return;
                    }
                }
                #endregion JBP


                #region BUB
                else if (plantCode == "TM03")
                {
                    if ((cmb_Group != "Shop 1") && (cmb_Group != "Shop 2") && (cmb_Group != "Shop 3") && (cmb_Group != "Shop 4"))
                    {
                        // cn.Open();
                        // Conn.Open();
                        int notch = 0, meel = 0, weld = 0, bend = 0, total = 0, i = 0; ;
                        int[] ar = new int[10];
                        string opstatus = "", PO = "";
                        SqlCommand cd = new SqlCommand("select * from Operations where Mirno= '" + mirno + "' and MachineName is null", Conn);
                        SqlDataReader dr1 = cd.ExecuteReader();
                        while (dr1.Read())
                        {
                            PO = dr1["RSNo"].ToString();
                            opstatus = dr1["OPStatus"].ToString();
                            i = 0;

                            #region All
                            if (opstatus.Contains("N") && opstatus.Contains("B") && opstatus.Contains("M"))
                            {
                                ar[0] = 10;
                                ar[1] = 40;
                                ar[2] = 50;
                                ar[3] = 70;
                                ar[4] = 90;
                                ar[5] = 100;

                                //string All = Application.StartupPath + "\\" + "All.xml";
                                xml1.Load(All);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(All);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName;
                                    xml1.Save(All);
                                }
                                nodes = xml1.SelectNodes("//IROPNO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = ar[i].ToString();
                                    xml1.Save(All);
                                    i++;
                                }
                                try
                                {
                                    sr = new StreamReader(WCAssignini);
                                    uri = sr.ReadLine();
                                    sr.Close();

                                    credentials = new NetworkCredential(UserName, Password);
                                    cc = new CredentialCache();
                                    cc.Add(new Uri(uri), "Basic", credentials);
                                    req = WebRequest.Create(uri);
                                    req.Method = "POST";
                                    req.ContentType = "text/xml";
                                    writer = new StreamWriter(req.GetRequestStream());
                                    writer.WriteLine(this.GetTextFromXMLFile(All));
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
                                    //return;
                                }
                            }
                            #endregion All
                            #region N&B
                            else if (opstatus.Contains("N") && opstatus.Contains("B"))
                            {
                                ar[0] = 10;
                                ar[1] = 40;
                                ar[2] = 50;
                                ar[3] = 70;
                                ar[4] = 100;

                                //string NB = Application.StartupPath + "\\" + "NB.xml";
                                xml1.Load(NB);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(NB);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName;
                                    xml1.Save(NB);
                                }
                                nodes = xml1.SelectNodes("//IROPNO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = ar[i].ToString();
                                    xml1.Save(NB);
                                    i++;
                                }
                                try
                                {
                                    sr = new StreamReader(WCAssignini);
                                    uri = sr.ReadLine();
                                    sr.Close();

                                    credentials = new NetworkCredential(UserName, Password);
                                    cc = new CredentialCache();
                                    cc.Add(new Uri(uri), "Basic", credentials);
                                    req = WebRequest.Create(uri);
                                    req.Method = "POST";
                                    req.ContentType = "text/xml";
                                    writer = new StreamWriter(req.GetRequestStream());
                                    writer.WriteLine(this.GetTextFromXMLFile(NB));
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
                                    //return;
                                }
                            }
                            #endregion N&B
                            #region N&M
                            else if (opstatus.Contains("N") && opstatus.Contains("M"))
                            {
                                ar[0] = 10;
                                ar[1] = 40;
                                ar[2] = 50;
                                ar[3] = 70;
                                ar[4] = 90;

                                //string NM = Application.StartupPath + "\\" + "NM.xml";
                                xml1.Load(NM);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(NM);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName;
                                    xml1.Save(NM);
                                }
                                nodes = xml1.SelectNodes("//IROPNO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = ar[i].ToString();
                                    xml1.Save(NM);
                                    i++;
                                }
                                try
                                {
                                    sr = new StreamReader(WCAssignini);
                                    uri = sr.ReadLine();
                                    sr.Close();

                                    credentials = new NetworkCredential(UserName, Password);
                                    cc = new CredentialCache();
                                    cc.Add(new Uri(uri), "Basic", credentials);
                                    req = WebRequest.Create(uri);
                                    req.Method = "POST";
                                    req.ContentType = "text/xml";
                                    writer = new StreamWriter(req.GetRequestStream());
                                    writer.WriteLine(this.GetTextFromXMLFile(NM));
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
                                    //return;
                                }
                            }
                            #endregion N&M
                            #region B&M
                            else if (opstatus.Contains("B") && opstatus.Contains("M"))
                            {
                                ar[0] = 10;
                                ar[1] = 40;
                                ar[2] = 50;
                                ar[3] = 90;
                                ar[4] = 100;

                                //string BM = Application.StartupPath + "\\" + "BM.xml";
                                xml1.Load(BM);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(BM);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName;
                                    xml1.Save(BM);
                                }
                                nodes = xml1.SelectNodes("//IROPNO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = ar[i].ToString();
                                    xml1.Save(BM);
                                    i++;
                                }
                                try
                                {
                                    sr = new StreamReader(WCAssignini);
                                    uri = sr.ReadLine();
                                    sr.Close();

                                    credentials = new NetworkCredential(UserName, Password);
                                    cc = new CredentialCache();
                                    cc.Add(new Uri(uri), "Basic", credentials);
                                    req = WebRequest.Create(uri);
                                    req.Method = "POST";
                                    req.ContentType = "text/xml";
                                    writer = new StreamWriter(req.GetRequestStream());
                                    writer.WriteLine(this.GetTextFromXMLFile(BM));
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
                                    //return;
                                }
                            }
                            #endregion N&M
                            #region N
                            else if (opstatus.Contains("N"))
                            {
                                ar[0] = 10;
                                ar[1] = 40;
                                ar[2] = 50;
                                ar[3] = 70;

                                //string Notch = Application.StartupPath + "\\" + "Notch.xml";
                                xml1.Load(Notch);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(Notch);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName;
                                    xml1.Save(Notch);
                                }
                                nodes = xml1.SelectNodes("//IROPNO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = ar[i].ToString();
                                    xml1.Save(Notch);
                                    i++;
                                }
                                try
                                {
                                    sr = new StreamReader(WCAssignini);
                                    uri = sr.ReadLine();
                                    sr.Close();

                                    credentials = new NetworkCredential(UserName, Password);
                                    cc = new CredentialCache();
                                    cc.Add(new Uri(uri), "Basic", credentials);
                                    req = WebRequest.Create(uri);
                                    req.Method = "POST";
                                    req.ContentType = "text/xml";
                                    writer = new StreamWriter(req.GetRequestStream());
                                    writer.WriteLine(this.GetTextFromXMLFile(Notch));
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
                                    //return;
                                }
                            }
                            #endregion N
                            #region M
                            else if (opstatus.Contains("M"))
                            {
                                ar[0] = 10;
                                ar[1] = 40;
                                ar[2] = 50;
                                ar[3] = 90;

                                //string HM = Application.StartupPath + "\\" + "HM.xml";
                                xml1.Load(HM);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(HM);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName;
                                    xml1.Save(HM);
                                }
                                nodes = xml1.SelectNodes("//IROPNO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = ar[i].ToString();
                                    xml1.Save(HM);
                                    i++;
                                }
                                try
                                {
                                    sr = new StreamReader(WCAssignini);
                                    uri = sr.ReadLine();
                                    sr.Close();

                                    credentials = new NetworkCredential(UserName, Password);
                                    cc = new CredentialCache();
                                    cc.Add(new Uri(uri), "Basic", credentials);
                                    req = WebRequest.Create(uri);
                                    req.Method = "POST";
                                    req.ContentType = "text/xml";
                                    writer = new StreamWriter(req.GetRequestStream());
                                    writer.WriteLine(this.GetTextFromXMLFile(HM));
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
                                    //return;
                                }
                            }
                            #endregion M
                            #region B
                            else if (opstatus.Contains("B"))
                            {
                                ar[0] = 10;
                                ar[1] = 40;
                                ar[2] = 50;
                                ar[3] = 100;

                                //string Bend = Application.StartupPath + "\\" + "Bend.xml";
                                xml1.Load(Bend);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(Bend);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName;
                                    xml1.Save(Bend);
                                }
                                nodes = xml1.SelectNodes("//IROPNO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = ar[i].ToString();
                                    xml1.Save(Bend);
                                    i++;
                                }
                                try
                                {
                                    sr = new StreamReader(WCAssignini);
                                    uri = sr.ReadLine();
                                    sr.Close();

                                    credentials = new NetworkCredential(UserName, Password);
                                    cc = new CredentialCache();
                                    cc.Add(new Uri(uri), "Basic", credentials);
                                    req = WebRequest.Create(uri);
                                    req.Method = "POST";
                                    req.ContentType = "text/xml";
                                    writer = new StreamWriter(req.GetRequestStream());
                                    writer.WriteLine(this.GetTextFromXMLFile(Bend));
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
                                    //return;
                                }
                            }
                            #endregion B

                            #region OnlyFab
                            else
                            {
                                xml1.Load(WCAssign);
                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = PO.ToString();
                                    xml1.Save(WCAssign);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName;
                                    xml1.Save(WCAssign);
                                }
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
                                    sr = new StreamReader(WCAssignini);
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
                                catch (Exception e)
                                {
                                    return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                    //return;
                                }
                            }

                            #endregion OnlyFab

                            SqlCommand c3 = new SqlCommand("update Operations set PlanningDate='" + scan_Date1 + "',  MachineName='" + comboBox_MachineName + "',PrimaryWC='" + comboBox_MachineName + "',PlanningShift='" + cmbShift + "',RackDetails='" + txtRack + "' where RSNo='" + PO.Trim() + "'", Conn);
                            c3.ExecuteNonQuery();
                            //progressBar2.Value += 1;
                            releasedWC += 1;
                        }
                        dr1.Close();
                        //MessageBox.Show(comboBox_MachineName.Text + " WorkCenter is assigned to MIR:" + txtMirno.Text);

                        #region Workcenter Change
                        if (totalPO == completedPO)
                        {
                            completedPO = 0;
                            //progressBar2.Maximum = totalPO;
                            c1 = new SqlCommand("select * from Operations where Mirno= '" + mirno + "' ", Conn); //and POType='Primary'
                            SqlDataReader drChange = c1.ExecuteReader();
                            while (drChange.Read())
                            {
                                PO = drChange["RSNo"].ToString();
                                opstatus = drChange["OPStatus"].ToString();
                                i = 0;

                                #region All
                                if (opstatus.Contains("N") && opstatus.Contains("B") && opstatus.Contains("M"))
                                {
                                    ar[0] = 10;
                                    ar[1] = 40;
                                    ar[2] = 50;
                                    ar[3] = 70;
                                    ar[4] = 90;
                                    ar[5] = 100;

                                    //string All = Application.StartupPath + "\\" + "All.xml";
                                    xml1.Load(All);
                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(All);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(All);
                                    }
                                    nodes = xml1.SelectNodes("//IROPNO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = ar[i].ToString();
                                        xml1.Save(All);
                                        i++;
                                    }
                                    try
                                    {
                                        sr = new StreamReader(WCAssignini);
                                        uri = sr.ReadLine();
                                        sr.Close();

                                        credentials = new NetworkCredential(UserName, Password);
                                        cc = new CredentialCache();
                                        cc.Add(new Uri(uri), "Basic", credentials);
                                        req = WebRequest.Create(uri);
                                        req.Method = "POST";
                                        req.ContentType = "text/xml";
                                        writer = new StreamWriter(req.GetRequestStream());
                                        writer.WriteLine(this.GetTextFromXMLFile(All));
                                        writer.Close();
                                        req.Credentials = cc;
                                        rsp = req.GetResponse();
                                        sr = new StreamReader(rsp.GetResponseStream());
                                        result = sr.ReadToEnd();
                                        sr.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                        //return;
                                    }
                                }
                                #endregion All
                                #region N&B
                                else if (opstatus.Contains("N") && opstatus.Contains("B"))
                                {
                                    ar[0] = 10;
                                    ar[1] = 40;
                                    ar[2] = 50;
                                    ar[3] = 70;
                                    ar[4] = 100;

                                    //string NB = Application.StartupPath + "\\" + "NB.xml";
                                    xml1.Load(NB);
                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(NB);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(NB);
                                    }
                                    nodes = xml1.SelectNodes("//IROPNO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = ar[i].ToString();
                                        xml1.Save(NB);
                                        i++;
                                    }
                                    try
                                    {
                                        sr = new StreamReader(WCAssignini);
                                        uri = sr.ReadLine();
                                        sr.Close();

                                        credentials = new NetworkCredential(UserName, Password);
                                        cc = new CredentialCache();
                                        cc.Add(new Uri(uri), "Basic", credentials);
                                        req = WebRequest.Create(uri);
                                        req.Method = "POST";
                                        req.ContentType = "text/xml";
                                        writer = new StreamWriter(req.GetRequestStream());
                                        writer.WriteLine(this.GetTextFromXMLFile(NB));
                                        writer.Close();
                                        req.Credentials = cc;
                                        rsp = req.GetResponse();
                                        sr = new StreamReader(rsp.GetResponseStream());
                                        result = sr.ReadToEnd();
                                        sr.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                        //return;
                                    }
                                }
                                #endregion N&B
                                #region N&M
                                else if (opstatus.Contains("N") && opstatus.Contains("M"))
                                {
                                    ar[0] = 10;
                                    ar[1] = 40;
                                    ar[2] = 50;
                                    ar[3] = 70;
                                    ar[4] = 90;

                                    //string NM = Application.StartupPath + "\\" + "NM.xml";
                                    xml1.Load(NM);
                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(NM);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(NM);
                                    }
                                    nodes = xml1.SelectNodes("//IROPNO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = ar[i].ToString();
                                        xml1.Save(NM);
                                        i++;
                                    }
                                    try
                                    {
                                        sr = new StreamReader(WCAssignini);
                                        uri = sr.ReadLine();
                                        sr.Close();

                                        credentials = new NetworkCredential(UserName, Password);
                                        cc = new CredentialCache();
                                        cc.Add(new Uri(uri), "Basic", credentials);
                                        req = WebRequest.Create(uri);
                                        req.Method = "POST";
                                        req.ContentType = "text/xml";
                                        writer = new StreamWriter(req.GetRequestStream());
                                        writer.WriteLine(this.GetTextFromXMLFile(NM));
                                        writer.Close();
                                        req.Credentials = cc;
                                        rsp = req.GetResponse();
                                        sr = new StreamReader(rsp.GetResponseStream());
                                        result = sr.ReadToEnd();
                                        sr.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                        //return;
                                    }
                                }
                                #endregion N&M
                                #region B&M
                                else if (opstatus.Contains("B") && opstatus.Contains("M"))
                                {

                                    ar[0] = 10;
                                    ar[1] = 40;
                                    ar[2] = 50;
                                    ar[3] = 90;
                                    ar[4] = 100;

                                    //string BM = Application.StartupPath + "\\" + "BM.xml";
                                    xml1.Load(BM);
                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(BM);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(BM);
                                    }
                                    nodes = xml1.SelectNodes("//IROPNO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = ar[i].ToString();
                                        xml1.Save(BM);
                                        i++;
                                    }
                                    try
                                    {
                                        sr = new StreamReader(WCAssignini);
                                        uri = sr.ReadLine();
                                        sr.Close();

                                        credentials = new NetworkCredential(UserName, Password);
                                        cc = new CredentialCache();
                                        cc.Add(new Uri(uri), "Basic", credentials);
                                        req = WebRequest.Create(uri);
                                        req.Method = "POST";
                                        req.ContentType = "text/xml";
                                        writer = new StreamWriter(req.GetRequestStream());
                                        writer.WriteLine(this.GetTextFromXMLFile(BM));
                                        writer.Close();
                                        req.Credentials = cc;
                                        rsp = req.GetResponse();
                                        sr = new StreamReader(rsp.GetResponseStream());
                                        result = sr.ReadToEnd();
                                        sr.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                        //return;
                                    }
                                }
                                #endregion N&M
                                #region N
                                else if (opstatus.Contains("N"))
                                {
                                    ar[0] = 10;
                                    ar[1] = 40;
                                    ar[2] = 50;
                                    ar[3] = 70;

                                    //string Notch = Application.StartupPath + "\\" + "Notch.xml";
                                    xml1.Load(Notch);
                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(Notch);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(Notch);
                                    }
                                    nodes = xml1.SelectNodes("//IROPNO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = ar[i].ToString();
                                        xml1.Save(Notch);
                                        i++;
                                    }
                                    try
                                    {
                                        sr = new StreamReader(WCAssignini);
                                        uri = sr.ReadLine();
                                        sr.Close();

                                        credentials = new NetworkCredential(UserName, Password);
                                        cc = new CredentialCache();
                                        cc.Add(new Uri(uri), "Basic", credentials);
                                        req = WebRequest.Create(uri);
                                        req.Method = "POST";
                                        req.ContentType = "text/xml";
                                        writer = new StreamWriter(req.GetRequestStream());
                                        writer.WriteLine(this.GetTextFromXMLFile(Notch));
                                        writer.Close();
                                        req.Credentials = cc;
                                        rsp = req.GetResponse();
                                        sr = new StreamReader(rsp.GetResponseStream());
                                        result = sr.ReadToEnd();
                                        sr.Close();
                                    }
                                    catch (Exception e)
                                    {
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                        //return;
                                    }
                                }
                                #endregion N
                                #region M
                                else if (opstatus.Contains("M"))
                                {
                                    ar[0] = 10;
                                    ar[1] = 40;
                                    ar[2] = 50;
                                    ar[3] = 90;

                                    //string HM = Application.StartupPath + "\\" + "HM.xml";
                                    xml1.Load(HM);
                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(HM);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(HM);
                                    }
                                    nodes = xml1.SelectNodes("//IROPNO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = ar[i].ToString();
                                        xml1.Save(HM);
                                        i++;
                                    }
                                    try
                                    {
                                        sr = new StreamReader(WCAssignini);
                                        uri = sr.ReadLine();
                                        sr.Close();

                                        credentials = new NetworkCredential(UserName, Password);
                                        cc = new CredentialCache();
                                        cc.Add(new Uri(uri), "Basic", credentials);
                                        req = WebRequest.Create(uri);
                                        req.Method = "POST";
                                        req.ContentType = "text/xml";
                                        writer = new StreamWriter(req.GetRequestStream());
                                        writer.WriteLine(this.GetTextFromXMLFile(HM));
                                        writer.Close();
                                        req.Credentials = cc;
                                        rsp = req.GetResponse();
                                        sr = new StreamReader(rsp.GetResponseStream());
                                        result = sr.ReadToEnd();
                                        sr.Close();
                                    }
                                    catch (Exception ex)
                                    {
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                        //return;
                                    }
                                }
                                #endregion M
                                #region B
                                else if (opstatus.Contains("B"))
                                {
                                    ar[0] = 10;
                                    ar[1] = 40;
                                    ar[2] = 50;
                                    ar[3] = 100;

                                    //string Bend = Application.StartupPath + "\\" + "Bend.xml";
                                    xml1.Load(Bend);
                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(Bend);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(Bend);
                                    }
                                    nodes = xml1.SelectNodes("//IROPNO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = ar[i].ToString();
                                        xml1.Save(Bend);
                                        i++;
                                    }
                                    try
                                    {
                                        sr = new StreamReader(WCAssignini);
                                        uri = sr.ReadLine();
                                        sr.Close();

                                        credentials = new NetworkCredential(UserName, Password);
                                        cc = new CredentialCache();
                                        cc.Add(new Uri(uri), "Basic", credentials);
                                        req = WebRequest.Create(uri);
                                        req.Method = "POST";
                                        req.ContentType = "text/xml";
                                        writer = new StreamWriter(req.GetRequestStream());
                                        writer.WriteLine(this.GetTextFromXMLFile(Bend));
                                        writer.Close();
                                        req.Credentials = cc;
                                        rsp = req.GetResponse();
                                        sr = new StreamReader(rsp.GetResponseStream());
                                        result = sr.ReadToEnd();
                                        sr.Close();
                                    }
                                    catch (Exception ex)
                                    {
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                        //return;
                                    }
                                }
                                #endregion B

                                #region OnlyFab
                                else
                                {
                                    xml1.Load(WCAssign);

                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = PO.ToString();
                                        xml1.Save(WCAssign);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName;
                                        xml1.Save(WCAssign);
                                    }
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
                                        sr = new StreamReader(WCAssignini);
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
                                        //progressBar2.Visible = false;
                                        //lblUpdationBar.Visible = false;
                                        return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                        //return;
                                    }
                                }

                                #endregion OnlyFab
                                //progressBar2.Value += 1;
                                releasedWC += 1;
                            }
                            drChange.Close();

                            c1 = new SqlCommand("insert into Operations(RSNo,FGItem,Operation,Pices,ActualPiece,Wheight,Setups,SctDinemtion,PlanningDate,PlanningShift,Operator,Status,Mirno,LotCode,BP,CountStatus,TotalWt,Tot_OPS,Length,MachineName,JDDate,CompletedQty,OPStatus,RunTime,Flag_Ack,PrimaryWC,POType) select RSNo,FGItem,Operation,Pices,ActualPiece,Wheight,Setups,SctDinemtion,'" + scan_Date1 + "','" + cmbShift + "',Operator,Status,Mirno,LotCode,BP,CountStatus,TotalWt,Tot_OPS,Length,'" + comboBox_MachineName + "',JDDate,CompletedQty,OPStatus,RunTime,Flag_Ack,PrimaryWC,'Duplicate' from Operations where Mirno='" + mirno + "' and MachineName=(select TOP 1 MachineName from Operations where Mirno='" + mirno + "') and PlanningDate=(select TOP 1 PlanningDate from Operations where Mirno='" + mirno + "') and PlanningShift=(select TOP 1 PlanningShift from Operations where Mirno='" + mirno + "')", Conn);
                            c1.ExecuteNonQuery();

                        }
                        #endregion WorkCenter Change

                        //***Code moved in common function : bindDataGridAfterAllocate

                        //string CommandText = "SELECT RSNo as [PO.NO.],FGItem as [Item No],Mirno as MIRNO,MachineName as WorkCenter,SctDinemtion as SECTION,OPStatus,LotCode as [Billable Lot],Pices as QTY,Length,Wheight as Weight,RackDetails,PlanningShift,SAPPulledDate as PulledDate,JDDate as ReleaseDate,PlanningDate,Operation,Tot_OPS,RunTime,Status,TotalWt FROM Operations where Mirno='" + mirno + "' and BP='" + plantCode + "' and Flag_Fab is null order by [index] desc ";
                        //SqlCommand myCommand = new SqlCommand(CommandText, Conn);
                        //SqlDataAdapter myAdapter = new SqlDataAdapter();
                        //myAdapter.SelectCommand = myCommand;
                        //DataSet myDataSet = new DataSet();
                        //myAdapter.Fill(myDataSet);
                        //dataGridView1.DataSource = myDataSet.Tables[0];

                        int totalcompleted = completedPO + releasedWC;
                        //MessageBox.Show("Workcenter " + comboBox_MachineName.Text + " is assigned to total " + totalcompleted + "Production Orders. out of" + totalPO + " MIR No: " + txtAckMirno.Text);
                        return Ok("Workcenter " + comboBox_MachineName + " is assigned to total " + totalcompleted + " Production Orders out of " + totalPO + " MIR No: " + mirno);
                        // MessageBox.Show(" Workcenter assigned Workorders: " + releasedWC + " Total Workorders: " + totalPO + " MIR No: " + txtAckMirno.Text + " Workcenter Name " + comboBox_MachineName.Text);
                        //progressBar2.Visible = false;
                        //lblUpdationBar.Visible = false;
                        //return;
                    }
                }

                #endregion Bub

                #endregion ShopWCAssign

                //txtwheight.Text = "";
                //txtRs.Text = "";
                //txtOprns.Text = "";
                //txtRunTime.Text = "";
                c1 = new SqlCommand("select MachineType from WorkCenterMaster where WorkCenterCode = '" + comboBox_MachineName + "' ", Conn);
                SqlDataReader dr2 = c1.ExecuteReader();
                if (dr2.Read())
                {
                    maachinetype = dr2["MachineType"].ToString();
                }
                dr2.Close();
                //DateTime scandate = DateTime.Parse(dateTimePicker1.Text);
                DateTime scandate = DateTime.Now;
                string scan_Date = scandate.Month + "/" + scandate.Day + "/" + scandate.Year + " " + scandate.TimeOfDay;

                if (mirno != "" && cmbShift != "" && comboBox_MachineName != "")
                {
                    string MCName = "", shift = "";
                    DateTime plandate = DateTime.Now;
                    SqlDataReader dr = null;
                    if (totalPO == completedPO)
                    {
                        c1 = new SqlCommand("select * from Operations where Mirno = '" + mirno + "' and MachineName='" + comboBox_MachineName + "' and PlanningShift='" + cmbShift + "' and PlanningDate='" + scan_Date + "'", Conn);
                        SqlDataReader dr1 = c1.ExecuteReader();
                        if (dr1.Read())
                        {
                            dr1.Close();
                            return Ok("MIR Already assign for this  machine in selected shift ");
                            //cn.Close();
                            //return;
                        }
                        dr1.Close();
                        c1 = new SqlCommand("select * from Operations where Mirno='" + mirno + "' and MachineName is not null", Conn);
                        dr = c1.ExecuteReader();
                        if (dr.Read())
                        {
                            String AssignedMC = dr["MachineName"].ToString();
                            string AssignedShift = dr["PlanningShift"].ToString();
                            DateTime PlanDateCheck = DateTime.Parse(dr["PlanningDate"].ToString());
                            string Plan_Date = PlanDateCheck.Month + "/" + PlanDateCheck.Day + "/" + PlanDateCheck.Year + " " + PlanDateCheck.TimeOfDay;
                            dr.Close();

                            //DialogResult diaResult = MessageBox.Show("MIR already assigned: Do you want continue", "warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                            //if (diaResult == DialogResult.Yes)
                            //{
                            if (AssignedMC == comboBox_MachineName && AssignedShift != cmbShift)
                            {
                                c1 = new SqlCommand("insert into Operations(RSNo,FGItem,Operation,Pices,ActualPiece,Wheight,Setups,SctDinemtion,PlanningDate,PlanningShift,Operator,Status,Mirno,LotCode,BP,CountStatus,TotalWt,Tot_OPS,Length,MachineName,JDDate,CompletedQty,OPStatus,RunTime,Flag_Ack,PrimaryWC,POType) select RSNo,FGItem,Operation,Pices,ActualPiece,Wheight,Setups,SctDinemtion,'" + scan_Date + "','" + cmbShift + "',Operator,Status,Mirno,LotCode,BP,CountStatus,TotalWt,Tot_OPS,Length,'" + comboBox_MachineName + "',JDDate,CompletedQty,OPStatus,RunTime,Flag_Ack,PrimaryWC,'Duplicate' from Operations where Mirno='" + mirno + "' and MachineName=(select TOP 1 MachineName from Operations where Mirno='" + mirno + "') and PlanningDate=(select TOP 1 PlanningDate from Operations where Mirno='" + mirno + "') and PlanningShift=(select TOP 1 PlanningShift from Operations where Mirno='" + mirno + "')", Conn);
                                c1.ExecuteNonQuery();
                            }
                            else
                            {
                                c2 = new SqlCommand("update Operations set PlanningDate='" + scan_Date + "',  MachineName='" + comboBox_MachineName + "',PlanningShift='" + cmbShift + "',RunTime=" + Runtime + ",PrimaryWC='" + comboBox_MachineName + "' where  Mirno='" + mirno + "' ", Conn); //and POType='Primary'
                                c2.ExecuteNonQuery();

                                c2 = new SqlCommand("delete from Operations where Mirno='" + mirno + "' and POType='Duplicate'", Conn);
                                c2.ExecuteNonQuery();
                            }
                            c1 = new SqlCommand("select RSNo  from Operations where Mirno='" + mirno + "' and MachineName=(select TOP 1 MachineName from Operations where Mirno='" + mirno + "') and PlanningDate=(select TOP 1 PlanningDate from Operations where Mirno='" + mirno + "') and PlanningShift=(select TOP 1 PlanningShift from Operations where Mirno='" + mirno + "')", Conn);
                            dr = c1.ExecuteReader();
                            while (dr.Read())
                            {
                                routeSheet = dr["RSNo"].ToString();
                                //****************************SAP Update**************************//
                                //else
                                {
                                    xml1.Load(WCAssign);
                                    nodes = xml1.SelectNodes("//WADOCO");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = routeSheet;
                                        xml1.Save(WCAssign);
                                    }
                                    nodes = xml1.SelectNodes("//Machine");
                                    foreach (XmlElement element1 in nodes)
                                    {
                                        element1.InnerText = comboBox_MachineName.Trim();
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
                                }
                                try
                                {
                                    sr = new StreamReader(WCAssignini);
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
                                    //progressBar2.Value += 1;
                                }
                                catch (Exception e)
                                {
                                    return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                    //return;
                                }
                                //c1 = new SqlCommand("insert into Operations(RSNo,FGItem,Operation,Pices,ActualPiece,Wheight,Setups,SctDinemtion,PlanningDate,PlanningShift,Operator,Status,Mirno,LotCode,BP,CountStatus,TotalWt,Tot_OPS,Length,MachineName,JDDate,CompletedQty,OPStatus,RunTime) select RSNo,FGItem,Operation,Pices,ActualPiece,Wheight,Setups,SctDinemtion,'" + scan_Date + "','" + cmbShift.Text + "',Operator,Status,Mirno,LotCode,BP,CountStatus,TotalWt,Tot_OPS,Length,'" + comboBox_MachineName.Text + "',JDDate,CompletedQty,OPStatus,RunTime from Operations where Mirno='" + txtMirno.Text + "' and MachineName=(select TOP 1 MachineName from Operations where Mirno='" + txtMirno.Text + "') and PlanningDate=(select TOP 1 PlanningDate from Operations where Mirno='" + txtMirno.Text + "') and PlanningShift=(select TOP 1 PlanningShift from Operations where Mirno='" + txtMirno.Text + "')", cn);
                                //c1.ExecuteNonQuery();
                            }
                            //****************************************************************//
                            dr.Close();

                            //***Code moved in common function : getCalculationsAfterAllocate

                            //c1 = new SqlCommand("select sum(TotalWt)/1000 as TotalWheight,count(RSNo) as RSno,sum(Tot_OPS) as TotalOpns,sum(RunTime) as RunTime from Operations where PlanningDate='" + scan_Date + "'and  MachineName='" + comboBox_MachineName + "' and PlanningShift='" + cmbShift + "' and POType='Primary' ", Conn);

                            //dr = c1.ExecuteReader();
                            //if (dr.Read())
                            //{
                            //    if (dr["TotalWheight"].ToString() != "")
                            //        wheight = Math.Round(double.Parse(dr["TotalWheight"].ToString()), 3);

                            //    txtRs.Text = dr["RSno"].ToString();
                            //    if (dr["TotalOpns"].ToString() != "")
                            //        Operations = double.Parse(dr["TotalOpns"].ToString());
                            //    if (dr["RunTime"].ToString() != "")
                            //        Runtime1 = Math.Round(double.Parse(dr["RunTime"].ToString()), 2);
                            //}
                            //txtwheight.Text = wheight.ToString();
                            //txtOprns.Text = Operations.ToString();
                            //txtRunTime.Text = Runtime1.ToString();
                            //dr.Close();
                            //
                            //}
                        }
                    }
                    else
                    {
                        //development
                        //dr.Close();
                        c1 = new SqlCommand("select * from Operations where Mirno='" + mirno + "' and MachineName is null", Conn);
                        dr = c1.ExecuteReader();
                        //Conn.Open();
                        while (dr.Read())
                        {
                            routeSheet = dr["RSNo"].ToString();
                            length = float.Parse(dr["Length"].ToString());
                            pices = float.Parse(dr["Pices"].ToString());
                            operations = float.Parse(dr["Tot_OPS"].ToString());
                            section = dr["SctDinemtion"].ToString().Trim();
                            xml1 = new XmlDocument();
                            if (comboBox_MachineName.Contains("GALV"))
                            {
                                xml1.Load(GWCAssign);

                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = routeSheet;
                                    xml1.Save(GWCAssign);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName.Trim();
                                    xml1.Save(GWCAssign);
                                }
                                nodes = xml1.SelectNodes("//IROPNO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    int status = 230;
                                    element1.InnerText = status.ToString();
                                    xml1.Save(GWCAssign);
                                }
                            }
                            else
                            {
                                xml1.Load(WCAssign);

                                nodes = xml1.SelectNodes("//WADOCO");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = routeSheet;
                                    xml1.Save(WCAssign);
                                }
                                nodes = xml1.SelectNodes("//Machine");
                                foreach (XmlElement element1 in nodes)
                                {
                                    element1.InnerText = comboBox_MachineName.Trim();
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
                            }
                            sr = new StreamReader(WCAssignini);
                            uri = sr.ReadLine();
                            sr.Close();
                            try
                            {
                                credentials = new NetworkCredential(UserName, Password);
                                cc = new CredentialCache();
                                cc.Add(new Uri(uri), "Basic", credentials);
                                req = WebRequest.Create(uri);
                                req.Method = "POST";
                                req.ContentType = "text/xml";
                                writer = new StreamWriter(req.GetRequestStream());
                                if (comboBox_MachineName.Contains("GALV"))
                                {
                                    writer.WriteLine(this.GetTextFromXMLFile(GWCAssign));
                                }
                                else
                                {
                                    writer.WriteLine(this.GetTextFromXMLFile(WCAssign));
                                }
                                writer.Close();
                                req.Credentials = cc;
                                rsp = req.GetResponse();
                                sr = new StreamReader(rsp.GetResponseStream());
                                result = sr.ReadToEnd();
                                sr.Close();
                            }
                            catch (Exception ex)
                            {
                                return Ok("Connectivity failed with SAP server!!!.  Please try again..!!!");
                                //return;
                            }
                            if (section == "")
                            {
                                if (showMessage == true)
                                {
                                    showMessage = false;
                                    Console.Write("Section is not available for given MIR. It will affect on Runtime Calculation");
                                }
                                Thickness = 0;
                            }
                            else
                            {
                                section = section.Remove(0, (section.Length - 2));//comment for testing
                                if (section.Contains("X"))
                                {
                                    Thickness = float.Parse(section.Remove(0, 1));
                                }
                                else
                                {
                                    bool containsLetter = Regex.IsMatch(section, "[A-Z]");
                                    if (containsLetter)
                                    {
                                        if (Regex.IsMatch(section.Remove(0, 1), "[A-Z]"))
                                            Thickness = 0;
                                        else
                                            Thickness = float.Parse(section.Remove(0, 1));
                                    }
                                    if (!containsLetter)
                                    {

                                        Thickness = float.Parse(section);
                                    }

                                }
                            }

                            if (Thickness >= 4 && Thickness <= 10)
                            {
                                YFactoreConst = 35;
                            }
                            else
                            {
                                YFactoreConst = 45;
                            }
                            //if (rdrFicep.Checked)
                            if (maachinetype == "FICEP")
                            {
                                c2 = new SqlCommand("select * from RunTimeCalc where MachineType = 'FICEP' and Thickness=" + Thickness + "", Conn);
                                drpara = c2.ExecuteReader();
                                if (drpara.Read())
                                {
                                    stampara = float.Parse(drpara["stamping"].ToString());
                                    punchpara = float.Parse(drpara["Punching"].ToString());
                                    cutpara = float.Parse(drpara["Cutting"].ToString());
                                }
                                drpara.Close();
                                GripAllow = 2;
                                CarriageSpeed = 60000;
                                FactoreSpeed = 40000;
                            }
                            else if (maachinetype == "VERNET")
                            {
                                c2 = new SqlCommand("select * from RunTimeCalc where MachineType = 'VERNET' and Thickness=" + Thickness + "", Conn);
                                drpara = c2.ExecuteReader();
                                if (drpara.Read())
                                {
                                    stampara = float.Parse(drpara["stamping"].ToString());
                                    punchpara = float.Parse(drpara["Punching"].ToString());
                                    cutpara = float.Parse(drpara["Cutting"].ToString());
                                }
                                drpara.Close();
                                GripAllow = 3;
                                CarriageSpeed = 30000;
                                FactoreSpeed = 30000;

                            }
                            if (FactoreSpeed == 0)
                                FactoreSpeed = 100;
                            if (CarriageSpeed == 0)
                                CarriageSpeed = 100;
                            if (GripAllow == 0)
                                GripAllow = 1000;
                            xfactore = (length * pices) / (FactoreSpeed * 1);
                            punch = (operations * punchpara) / (1 * 1);
                            yfactore = (YFactoreConst * operations) / (7000 * 1);
                            cutting = (cutpara * pices) / 1;
                            stamping = (stampara * pices) / 1;
                            setup = 2.75;
                            inspection = 1.5;
                            carriage = (pices * length) / (CarriageSpeed * 1);
                            gripping = (0.45 * pices) / GripAllow;
                            Runtime = xfactore + punch + yfactore + cutting + stamping + setup + inspection + carriage + gripping;
                            c2 = new SqlCommand("update Operations set PlanningDate='" + scan_Date + "',  MachineName='" + comboBox_MachineName + "',PlanningShift='" + cmbShift + "',RunTime=" + Runtime + ",PrimaryWC='" + comboBox_MachineName + "' where RSNo='" + routeSheet + "'", Conn);
                            c2.ExecuteNonQuery();
                            //progressBar2.Value += 1;
                            releasedWC += 1;

                        }
                        dr.Close();
                        //c1 = new SqlCommand("update operations set PlanningDate='" + scan_Date + "',  MachineName='" + comboBox_MachineName.Text + "',PlanningShift='" + cmbShift.Text + "' where Mirno='" + txtMirno.Text + "'", cn);
                        //c1.ExecuteNonQuery();
                    }

                    ////Code moved in common function : bindDataGridAfterAllocate

                    //string CommandText = "SELECT RSNo as [PO.NO.],FGItem as [Item No],Mirno as MIRNO,MachineName as WorkCenter,SctDinemtion as SECTION,OPStatus,LotCode as [Billable Lot],Pices as QTY,Length,Wheight as Weight,Diameter,RackDetails,PlanningShift,SAPPulledDate as PulledDate,JDDate as ReleaseDate,PlanningDate,Operation,Tot_OPS,RunTime,Status,TotalWt FROM Operations where Mirno='" + mirno + "' and BP='" + plantCode + "' and Flag_Fab is null order by [index] desc ";
                    //SqlCommand myCommand = new SqlCommand(CommandText, Conn);
                    //SqlDataAdapter myAdapter = new SqlDataAdapter();
                    //myAdapter.SelectCommand = myCommand;
                    //DataSet myDataSet = new DataSet();
                    //myAdapter.Fill(myDataSet);
                    //dataGridView1.DataSource = myDataSet.Tables[0];

                    ///SqlCommand c = new SqlCommand("select sum(TotalWt)/1000 as TotalWheight,count(RSNo) as RSno,sum(Tot_OPS) as TotalOpns,sum(RunTime) as RunTime from Operations where PlanningDate='" + scan_Date + "'and  MachineName='" + comboBox_MachineName.Text + "' and PlanningShift='" + cmbShift.Text + "' and POType='Primary' ", cn);
 
                    //***Code moved in common function : getCalculationsAfterAllocate

                    // SqlCommand c = new SqlCommand("select sum(TotalWt)/1000 as TotalWheight,count(RSNo) as RSno,sum(Tot_OPS) as TotalOpns,sum(RunTime) as RunTime from Operations where POType='Primary' ", cn);
                    //dr = c.ExecuteReader();
                    //if (dr.Read())
                    //{
                    //    if (dr["TotalWheight"].ToString() != "")
                    //        wheight = Math.Round(double.Parse(dr["TotalWheight"].ToString()), 3);

                    //    txtRs.Text = dr["RSno"].ToString();
                    //    if (dr["TotalOpns"].ToString() != "")
                    //        Operations = double.Parse(dr["TotalOpns"].ToString());
                    //    if (dr["RunTime"].ToString() != "")
                    //        Runtime1 = Math.Round(double.Parse(dr["RunTime"].ToString()), 2);
                    //}
                    //txtwheight.Text = wheight.ToString();
                    //txtOprns.Text = Operations.ToString();
                    //txtRunTime.Text = Runtime1.ToString();
                    //lblUpdationBar.Visible = false;
                    //dr.Close();

                    int totalcompleted = completedPO + releasedWC;
                    //MessageBox.Show(comboBox_MachineName.Text + " WorkCenter is assigned to MIR:" + txtMirno.Text);
                    return Ok("Workcenter " + comboBox_MachineName + " is assigned to total " + totalcompleted + " Production Orders out of " + totalPO + " MIR No: " + mirno);
                    //progressBar2.Visible = false;
                }
                else
                {
                    return Ok("Please Seelct All fields for Allocation");
                }
                //txtMirno.Focus();
            }
            catch (Exception ex)
            {
                //progressBar2.Visible = false;
                //lblUpdationBar.Visible = false;
                //MessageBox.Show(ex.ToString());
                return Ok("Exception Found:" +  ex.Message);
            }
            finally
            {
                Conn.Close();
            }
            return Ok("");
        }


        [HttpGet]
        [Route("bindDataGridAfterAllocate")]
        public async Task<IActionResult> BindDataGridAfterAllocate(string mirno, string plantCode)
        {
            if (Conn.State == ConnectionState.Closed)
                Conn.Open();
            string CommandText = "SELECT RSNo as rsNo,FGItem as Item_No,Mirno as mirNo,MachineName as WorkCenter,SctDinemtion as SECTION,OPStatus,LotCode as billable_Lot,Pices as QTY,Length,Wheight as Wt_Pcs, FORMAT (JDDate, 'dd/MM/yyyy ')  as ReleasedDate,TotalWt,RackDetails,Operation as OPRPCs,Tot_OPS,PlanningShift, FORMAT (PlanningDate, 'dd/MM/yyyy ') as PlanningDate,RunTime,Status FROM Operations where Mirno='" + mirno + "' and BP='" + plantCode + "' and Flag_Fab is null order by [index] desc ";
            SqlCommand myCommand = new SqlCommand(CommandText, Conn);
            SqlDataAdapter myAdapter = new SqlDataAdapter();
            myAdapter.SelectCommand = myCommand;
            DataSet myDataSet = new DataSet();
            myAdapter.Fill(myDataSet);
            Conn.Close();
            return Ok(myDataSet.Tables[0]);
        }

        [HttpGet]
        [Route("getCalculationsAfterAllocate")]
        public async Task<IActionResult> GetCalculationsAfterAllocate(string mirno, string plantCode, string poType)
        {
            if (Conn.State == ConnectionState.Closed)
                Conn.Open();
            string CommandText = "select round(sum(TotalWt)/1000,3) as TotalWheight,count(RSNo) as RSno,round(sum(Tot_OPS),2) as TotalOpns,round(sum(RunTime),2) as RunTime from Operations where Mirno='" + mirno + "' and BP='" + plantCode + "' and Flag_Fab is null "; //and POType='Primary'
            SqlCommand myCommand = new SqlCommand(CommandText, Conn);
            SqlDataAdapter myAdapter = new SqlDataAdapter();
            myAdapter.SelectCommand = myCommand;
            DataSet myDataSet = new DataSet();
            myAdapter.Fill(myDataSet);
            Conn.Close();
            return Ok(myDataSet.Tables[0]);
        }

        [HttpGet]
        [Route("insertUpdateDelete")]
        public async Task<IActionResult> InsertUpdateDelete(string query)
        {
            try
            {
                if (Conn.State == ConnectionState.Closed)
                    Conn.Open();
                SqlCommand com = new SqlCommand(query, Conn);
                int numberOfRecords = com.ExecuteNonQuery();
                if (numberOfRecords > 0)
                    return Ok("true");
                else
                    return Ok("false");
            }
            catch (SqlException ex)
            {
                //Debug.Write(ex.ToString());
                return Ok("false");
            }
            finally
            {
                Conn.Close();
            }
        }


        [HttpGet]
        [Route("onLoadScreen")]
        public async Task<IActionResult> OnLoadScreen(string mirno, bool rdrFicep, bool rdrVernet, bool rdrDrilling)
        {
            double xfactore = 0, Thickness = 0, GripAllow = 0, CarriageSpeed = 0, YFactoreConst = 0, FactoreSpeed = 0, punchpara = 0, cutpara = 0, stampara = 0, yfactore = 0, punch = 0, cutting = 0, stamping = 0, carriage = 0, gripping = 0, length = 0, pices = 0, operations = 0, setup = 0, inspection = 0, Runtime = 0;
            string routeSheet = "", section = "";

            SqlCommand c1 = new SqlCommand();
            SqlCommand c2 = new SqlCommand();
            SqlDataReader dr;
            SqlDataReader drpara;

            if (Conn.State == ConnectionState.Closed)
                Conn.Open();

            c1 = new SqlCommand("select * from operations where Mirno='" +mirno + "' and PlanningDate is null", Conn);
            dr = c1.ExecuteReader();

            //Conn.Open();
            while (dr.Read())
            {
                routeSheet = dr["RSNo"].ToString();
                length = float.Parse(dr["Length"].ToString());
                pices = float.Parse(dr["Pices"].ToString());
                operations = float.Parse(dr["Tot_OPS"].ToString());
                section = dr["SctDinemtion"].ToString().Trim();
                section = section.Remove(0, (section.Length - 2));
                if (section.Contains("X"))
                {
                    Thickness = float.Parse(section.Remove(0, 1));
                }
                else
                {
                    Thickness = float.Parse(section);
                }
                if (Thickness >= 4 && Thickness <= 10)
                {
                    YFactoreConst = 35;
                }
                else
                {
                    YFactoreConst = 45;
                }
                if (rdrFicep == true)
                {

                    c2 = new SqlCommand("select * from RunTimeCalc where MachineType = 'FICEP' and Thickness=" + Thickness + "", Conn);
                    drpara = c2.ExecuteReader();
                    if (drpara.Read())
                    {
                        stampara = float.Parse(drpara["stamping"].ToString());
                        punchpara = float.Parse(drpara["Punching"].ToString());
                        cutpara = float.Parse(drpara["Cutting"].ToString());
                    }
                    drpara.Close();
                    GripAllow = 2;
                    CarriageSpeed = 60000;
                    FactoreSpeed = 40000;
                }
                else if (rdrVernet == true)
                {
                    c2 = new SqlCommand("select * from RunTimeCalc where MachineType = 'VERNET' and Thickness=" + Thickness + "", Conn);
                    drpara = c2.ExecuteReader();
                    if (drpara.Read())
                    {
                        stampara = float.Parse(drpara["stamping"].ToString());
                        punchpara = float.Parse(drpara["Punching"].ToString());
                        cutpara = float.Parse(drpara["Cutting"].ToString());
                    }
                    drpara.Close();
                    GripAllow = 3;
                    CarriageSpeed = 30000;
                    FactoreSpeed = 30000;

                }
                xfactore = (length * pices) / (FactoreSpeed * 1);
                punch = (operations * punchpara) / (1 * 1);
                yfactore = (YFactoreConst * operations) / (7000 * 1);
                cutting = (cutpara * pices) / 1;
                stamping = (stampara * pices) / 1;
                setup = 2.75;
                inspection = 1.5;
                carriage = (pices * length) / (CarriageSpeed * 1);
                gripping = (0.45 * pices) / GripAllow;
                Runtime = xfactore + punch + yfactore + cutting + stamping + setup + inspection + carriage + gripping;
                c2 = new SqlCommand("update operations set RunTime=" + Runtime + " where RSNo='" + routeSheet + "'", Conn);
                c2.ExecuteNonQuery();
            }

            dr.Close(); 

            //double wheight = 0, Operations = 0, Runtime1 = 0;
            //int rs = 0;

            string CommandText = "select round(sum(TotalWt)/1000,3) as TotalWheight,count(RSNo) as RSno,round(sum(Tot_OPS),2) as TotalOpns,round(sum(RunTime),2) as RunTime from Operations where Mirno ='" + mirno + "' and MachineName is null ";
            SqlCommand myCommand = new SqlCommand(CommandText, Conn);
            SqlDataAdapter myAdapter = new SqlDataAdapter();
            myAdapter.SelectCommand = myCommand;
            DataSet myDataSet = new DataSet();
            myAdapter.Fill(myDataSet);
            Conn.Close();
            return Ok(myDataSet.Tables[0]);
        }

        [HttpGet]
        [Route("totweight1")]
        public async Task<IActionResult> Totweight1(string mirno)
        {
            if (Conn.State == ConnectionState.Closed)
                Conn.Open();

            DateTime scandate = DateTime.Parse(DateTime.Now.ToString());
            string scan_Date = scandate.Month + "/" + scandate.Day + "/" + scandate.Year + " " + scandate.TimeOfDay;

            string CommandText = "select sum(TotalWt)/1000 as TotalWheight,count(RSNo) as RSno,sum(Tot_OPS) as TotalOpns,sum(RunTime) as RunTime from Operations where Mirno='" + mirno + "' and MachineName=(select TOP 1 MachineName from Operations where Mirno='" + mirno + "') and PlanningDate=(select TOP 1 PlanningDate from Operations where Mirno='" + mirno + "') and PlanningShift=(select TOP 1 PlanningShift from Operations where Mirno='" + mirno + "')";
            SqlCommand myCommand = new SqlCommand(CommandText, Conn);
            SqlDataAdapter myAdapter = new SqlDataAdapter();
            myAdapter.SelectCommand = myCommand;
            DataSet myDataSet = new DataSet();
            myAdapter.Fill(myDataSet);
            Conn.Close();
            return Ok(myDataSet.Tables[0]);
        }

        [HttpGet]
        [Route("onBalPunchMIR")]
        public async Task<IActionResult> OnBalPunchMIR(DateTime dateTimePicker1)
        {
            if (Conn.State == ConnectionState.Closed)
                Conn.Open();

            DateTime scandate = DateTime.Parse(dateTimePicker1.ToString());
            string scan_Date = scandate.Month + "/" + scandate.Day + "/" + scandate.Year + " " + scandate.TimeOfDay;

            string CommandText = " select round(sum(TotalWt)/1000,3) as TotalWheight,count(RSNo) as RSno,round(sum(Tot_OPS),2) as TotalOpns, round(sum(RunTime),2) as RunTime from Operations where PlanningDate='" + scan_Date + "' ";
            SqlCommand myCommand = new SqlCommand(CommandText, Conn);
            SqlDataAdapter myAdapter = new SqlDataAdapter();
            myAdapter.SelectCommand = myCommand;
            DataSet myDataSet = new DataSet();
            myAdapter.Fill(myDataSet);
            Conn.Close();
            return Ok(myDataSet.Tables[0]);
        }

        [HttpGet]
        [Route("bindDataGridOnPunchMIR")]
        public async Task<IActionResult> BindDataGridOnPunchMIR(string plantCode, DateTime dateTimePicker1)
        {
            if (Conn.State == ConnectionState.Closed)
                Conn.Open();
            DateTime scandate = DateTime.Parse(dateTimePicker1.ToString());
            string scan_Date = scandate.Month + "/" + scandate.Day + "/" + scandate.Year + " " + scandate.TimeOfDay;

            string CommandText = "SELECT RSNo as rsNo,FGItem as Item_No,Mirno as mirNo,MachineName as WorkCenter,SctDinemtion as SECTION,OPStatus,LotCode as billable_Lot,Pices as QTY,Length,Wheight as Wt_Pcs, FORMAT (JDDate, 'dd/MM/yyyy ')  as ReleasedDate,TotalWt,RackDetails,Operation as OPRPCs,Tot_OPS,PlanningShift, FORMAT (PlanningDate, 'dd/MM/yyyy ') as PlanningDate,RunTime,Status FROM Operations where PlanningDate='" + scan_Date + "' and BP='" + plantCode + "' and Flag_Fab is null order by [index] desc ";
            SqlCommand myCommand = new SqlCommand(CommandText, Conn);
            SqlDataAdapter myAdapter = new SqlDataAdapter();
            myAdapter.SelectCommand = myCommand;
            DataSet myDataSet = new DataSet();
            myAdapter.Fill(myDataSet);
            Conn.Close();
            return Ok(myDataSet.Tables[0]);
        }


        [HttpGet]
        [Route("onBalAllocateMIR")]
        public async Task<IActionResult> OnBalAllocateMIR()
        {
            if (Conn.State == ConnectionState.Closed)
                Conn.Open(); 
            string CommandText = " select round(sum(TotalWt)/1000,3) as TotalWheight,count(RSNo) as RSno,round(sum(Tot_OPS),2) as TotalOpns, round(sum(RunTime),2) as RunTime from Operations where PlanningDate is null ";
            SqlCommand myCommand = new SqlCommand(CommandText, Conn);
            SqlDataAdapter myAdapter = new SqlDataAdapter();
            myAdapter.SelectCommand = myCommand;
            DataSet myDataSet = new DataSet();
            myAdapter.Fill(myDataSet);
            Conn.Close();
            return Ok(myDataSet.Tables[0]);
        }

        [HttpGet]
        [Route("bindDataGridOnAllocateMIR")]
        public async Task<IActionResult> BindDataGridOnAllocateMIR(string plantCode)
        {
            if (Conn.State == ConnectionState.Closed)
                Conn.Open();
            string CommandText = " SELECT RSNo as rsNo,FGItem as Item_No,Mirno as mirNo,MachineName as WorkCenter,SctDinemtion as SECTION,OPStatus,LotCode as billable_Lot,Pices as QTY,Length,Wheight as Wt_Pcs, FORMAT (JDDate, 'dd/MM/yyyy ')  as ReleasedDate,TotalWt,RackDetails,Operation as OPRPCs,Tot_OPS,PlanningShift, FORMAT (PlanningDate, 'dd/MM/yyyy ') as PlanningDate,RunTime,Status FROM Operations where PlanningDate is null and BP='" + plantCode + "' and Flag_Fab is null order by [index] desc  ";
            SqlCommand myCommand = new SqlCommand(CommandText, Conn);
            SqlDataAdapter myAdapter = new SqlDataAdapter();
            myAdapter.SelectCommand = myCommand;
            DataSet myDataSet = new DataSet();
            myAdapter.Fill(myDataSet);
            Conn.Close();
            return Ok(myDataSet.Tables[0]);
        }

        [HttpGet]
        [Route("bindDataGridOnradioButton1")]
        public async Task<IActionResult> BindDataGridOnradioButton1(string plantCode)
        {
            if (Conn.State == ConnectionState.Closed)
                Conn.Open();
            string CommandText = " SELECT RSNo as rsNo,FGItem as Item_No,Mirno as mirNo,MachineName as WorkCenter,SctDinemtion as SECTION,OPStatus,LotCode as billable_Lot,Pices as QTY,Length,Wheight as Wt_Pcs, FORMAT (JDDate, 'dd/MM/yyyy ')  as ReleasedDate,TotalWt,RackDetails,Operation as OPRPCs,Tot_OPS,PlanningShift, FORMAT (PlanningDate, 'dd/MM/yyyy ') as PlanningDate,RunTime,Status FROM Operations where Flag_Galva is null and BP='" + plantCode + "' ";
            SqlCommand myCommand = new SqlCommand(CommandText, Conn);
            SqlDataAdapter myAdapter = new SqlDataAdapter();
            myAdapter.SelectCommand = myCommand;
            DataSet myDataSet = new DataSet();
            myAdapter.Fill(myDataSet);
            Conn.Close();
            return Ok(myDataSet.Tables[0]);
        }

        ////GET api/values/5
        //[HttpGet("{id}")]
        //public string Get(int id)
        //{
        //    return "value sammm";
        //}

        //// POST api/values
        //[HttpPost]
        //public void Post([FromBody] string value)
        //{
        //}

        //// PUT api/values/5
        //[HttpPut("{id}")]
        //public void Put(int id, [FromBody] string value)
        //{
        //}

        //// DELETE api/values/5
        //[HttpDelete("{id}")]
        //public void Delete(int id)
        //{
        //}
    }
}
