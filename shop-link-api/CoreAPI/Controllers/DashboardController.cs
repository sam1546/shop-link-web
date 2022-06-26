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

using Excel123 = Microsoft.Office.Interop.Excel;
using System.Reflection;

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
        string exc = "";

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

            exc = _env.ContentRootPath + "//files//xlsx//Book1.xlsx";


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
                        //c1 = new SqlCommand("select sum(TotalWt)/1000 as TotalWheight,count(RSNo) as RSno,sum(Tot_OPS) as TotalOpns,sum(RunTime) as RunTime from Operations where Mirno='" + txtMirno.Text + "' and BP='" + lblPlantCode + "' and Flag_Fab is null and POType='Primary'", cn);
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


        [HttpGet]
        [Route("shop1")]
        public async Task<IActionResult> Shop1(string plantCode)
        {
            string lblPlantCode = plantCode;
            string dateTimePicker1 = DateTime.Now.ToString();
            Excel123.ApplicationClass excelApp = new Excel123.ApplicationClass();
            Excel123.Workbook workbook = (Excel123.Workbook)excelApp.Workbooks.Add(Missing.Value);
            Excel123.Worksheet wrksheet;

            if (Conn.State == ConnectionState.Closed)
                Conn.Open();

            /*StreamReader sr = new StreamReader(Application.StartupPath + "\\" + "baudsetting.ini");
            string exc = sr.ReadLine();
            sr.Close();*/

            //string exc = Application.StartupPath + "\\" + "Book1.xlsx";

            if (lblPlantCode == "TM01")
                exc = _env.ContentRootPath + "//files//xlsx//TM01.xlsx";
            else if (lblPlantCode == "TM02")
                exc = _env.ContentRootPath + "//files//xlsx//TM02.xlsx";
            else if (lblPlantCode == "TM03")
                exc = _env.ContentRootPath + "//files//xlsx//TM03.xlsx";
            else if (lblPlantCode == "TMD1")
                exc = _env.ContentRootPath + "//files//xlsx//TMD1.xlsx";
            try
            {
                DateTime currentTime = DateTime.Now;
                workbook = excelApp.Workbooks.Open(exc, 0, false, 5, "", "", true, Excel123.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                //get 1st worksheet
                wrksheet = (Excel123.Worksheet)workbook.Sheets.get_Item(1);

                SqlCommand cmd = new SqlCommand();
                SqlDataReader dr;
                 
                int CK = 0, CL = 0, CM = 0;
                string queryadp = "";
                for (int i = 0; i < 3; i++)
                {
                    CK = 60 + 3 * i;
                    CL = 61 + 3 * i;
                    CM = 62 + 3 * i;
                    DateTime scandate = DateTime.Parse(dateTimePicker1.ToString());
                    scandate = scandate.AddDays(i);

                    if (lblPlantCode == "TM02")

                    {
                        string scan_Date = scandate.Month + "/" + scandate.Day + "/" + scandate.Year + " " + scandate.TimeOfDay;
                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";

                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string sec = "", lot = "", st = "", tot = "", Trsno = "", Toprns = "";
                        while (dr.Read())
                        {
                            if (!sec.Contains(dr["Mirno"].ToString()))
                                sec += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lot += dr["LotCode"].ToString().Remove(5);
                        }
                        dr.Close();

                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP01' and BP='" + lblPlantCode + "' and Flag_Fab is null";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();

                        while (dr.Read())
                        {
                            st += dr["OPStatus"].ToString();
                        }

                        dr.Close();
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            tot = dr["TWheight"].ToString();
                            Trsno = dr["TRSNo"].ToString();
                            Toprns = dr["TOPsn"].ToString();
                            lot = dr["LotCode"].ToString();
                        }
                        dr.Close();
                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP01' and BP='" + lblPlantCode + "' and Flag_Fab is null";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secC1 = "", lotC1 = "", mirC1 = "", stc1 = "", TrsnoC1 = "", ToprnsC1 = "";
                        while (dr.Read())
                        {
                            if (!secC1.Contains(dr["Mirno"].ToString()))
                                secC1 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotC1 += dr["LotCode"].ToString().Remove(5);
                            //mirC1+= dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();

                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();

                        while (dr.Read())
                        {

                            stc1 += dr["OPStatus"].ToString();


                        }
                        dr.Close();
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP01' and BP='" + lblPlantCode + "' and Flag_Fab is null";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirC1 = dr["TWheight"].ToString();
                            TrsnoC1 = dr["TRSNo"].ToString();
                            ToprnsC1 = dr["TOPsn"].ToString();
                            lotC1 = dr["LotCode"].ToString();
                        }
                        dr.Close();
                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD1 = "", lotD1 = "", mirD1 = "", std1 = "";
                        if (dr.Read())
                        {
                            if (!secD1.Contains(dr["Mirno"].ToString()))
                                secD1 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD1 += dr["LotCode"].ToString().Remove(5);
                            //lotD1 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();

                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();

                        while (dr.Read())
                        {

                            std1 += dr["OPStatus"].ToString();


                        }
                        dr.Close();
                        string TrsnoD1 = "", ToprnsD1 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD1 = dr["TWheight"].ToString();
                            TrsnoD1 = dr["TRSNo"].ToString();
                            ToprnsD1 = dr["TOPsn"].ToString();
                            lotD1 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb2 = "", lotb2 = "", mirb2 = "", stb2 = "";
                        while (dr.Read())
                        {
                            if (!secb2.Contains(dr["Mirno"].ToString()))
                                secb2 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            // lotb2 += dr["LotCode"].ToString().Remove(5);
                            //mirb2 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();

                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb2 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob2 = "", Toprnsb2 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb2 = dr["TWheight"].ToString();
                            Trsnob2 = dr["TRSNo"].ToString();
                            Toprnsb2 = dr["TOPsn"].ToString();
                            lotb2 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc2 = "", lotc2 = "", mirc2 = "", stc2 = "";
                        while (dr.Read())
                        {
                            if (!secc2.Contains(dr["Mirno"].ToString()))
                                secc2 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc2 += dr["LotCode"].ToString().Remove(5);
                            //mirc2 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc2 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc2 = "", Toprnsc2 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc2 = dr["TWheight"].ToString();
                            Trsnoc2 = dr["TRSNo"].ToString();
                            Toprnsc2 = dr["TOPsn"].ToString();
                            lotc2 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD2 = "", lotD2 = "", mirD2 = "", std2 = "";
                        while (dr.Read())
                        {
                            if (!secD2.Contains(dr["Mirno"].ToString()))
                                secD2 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD2 += dr["LotCode"].ToString().Remove(5);
                            // mirD2 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std2 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD2 = "", ToprnsD2 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD2 = dr["TWheight"].ToString();
                            TrsnoD2 = dr["TRSNo"].ToString();
                            ToprnsD2 = dr["TOPsn"].ToString();
                            lotD2 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb3 = "", lotb3 = "", mirb3 = "", stb3 = "";
                        while (dr.Read())
                        {
                            if (!secb3.Contains(dr["Mirno"].ToString()))
                                secb3 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb3 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob3 = "", Toprnsb3 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb3 = dr["TWheight"].ToString();
                            Trsnob3 = dr["TRSNo"].ToString();
                            Toprnsb3 = dr["TOPsn"].ToString();
                            lotb3 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc3 = "", lotc3 = "", mirc3 = "", stc3 = "";
                        while (dr.Read())
                        {
                            if (!secc3.Contains(dr["Mirno"].ToString()))
                                secc3 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc3 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc3 = "", Toprnsc3 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc3 = dr["TWheight"].ToString();
                            Trsnoc3 = dr["TRSNo"].ToString();
                            Toprnsc3 = dr["TOPsn"].ToString();
                            lotc3 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD3 = "", lotD3 = "", mirD3 = "", std3 = "";
                        while (dr.Read())
                        {
                            if (!secD3.Contains(dr["Mirno"].ToString()))
                                secD3 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std3 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD3 = "", ToprnsD3 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD3 = dr["TWheight"].ToString();
                            TrsnoD3 = dr["TRSNo"].ToString();
                            ToprnsD3 = dr["TOPsn"].ToString();
                            lotD3 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb4 = "", lotb4 = "", mirb4 = "", stb4 = "";
                        while (dr.Read())
                        {
                            if (!secb4.Contains(dr["Mirno"].ToString()))
                                secb4 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb4 += dr["LotCode"].ToString().Remove(5);
                            //mirb4 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            //if (!(stb4.Contains("B")||!stb4.Contains("N")||!stb4.Contains("M"))
                            stb4 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob4 = "", Toprnsb4 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb4 = dr["TWheight"].ToString();
                            Trsnob4 = dr["TRSNo"].ToString();
                            Toprnsb4 = dr["TOPsn"].ToString();
                            lotb4 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc4 = "", lotc4 = "", mirc4 = "", stc4 = "";
                        while (dr.Read())
                        {
                            if (!secc4.Contains(dr["Mirno"].ToString()))
                                secc4 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc4 += dr["LotCode"].ToString().Remove(5);
                            //mirc4 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc4 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc4 = "", Toprnsc4 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc4 = dr["TWheight"].ToString();
                            Trsnoc4 = dr["TRSNo"].ToString();
                            Toprnsc4 = dr["TOPsn"].ToString();
                            lotc4 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD4 = "", lotD4 = "", mirD4 = "", std4 = "";
                        while (dr.Read())
                        {
                            if (!secD4.Contains(dr["Mirno"].ToString()))
                                secD4 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD4 += dr["LotCode"].ToString().Remove(5);
                            // mirD4 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std4 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD4 = "", ToprnsD4 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD4 = dr["TWheight"].ToString();
                            TrsnoD4 = dr["TRSNo"].ToString();
                            ToprnsD4 = dr["TOPsn"].ToString();
                            lotD4 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP05' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb5 = "", lotb5 = "", mirb5 = "", stb5 = "";
                        while (dr.Read())
                        {
                            if (!secb5.Contains(dr["Mirno"].ToString()))
                                secb5 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb5 += dr["LotCode"].ToString().Remove(5);
                            // mirb5 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP05' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb5 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob5 = "", Toprnsb5 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP05' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb5 = dr["TWheight"].ToString();
                            Trsnob5 = dr["TRSNo"].ToString();
                            Toprnsb5 = dr["TOPsn"].ToString();
                            lotb5 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP05' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc5 = "", lotc5 = "", mirc5 = "", stc5 = "";
                        while (dr.Read())
                        {
                            if (!secc5.Contains(dr["Mirno"].ToString()))
                                secc5 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            // lotc5 += dr["LotCode"].ToString().Remove(5);
                            //mirc5 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP05' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc5 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc5 = "", Toprnsc5 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP05' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc5 = dr["TWheight"].ToString();
                            Trsnoc5 = dr["TRSNo"].ToString();
                            Toprnsc5 = dr["TOPsn"].ToString();
                            lotc5 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP05' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD5 = "", lotD5 = "", mirD5 = "", std5 = "";
                        while (dr.Read())
                        {
                            if (!secD5.Contains(dr["Mirno"].ToString()))
                                secD5 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD5 += dr["LotCode"].ToString().Remove(5);
                            //mirD5 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP05' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std5 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD5 = "", ToprnsD5 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP05' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD5 = dr["TWheight"].ToString();
                            TrsnoD5 = dr["TRSNo"].ToString();
                            ToprnsD5 = dr["TOPsn"].ToString();
                            lotD5 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP06' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb6 = "", lotb6 = "", mirb6 = "", stb6 = "";
                        while (dr.Read())
                        {
                            if (!secb6.Contains(dr["Mirno"].ToString()))
                                secb6 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb6 += dr["LotCode"].ToString().Remove(5);
                            //mirb6 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP06' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb6 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob6 = "", Toprnsb6 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP06' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb6 = dr["TWheight"].ToString();
                            Trsnob6 = dr["TRSNo"].ToString();
                            Toprnsb6 = dr["TOPsn"].ToString();
                            lotb6 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP06' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc6 = "", lotc6 = "", mirc6 = "", stc6 = "";
                        while (dr.Read())
                        {
                            if (!secc6.Contains(dr["Mirno"].ToString()))
                                secc6 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc6 += dr["LotCode"].ToString().Remove(5);
                            //mirc6 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP06' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc6 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc6 = "", Toprnsc6 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP06' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc6 = dr["TWheight"].ToString();
                            Trsnoc6 = dr["TRSNo"].ToString();
                            Toprnsc6 = dr["TOPsn"].ToString();
                            lotc6 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP06' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD6 = "", lotD6 = "", mirD6 = "", stD6 = "";
                        while (dr.Read())
                        {
                            if (!secD6.Contains(dr["Mirno"].ToString()))
                                secD6 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD6 += dr["LotCode"].ToString().Remove(5);
                            //mirD6 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP06' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stD6 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD6 = "", ToprnsD6 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP06' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD6 = dr["TWheight"].ToString();
                            TrsnoD6 = dr["TRSNo"].ToString();
                            ToprnsD6 = dr["TOPsn"].ToString();
                            lotD6 = dr["LotCode"].ToString();
                        }
                        dr.Close();



                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='P-81' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb7 = "", lotb7 = "", mirb7 = "";
                        while (dr.Read())
                        {
                            if (!secb7.Contains(dr["Mirno"].ToString()))
                                secb7 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb7 += dr["LotCode"].ToString().Remove(5);
                            //mirb7 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='P-81' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc7 = "", lotc7 = "", mirc7 = "";
                        while (dr.Read())
                        {
                            if (!secc7.Contains(dr["Mirno"].ToString()))
                                secc7 += (dr["SctDinemtion"].ToString().Remove(1, 10) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc7 += dr["LotCode"].ToString().Remove(5);
                            //mirc7 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='P-81' and BP='" + lblPlantCode + "' and Flag_Fab is null  ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD7 = "", lotD7 = "", mirD7 = "";
                        while (dr.Read())
                        {
                            if (!secD7.Contains(dr["Mirno"].ToString()))
                                secD7 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD7 += dr["LotCode"].ToString().Remove(5);
                            //mirD7 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();

                        //*************************************CD01 to CP07*******************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP07' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb8 = "", lotb8 = "", mirb8 = "", stb8 = "";
                        while (dr.Read())
                        {
                            if (!secb8.Contains(dr["Mirno"].ToString()))
                                secb8 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb8 += dr["LotCode"].ToString().Remove(5);
                            // mirb8 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP07' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb8 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob8 = "", Toprnsb8 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP07' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb8 = dr["TWheight"].ToString();
                            Trsnob8 = dr["TRSNo"].ToString();
                            Toprnsb8 = dr["TOPsn"].ToString();
                            lotb8 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP07' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc8 = "", lotc8 = "", mirc8 = "", stc8 = "";
                        if (dr.Read())
                        {
                            if (!secc8.Contains(dr["Mirno"].ToString()))
                                secc8 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc8 += dr["LotCode"].ToString().Remove(5);
                            //mirc7 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP07' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc8 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc8 = "", Toprnsc8 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP07' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc8 = dr["TWheight"].ToString();
                            Trsnoc8 = dr["TRSNo"].ToString();
                            Toprnsc8 = dr["TOPsn"].ToString();
                            lotc8 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP07' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD8 = "", lotD8 = "", mirD8 = "", std8 = "";
                        while (dr.Read())
                        {
                            if (!secD8.Contains(dr["Mirno"].ToString()))
                                secD8 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD8 += dr["LotCode"].ToString().Remove(5);
                            //mirD8 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP07' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std8 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD8 = "", ToprnsD8 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP07' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD8 = dr["TWheight"].ToString();
                            TrsnoD8 = dr["TRSNo"].ToString();
                            ToprnsD8 = dr["TOPsn"].ToString();
                            lotD8 = dr["LotCode"].ToString();

                        }
                        dr.Close();

                        //**************************************  cp07 to CD01  ****************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CD01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb9 = "", lotb9 = "", mirb9 = "", stb9 = "";
                        while (dr.Read())
                        {
                            if (!secb9.Contains(dr["Mirno"].ToString()))
                                secb9 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb6 += dr["LotCode"].ToString().Remove(5);
                            //mirb6 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CD01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb9 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob9 = "", Toprnsb9 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CD01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb9 = dr["TWheight"].ToString();
                            Trsnob9 = dr["TRSNo"].ToString();
                            Toprnsb9 = dr["TOPsn"].ToString();
                            lotb9 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CD01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc9 = "", lotc9 = "", mirc9 = "", stc9 = "";
                        while (dr.Read())
                        {
                            if (!secc9.Contains(dr["Mirno"].ToString()))
                                secc9 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc6 += dr["LotCode"].ToString().Remove(5);
                            //mirc6 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CD01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc9 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc9 = "", Toprnsc9 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CD01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc9 = dr["TWheight"].ToString();
                            Trsnoc9 = dr["TRSNo"].ToString();
                            Toprnsc9 = dr["TOPsn"].ToString();
                            lotc9 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CD01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD9 = "", lotD9 = "", mirD9 = "", stD9 = "";
                        while (dr.Read())
                        {
                            if (!secD9.Contains(dr["Mirno"].ToString()))
                                secD9 += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD6 += dr["LotCode"].ToString().Remove(5);
                            //mirD6 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CD01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stD9 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD9 = "", ToprnsD9 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CD01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD9 = dr["TWheight"].ToString();
                            TrsnoD9 = dr["TRSNo"].ToString();
                            ToprnsD9 = dr["TOPsn"].ToString();
                            lotD9 = dr["LotCode"].ToString();
                        }
                        dr.Close();







                        //cn.Close();
                        ((Excel123.Range)wrksheet.Cells["1", "BI"]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["2", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["4", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["4", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["4", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["5", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["5", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["5", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["6", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["6", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["6", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["7", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["7", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["7", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["8", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["8", CL]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["8", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["9", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["9", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["9", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["10", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["10", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["10", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["11", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["11", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["11", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["12", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["12", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["12", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["13", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["13", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["13", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["14", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["14", CL]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["14", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["15", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["15", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["15", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["16", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["16", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["16", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["17", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["17", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["17", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["18", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["18", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["18", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["19", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["19", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["19", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["20", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["20", CL]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["20", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["21", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["21", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["21", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["22", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["22", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["22", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["23", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["23", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["23", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["24", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["24", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["24", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["25", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["25", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["25", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["26", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["26", CL]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["26", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["27", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["27", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["27", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["28", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["28", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["28", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["29", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["29", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["29", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["30", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["30", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["30", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["31", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["31", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["31", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["32", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["32", CL]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["32", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["33", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["33", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["33", CM]).Value2 = "";


                        ((Excel123.Range)wrksheet.Cells["34", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["34", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["34", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["35", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["35", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["35", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["36", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["36", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["36", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["37", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["37", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["37", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["38", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["38", CL]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["38", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["39", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["39", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["39", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["40", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["40", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["40", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["41", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["41", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["41", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["42", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["42", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["42", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["43", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["43", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["43", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["44", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["44", CL]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["44", CM]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["45", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["45", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["45", CM]).Value2 = "";

                        //***********************CP07*****************************

                        ((Excel123.Range)wrksheet.Cells["46", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["46", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["46", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["47", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["47", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["47", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["48", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["48", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["48", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["49", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["49", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["49", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["50", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["50", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["50", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["51", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["51", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["51", CM]).Value2 = "";



                        ((Excel123.Range)wrksheet.Cells["5", CK]).Value2 = tot;
                        ((Excel123.Range)wrksheet.Cells["5", CL]).Value2 = mirC1;
                        ((Excel123.Range)wrksheet.Cells["5", CM]).Value2 = mirD1;
                        ((Excel123.Range)wrksheet.Cells["11", CK]).Value2 = mirb2;
                        ((Excel123.Range)wrksheet.Cells["11", CL]).Value2 = mirc2;
                        ((Excel123.Range)wrksheet.Cells["11", CM]).Value2 = mirD2;

                        ((Excel123.Range)wrksheet.Cells["17", CK]).Value2 = mirb3;
                        ((Excel123.Range)wrksheet.Cells["17", CL]).Value2 = mirc3;
                        ((Excel123.Range)wrksheet.Cells["17", CM]).Value2 = mirD3;
                        ((Excel123.Range)wrksheet.Cells["23", CK]).Value2 = mirb4;
                        ((Excel123.Range)wrksheet.Cells["23", CL]).Value2 = mirc4;
                        ((Excel123.Range)wrksheet.Cells["23", CM]).Value2 = mirD4;

                        ((Excel123.Range)wrksheet.Cells["29", CK]).Value2 = mirb5;
                        ((Excel123.Range)wrksheet.Cells["29", CL]).Value2 = mirc5;
                        ((Excel123.Range)wrksheet.Cells["29", CM]).Value2 = mirD5;
                        ((Excel123.Range)wrksheet.Cells["35", CK]).Value2 = mirb6;
                        ((Excel123.Range)wrksheet.Cells["35", CL]).Value2 = mirc6;
                        ((Excel123.Range)wrksheet.Cells["35", CM]).Value2 = mirD6;


                        ((Excel123.Range)wrksheet.Cells["1", "BI"]).Value2 = currentTime;
                        ((Excel123.Range)wrksheet.Cells["2", CL]).Value2 = scandate;
                        ((Excel123.Range)wrksheet.Cells["4", CK]).Value2 = sec;
                        ((Excel123.Range)wrksheet.Cells["4", CL]).Value2 = secC1;

                        ((Excel123.Range)wrksheet.Cells["4", CM]).Value2 = secD1;
                        // ((Excel123.Range)wrksheet.Cells["5", CK]).Value2 = tot;

                        ((Excel123.Range)wrksheet.Cells["10", CK]).Value2 = secb2;
                        ((Excel123.Range)wrksheet.Cells["10", CL]).Value2 = secc2;
                        ((Excel123.Range)wrksheet.Cells["10", CM]).Value2 = secD2;
                        ((Excel123.Range)wrksheet.Cells["16", CK]).Value2 = secb3;
                        ((Excel123.Range)wrksheet.Cells["16", CL]).Value2 = secc3;
                        ((Excel123.Range)wrksheet.Cells["16", CM]).Value2 = secD3;
                        ((Excel123.Range)wrksheet.Cells["22", CK]).Value2 = secb4;
                        ((Excel123.Range)wrksheet.Cells["22", CL]).Value2 = secc4;
                        ((Excel123.Range)wrksheet.Cells["22", CM]).Value2 = secD4;
                        ((Excel123.Range)wrksheet.Cells["28", CK]).Value2 = secb5;
                        ((Excel123.Range)wrksheet.Cells["28", CL]).Value2 = secc5;

                        ((Excel123.Range)wrksheet.Cells["28", CM]).Value2 = secD5;
                        ((Excel123.Range)wrksheet.Cells["34", CK]).Value2 = secb6;
                        ((Excel123.Range)wrksheet.Cells["34", CL]).Value2 = secc6;
                        ((Excel123.Range)wrksheet.Cells["34", CM]).Value2 = secD6;

                        //***************************

                        ((Excel123.Range)wrksheet.Cells["9", CK]).Value2 = lot;
                        ((Excel123.Range)wrksheet.Cells["9", CL]).Value2 = lotC1;
                        ((Excel123.Range)wrksheet.Cells["9", CM]).Value2 = lotD1;
                        ((Excel123.Range)wrksheet.Cells["15", CK]).Value2 = lotb2;
                        ((Excel123.Range)wrksheet.Cells["15", CL]).Value2 = lotc2;
                        ((Excel123.Range)wrksheet.Cells["15", CM]).Value2 = lotD2;
                        ((Excel123.Range)wrksheet.Cells["21", CK]).Value2 = lotb3;
                        ((Excel123.Range)wrksheet.Cells["21", CL]).Value2 = lotc3;
                        ((Excel123.Range)wrksheet.Cells["21", CM]).Value2 = lotD3;
                        ((Excel123.Range)wrksheet.Cells["27", CK]).Value2 = lotb4;
                        ((Excel123.Range)wrksheet.Cells["27", CL]).Value2 = lotc4;
                        ((Excel123.Range)wrksheet.Cells["27", CM]).Value2 = lotD4;
                        ((Excel123.Range)wrksheet.Cells["33", CK]).Value2 = lotb5;
                        ((Excel123.Range)wrksheet.Cells["33", CL]).Value2 = lotc5;

                        ((Excel123.Range)wrksheet.Cells["33", CM]).Value2 = lotD5;
                        ((Excel123.Range)wrksheet.Cells["39", CK]).Value2 = lotb6;
                        ((Excel123.Range)wrksheet.Cells["39", CL]).Value2 = lotc6;
                        ((Excel123.Range)wrksheet.Cells["39", CM]).Value2 = lotD6;



                        //*************************

                        ((Excel123.Range)wrksheet.Cells["8", CK]).Value2 = st;
                        ((Excel123.Range)wrksheet.Cells["8", CL]).Value2 = stc1;
                        ((Excel123.Range)wrksheet.Cells["8", CM]).Value2 = std1;
                        ((Excel123.Range)wrksheet.Cells["14", CK]).Value2 = stb2;
                        ((Excel123.Range)wrksheet.Cells["14", CL]).Value2 = stc2;
                        ((Excel123.Range)wrksheet.Cells["14", CM]).Value2 = std2;

                        ((Excel123.Range)wrksheet.Cells["20", CK]).Value2 = stb3;
                        ((Excel123.Range)wrksheet.Cells["20", CL]).Value2 = stc3;
                        ((Excel123.Range)wrksheet.Cells["20", CM]).Value2 = std3;
                        ((Excel123.Range)wrksheet.Cells["26", CK]).Value2 = stb4;
                        ((Excel123.Range)wrksheet.Cells["26", CL]).Value2 = stc4;
                        ((Excel123.Range)wrksheet.Cells["26", CM]).Value2 = std4;

                        ((Excel123.Range)wrksheet.Cells["32", CK]).Value2 = stb5;
                        ((Excel123.Range)wrksheet.Cells["32", CL]).Value2 = stc5;
                        ((Excel123.Range)wrksheet.Cells["32", CM]).Value2 = std5;
                        ((Excel123.Range)wrksheet.Cells["38", CK]).Value2 = stb6;
                        ((Excel123.Range)wrksheet.Cells["38", CL]).Value2 = stc6;
                        ((Excel123.Range)wrksheet.Cells["38", CM]).Value2 = stD6;


                        ((Excel123.Range)wrksheet.Cells["6", CK]).Value2 = Trsno;
                        ((Excel123.Range)wrksheet.Cells["6", CL]).Value2 = TrsnoC1;
                        ((Excel123.Range)wrksheet.Cells["6", CM]).Value2 = TrsnoD1;
                        ((Excel123.Range)wrksheet.Cells["12", CK]).Value2 = Trsnob2;
                        ((Excel123.Range)wrksheet.Cells["12", CL]).Value2 = Trsnoc2;
                        ((Excel123.Range)wrksheet.Cells["12", CM]).Value2 = TrsnoD2;

                        ((Excel123.Range)wrksheet.Cells["18", CK]).Value2 = Trsnob3;
                        ((Excel123.Range)wrksheet.Cells["18", CL]).Value2 = Trsnoc3;
                        ((Excel123.Range)wrksheet.Cells["18", CM]).Value2 = TrsnoD3;
                        ((Excel123.Range)wrksheet.Cells["24", CK]).Value2 = Trsnob4;
                        ((Excel123.Range)wrksheet.Cells["24", CL]).Value2 = Trsnoc4;
                        ((Excel123.Range)wrksheet.Cells["24", CM]).Value2 = TrsnoD4;

                        ((Excel123.Range)wrksheet.Cells["30", CK]).Value2 = Trsnob5;
                        ((Excel123.Range)wrksheet.Cells["30", CL]).Value2 = Trsnoc5;
                        ((Excel123.Range)wrksheet.Cells["30", CM]).Value2 = TrsnoD5;
                        ((Excel123.Range)wrksheet.Cells["36", CK]).Value2 = Trsnob6;
                        ((Excel123.Range)wrksheet.Cells["36", CL]).Value2 = Trsnoc6;
                        ((Excel123.Range)wrksheet.Cells["36", CM]).Value2 = TrsnoD6;


                        ((Excel123.Range)wrksheet.Cells["7", CK]).Value2 = Toprns;
                        ((Excel123.Range)wrksheet.Cells["7", CL]).Value2 = ToprnsC1;
                        ((Excel123.Range)wrksheet.Cells["7", CM]).Value2 = ToprnsD1;
                        ((Excel123.Range)wrksheet.Cells["13", CK]).Value2 = Toprnsb2;
                        ((Excel123.Range)wrksheet.Cells["13", CL]).Value2 = Toprnsc2;
                        ((Excel123.Range)wrksheet.Cells["13", CM]).Value2 = ToprnsD2;

                        ((Excel123.Range)wrksheet.Cells["19", CK]).Value2 = Toprnsb3;
                        ((Excel123.Range)wrksheet.Cells["19", CL]).Value2 = Toprnsc3;
                        ((Excel123.Range)wrksheet.Cells["19", CM]).Value2 = ToprnsD3;
                        ((Excel123.Range)wrksheet.Cells["25", CK]).Value2 = Toprnsb4;
                        ((Excel123.Range)wrksheet.Cells["25", CL]).Value2 = Toprnsc4;
                        ((Excel123.Range)wrksheet.Cells["25", CM]).Value2 = ToprnsD4;

                        ((Excel123.Range)wrksheet.Cells["31", CK]).Value2 = Toprnsb5;
                        ((Excel123.Range)wrksheet.Cells["31", CL]).Value2 = Toprnsc5;
                        ((Excel123.Range)wrksheet.Cells["31", CM]).Value2 = ToprnsD5;
                        ((Excel123.Range)wrksheet.Cells["37", CK]).Value2 = Toprnsb6;
                        ((Excel123.Range)wrksheet.Cells["37", CL]).Value2 = Toprnsc6;
                        ((Excel123.Range)wrksheet.Cells["37", CM]).Value2 = ToprnsD6;






                        //**********************CP07*************************

                        ((Excel123.Range)wrksheet.Cells["40", CK]).Value2 = secb9;
                        ((Excel123.Range)wrksheet.Cells["40", CL]).Value2 = secc9;
                        ((Excel123.Range)wrksheet.Cells["40", CM]).Value2 = secD9;

                        ((Excel123.Range)wrksheet.Cells["41", CK]).Value2 = mirb9;
                        ((Excel123.Range)wrksheet.Cells["41", CL]).Value2 = mirc9;
                        ((Excel123.Range)wrksheet.Cells["41", CM]).Value2 = mirD9;

                        ((Excel123.Range)wrksheet.Cells["42", CK]).Value2 = Trsnob9;
                        ((Excel123.Range)wrksheet.Cells["42", CL]).Value2 = Trsnoc9;
                        ((Excel123.Range)wrksheet.Cells["42", CM]).Value2 = TrsnoD9;

                        ((Excel123.Range)wrksheet.Cells["43", CK]).Value2 = Toprnsb9;
                        ((Excel123.Range)wrksheet.Cells["43", CL]).Value2 = Toprnsc9;
                        ((Excel123.Range)wrksheet.Cells["43", CM]).Value2 = ToprnsD9;

                        ((Excel123.Range)wrksheet.Cells["44", CK]).Value2 = stb9;
                        ((Excel123.Range)wrksheet.Cells["44", CL]).Value2 = stc9;
                        ((Excel123.Range)wrksheet.Cells["44", CM]).Value2 = stD9;


                        ((Excel123.Range)wrksheet.Cells["45", CK]).Value2 = lotb9;
                        ((Excel123.Range)wrksheet.Cells["45", CL]).Value2 = lotc9;
                        ((Excel123.Range)wrksheet.Cells["45", CM]).Value2 = lotD9;




                        //**********************CD-01***************************
                        ((Excel123.Range)wrksheet.Cells["46", CK]).Value2 = secb8;
                        ((Excel123.Range)wrksheet.Cells["46", CL]).Value2 = secc8;
                        ((Excel123.Range)wrksheet.Cells["46", CM]).Value2 = secD8;

                        ((Excel123.Range)wrksheet.Cells["47", CK]).Value2 = mirb8;
                        ((Excel123.Range)wrksheet.Cells["47", CL]).Value2 = mirc8;
                        ((Excel123.Range)wrksheet.Cells["47", CM]).Value2 = mirD8;

                        ((Excel123.Range)wrksheet.Cells["48", CK]).Value2 = Trsnob8;
                        ((Excel123.Range)wrksheet.Cells["48", CL]).Value2 = Trsnoc8;
                        ((Excel123.Range)wrksheet.Cells["48", CM]).Value2 = TrsnoD8;

                        ((Excel123.Range)wrksheet.Cells["49", CK]).Value2 = Toprnsb8;
                        ((Excel123.Range)wrksheet.Cells["49", CL]).Value2 = Toprnsc8;
                        ((Excel123.Range)wrksheet.Cells["49", CM]).Value2 = ToprnsD8;

                        ((Excel123.Range)wrksheet.Cells["50", CK]).Value2 = stb8;
                        ((Excel123.Range)wrksheet.Cells["50", CL]).Value2 = stc8;
                        ((Excel123.Range)wrksheet.Cells["50", CM]).Value2 = std8;

                        ((Excel123.Range)wrksheet.Cells["51", CK]).Value2 = lotb8;
                        ((Excel123.Range)wrksheet.Cells["51", CL]).Value2 = lotc8;
                        ((Excel123.Range)wrksheet.Cells["51", CM]).Value2 = lotD8;





                    }

                    else if (lblPlantCode == "TM03")
                    {
                        string scan_Date = scandate.Month + "/" + scandate.Day + "/" + scandate.Year + " " + scandate.TimeOfDay;

                        //*************************************CP22***************************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CD22' and BP='" + lblPlantCode + "' and Flag_Fab is null ";

                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string sec = "", lot = "", st = "", tot = "", Trsno = "", Toprns = "";
                        while (dr.Read())
                        {
                            if (!sec.Contains(dr["Mirno"].ToString()))
                                //sec += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                                sec += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";

                            //lot += dr["LotCode"].ToString().Remove(5);
                        }
                        dr.Close();

                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CD22' and BP='" + lblPlantCode + "' and Flag_Fab is null";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();

                        while (dr.Read())
                        {
                            st += dr["OPStatus"].ToString();
                        }

                        dr.Close();
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CD22' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            tot = dr["TWheight"].ToString();
                            Trsno = dr["TRSNo"].ToString();
                            Toprns = dr["TOPsn"].ToString();
                            lot = dr["LotCode"].ToString();
                        }
                        dr.Close();
                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CD22' and BP='" + lblPlantCode + "' and Flag_Fab is null";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secC1 = "", lotC1 = "", mirC1 = "", stc1 = "", TrsnoC1 = "", ToprnsC1 = "";
                        while (dr.Read())
                        {
                            if (!secC1.Contains(dr["Mirno"].ToString()))
                                secC1 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotC1 += dr["LotCode"].ToString().Remove(5);
                            //mirC1+= dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();

                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CD22' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();

                        while (dr.Read())
                        {

                            stc1 += dr["OPStatus"].ToString();


                        }
                        dr.Close();
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CD22' and BP='" + lblPlantCode + "' and Flag_Fab is null";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirC1 = dr["TWheight"].ToString();
                            TrsnoC1 = dr["TRSNo"].ToString();
                            ToprnsC1 = dr["TOPsn"].ToString();
                            lotC1 = dr["LotCode"].ToString();
                        }
                        dr.Close();
                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CD22' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD1 = "", lotD1 = "", mirD1 = "", std1 = "";
                        if (dr.Read())
                        {
                            if (!secD1.Contains(dr["Mirno"].ToString()))
                                secD1 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD1 += dr["LotCode"].ToString().Remove(5);
                            //lotD1 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();

                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CD22' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();

                        while (dr.Read())
                        {

                            std1 += dr["OPStatus"].ToString();


                        }
                        dr.Close();
                        string TrsnoD1 = "", ToprnsD1 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CD22' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD1 = dr["TWheight"].ToString();
                            TrsnoD1 = dr["TRSNo"].ToString();
                            ToprnsD1 = dr["TOPsn"].ToString();
                            lotD1 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        //***********************************CD14************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CD14' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb2 = "", lotb2 = "", mirb2 = "", stb2 = "";
                        while (dr.Read())
                        {
                            if (!secb2.Contains(dr["Mirno"].ToString()))
                                secb2 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            // lotb2 += dr["LotCode"].ToString().Remove(5);
                            //mirb2 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();

                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CD14' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb2 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob2 = "", Toprnsb2 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CD14' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb2 = dr["TWheight"].ToString();
                            Trsnob2 = dr["TRSNo"].ToString();
                            Toprnsb2 = dr["TOPsn"].ToString();
                            lotb2 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CD14' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc2 = "", lotc2 = "", mirc2 = "", stc2 = "";
                        while (dr.Read())
                        {
                            if (!secc2.Contains(dr["Mirno"].ToString()))
                                secc2 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc2 += dr["LotCode"].ToString().Remove(5);
                            //mirc2 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CD14' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc2 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc2 = "", Toprnsc2 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CD14' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc2 = dr["TWheight"].ToString();
                            Trsnoc2 = dr["TRSNo"].ToString();
                            Toprnsc2 = dr["TOPsn"].ToString();
                            lotc2 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CD14' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD2 = "", lotD2 = "", mirD2 = "", std2 = "";
                        while (dr.Read())
                        {
                            if (!secD2.Contains(dr["Mirno"].ToString()))
                                secD2 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD2 += dr["LotCode"].ToString().Remove(5);
                            // mirD2 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CD14' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std2 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD2 = "", ToprnsD2 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CD14' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD2 = dr["TWheight"].ToString();
                            TrsnoD2 = dr["TRSNo"].ToString();
                            ToprnsD2 = dr["TOPsn"].ToString();
                            lotD2 = dr["LotCode"].ToString();
                        }
                        dr.Close();


                        //*********************************** CP02 ***********************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb3 = "", lotb3 = "", mirb3 = "", stb3 = "";
                        while (dr.Read())
                        {
                            if (!secb3.Contains(dr["Mirno"].ToString()))
                                secb3 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb3 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob3 = "", Toprnsb3 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb3 = dr["TWheight"].ToString();
                            Trsnob3 = dr["TRSNo"].ToString();
                            Toprnsb3 = dr["TOPsn"].ToString();
                            lotb3 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc3 = "", lotc3 = "", mirc3 = "", stc3 = "";
                        while (dr.Read())
                        {
                            if (!secc3.Contains(dr["Mirno"].ToString()))
                                secc3 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc3 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc3 = "", Toprnsc3 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc3 = dr["TWheight"].ToString();
                            Trsnoc3 = dr["TRSNo"].ToString();
                            Toprnsc3 = dr["TOPsn"].ToString();
                            lotc3 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD3 = "", lotD3 = "", mirD3 = "", std3 = "";
                        while (dr.Read())
                        {
                            if (!secD3.Contains(dr["Mirno"].ToString()))
                                secD3 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std3 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD3 = "", ToprnsD3 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD3 = dr["TWheight"].ToString();
                            TrsnoD3 = dr["TRSNo"].ToString();
                            ToprnsD3 = dr["TOPsn"].ToString();
                            lotD3 = dr["LotCode"].ToString();
                        }
                        dr.Close();


                        //*************************************CP-03*************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb4 = "", lotb4 = "", mirb4 = "", stb4 = "";
                        while (dr.Read())
                        {
                            if (!secb4.Contains(dr["Mirno"].ToString()))
                                secb4 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb4 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob4 = "", Toprnsb4 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb4 = dr["TWheight"].ToString();
                            Trsnob4 = dr["TRSNo"].ToString();
                            Toprnsb4 = dr["TOPsn"].ToString();
                            lotb4 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc4 = "", lotc4 = "", mirc4 = "", stc4 = "";
                        while (dr.Read())
                        {
                            if (!secc4.Contains(dr["Mirno"].ToString()))
                                secc4 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc4 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc4 = "", Toprnsc4 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc4 = dr["TWheight"].ToString();
                            Trsnoc4 = dr["TRSNo"].ToString();
                            Toprnsc4 = dr["TOPsn"].ToString();
                            lotc4 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD4 = "", lotD4 = "", mirD4 = "", std4 = "";
                        while (dr.Read())
                        {
                            if (!secD4.Contains(dr["Mirno"].ToString()))
                                secD4 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std4 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD4 = "", ToprnsD4 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD4 = dr["TWheight"].ToString();
                            TrsnoD4 = dr["TRSNo"].ToString();
                            ToprnsD4 = dr["TOPsn"].ToString();
                            lotD4 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        //**********************************************CP-04****************************************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb5 = "", lotb5 = "", mirb5 = "", stb5 = "";
                        while (dr.Read())
                        {
                            if (!secb5.Contains(dr["Mirno"].ToString()))
                                secb5 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb5 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob5 = "", Toprnsb5 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb5 = dr["TWheight"].ToString();
                            Trsnob5 = dr["TRSNo"].ToString();
                            Toprnsb5 = dr["TOPsn"].ToString();
                            lotb5 = dr["LotCode"].ToString();
                        }

                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc5 = "", lotc5 = "", mirc5 = "", stc5 = "";
                        while (dr.Read())
                        {

                            if (!secc5.Contains(dr["Mirno"].ToString()))
                                secc5 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc5 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc5 = "", Toprnsc5 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc5 = dr["TWheight"].ToString();
                            Trsnoc5 = dr["TRSNo"].ToString();
                            Toprnsc5 = dr["TOPsn"].ToString();
                            lotc5 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD5 = "", lotD5 = "", mirD5 = "", std5 = "";
                        while (dr.Read())
                        {
                            if (!secD5.Contains(dr["Mirno"].ToString()))
                                secD5 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std5 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD5 = "", ToprnsD5 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD5 = dr["TWheight"].ToString();
                            TrsnoD5 = dr["TRSNo"].ToString();
                            ToprnsD5 = dr["TOPsn"].ToString();
                            lotD5 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        //************************************** CP-08   ***********************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP08' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb6 = "", lotb6 = "", mirb6 = "", stb6 = "";
                        while (dr.Read())
                        {
                            if (!secb6.Contains(dr["Mirno"].ToString()))
                                secb6 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP08' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb6 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob6 = "", Toprnsb6 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP08' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb6 = dr["TWheight"].ToString();
                            Trsnob6 = dr["TRSNo"].ToString();
                            Toprnsb6 = dr["TOPsn"].ToString();
                            lotb6 = dr["LotCode"].ToString();
                        }

                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP08' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc6 = "", lotc6 = "", mirc6 = "", stc6 = "";
                        while (dr.Read())
                        {

                            if (!secc6.Contains(dr["Mirno"].ToString()))
                                secc6 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP08' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc6 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc6 = "", Toprnsc6 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP08' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc6 = dr["TWheight"].ToString();
                            Trsnoc6 = dr["TRSNo"].ToString();
                            Toprnsc6 = dr["TOPsn"].ToString();
                            lotc6 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP08' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD6 = "", lotD6 = "", mirD6 = "", std6 = "";
                        while (dr.Read())
                        {
                            if (!secD6.Contains(dr["Mirno"].ToString()))
                                secD6 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP08' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std6 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD6 = "", ToprnsD6 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP08' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD6 = dr["TWheight"].ToString();
                            TrsnoD6 = dr["TRSNo"].ToString();
                            ToprnsD6 = dr["TOPsn"].ToString();
                            lotD6 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        //**************************************************CP-09*******************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP09' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb7 = "", lotb7 = "", mirb7 = "", stb7 = "";
                        while (dr.Read())
                        {
                            if (!secb7.Contains(dr["Mirno"].ToString()))
                                secb7 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP09' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb7 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob7 = "", Toprnsb7 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP09' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb7 = dr["TWheight"].ToString();
                            Trsnob7 = dr["TRSNo"].ToString();
                            Toprnsb7 = dr["TOPsn"].ToString();
                            lotb7 = dr["LotCode"].ToString();
                        }

                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP09' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc7 = "", lotc7 = "", mirc7 = "", stc7 = "";
                        while (dr.Read())
                        {

                            if (!secc7.Contains(dr["Mirno"].ToString()))
                                secc7 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP09' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc7 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc7 = "", Toprnsc7 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP09' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc7 = dr["TWheight"].ToString();
                            Trsnoc7 = dr["TRSNo"].ToString();
                            Toprnsc7 = dr["TOPsn"].ToString();
                            lotc7 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP09' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD7 = "", lotD7 = "", mirD7 = "", std7 = "";
                        while (dr.Read())
                        {
                            if (!secD7.Contains(dr["Mirno"].ToString()))
                                secD7 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP09' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std7 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD7 = "", ToprnsD7 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP09' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD7 = dr["TWheight"].ToString();
                            TrsnoD7 = dr["TRSNo"].ToString();
                            ToprnsD7 = dr["TOPsn"].ToString();
                            lotD7 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        //****************************CP-10*****************************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP10' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb8 = "", lotb8 = "", mirb8 = "", stb8 = "";
                        while (dr.Read())
                        {
                            if (!secb8.Contains(dr["Mirno"].ToString()))
                                secb8 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP11' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb8 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob8 = "", Toprnsb8 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP11' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb8 = dr["TWheight"].ToString();
                            Trsnob8 = dr["TRSNo"].ToString();
                            Toprnsb8 = dr["TOPsn"].ToString();
                            lotb8 = dr["LotCode"].ToString();
                        }

                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP11' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc8 = "", lotc8 = "", mirc8 = "", stc8 = "";
                        while (dr.Read())
                        {

                            if (!secc8.Contains(dr["Mirno"].ToString()))
                                secc8 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP11' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc8 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc8 = "", Toprnsc8 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP11' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc8 = dr["TWheight"].ToString();
                            Trsnoc8 = dr["TRSNo"].ToString();
                            Toprnsc8 = dr["TOPsn"].ToString();
                            lotc8 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP11' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD8 = "", lotD8 = "", mirD8 = "", std8 = "";
                        while (dr.Read())
                        {
                            if (!secD8.Contains(dr["Mirno"].ToString()))
                                secD8 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP11' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std8 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD8 = "", ToprnsD8 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP11' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD8 = dr["TWheight"].ToString();
                            TrsnoD8 = dr["TRSNo"].ToString();
                            ToprnsD8 = dr["TOPsn"].ToString();
                            lotD8 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        //****************************CP-11*****************************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP11' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb9 = "", lotb9 = "", mirb9 = "", stb9 = "";
                        while (dr.Read())
                        {
                            if (!secb9.Contains(dr["Mirno"].ToString()))
                                secb9 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP11' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb9 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob9 = "", Toprnsb9 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP11' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb9 = dr["TWheight"].ToString();
                            Trsnob9 = dr["TRSNo"].ToString();
                            Toprnsb9 = dr["TOPsn"].ToString();
                            lotb9 = dr["LotCode"].ToString();
                        }

                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP11' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc9 = "", lotc9 = "", mirc9 = "", stc9 = "";
                        while (dr.Read())
                        {

                            if (!secc9.Contains(dr["Mirno"].ToString()))
                                secc9 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP11' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc9 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc9 = "", Toprnsc9 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP11' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc9 = dr["TWheight"].ToString();
                            Trsnoc9 = dr["TRSNo"].ToString();
                            Toprnsc9 = dr["TOPsn"].ToString();
                            lotc9 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP11' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD9 = "", lotD9 = "", mirD9 = "", std9 = "";
                        while (dr.Read())
                        {
                            if (!secD9.Contains(dr["Mirno"].ToString()))
                                secD9 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP11' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std9 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD9 = "", ToprnsD9 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP11' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD9 = dr["TWheight"].ToString();
                            TrsnoD9 = dr["TRSNo"].ToString();
                            ToprnsD9 = dr["TOPsn"].ToString();
                            lotD9 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        //******************************CP-19**************************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP19' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb10 = "", lotb10 = "", mirb10 = "", stb10 = "";
                        while (dr.Read())
                        {
                            if (!secb10.Contains(dr["Mirno"].ToString()))
                                secb10 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP19' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb10 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob10 = "", Toprnsb10 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP19' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb10 = dr["TWheight"].ToString();
                            Trsnob10 = dr["TRSNo"].ToString();
                            Toprnsb10 = dr["TOPsn"].ToString();
                            lotb10 = dr["LotCode"].ToString();
                        }

                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP19' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc10 = "", lotc10 = "", mirc10 = "", stc10 = "";
                        while (dr.Read())
                        {

                            if (!secc10.Contains(dr["Mirno"].ToString()))
                                secc10 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP19' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc10 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc10 = "", Toprnsc10 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP19' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc10 = dr["TWheight"].ToString();
                            Trsnoc10 = dr["TRSNo"].ToString();
                            Toprnsc10 = dr["TOPsn"].ToString();
                            lotc10 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP19' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD10 = "", lotD10 = "", mirD10 = "", std10 = "";
                        while (dr.Read())
                        {
                            if (!secD10.Contains(dr["Mirno"].ToString()))
                                secD10 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP19' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std10 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD10 = "", ToprnsD10 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP19' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD10 = dr["TWheight"].ToString();
                            TrsnoD10 = dr["TRSNo"].ToString();
                            ToprnsD10 = dr["TOPsn"].ToString();
                            lotD10 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        //**************************CP20************************************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP20' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb11 = "", lotb11 = "", mirb11 = "", stb11 = "";
                        while (dr.Read())
                        {
                            if (!secb11.Contains(dr["Mirno"].ToString()))
                                secb11 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP20' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb11 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob11 = "", Toprnsb11 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP20' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb11 = dr["TWheight"].ToString();
                            Trsnob11 = dr["TRSNo"].ToString();
                            Toprnsb11 = dr["TOPsn"].ToString();
                            lotb11 = dr["LotCode"].ToString();
                        }

                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP20' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc11 = "", lotc11 = "", mirc11 = "", stc11 = "";
                        while (dr.Read())
                        {

                            if (!secc11.Contains(dr["Mirno"].ToString()))
                                secc11 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP20' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc11 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc11 = "", Toprnsc11 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP20' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc11 = dr["TWheight"].ToString();
                            Trsnoc11 = dr["TRSNo"].ToString();
                            Toprnsc11 = dr["TOPsn"].ToString();
                            lotc11 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP20' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD11 = "", lotD11 = "", mirD11 = "", std11 = "";
                        while (dr.Read())
                        {
                            if (!secD11.Contains(dr["Mirno"].ToString()))
                                secD11 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP20' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std11 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD11 = "", ToprnsD11 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP20' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD11 = dr["TWheight"].ToString();
                            TrsnoD11 = dr["TRSNo"].ToString();
                            ToprnsD11 = dr["TOPsn"].ToString();
                            lotD11 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        //*****************************CP-21***********************************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP21' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb12 = "", lotb12 = "", mirb12 = "", stb12 = "";
                        while (dr.Read())
                        {
                            if (!secb12.Contains(dr["Mirno"].ToString()))
                                secb12 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP21' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb12 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob12 = "", Toprnsb12 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP21' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb12 = dr["TWheight"].ToString();
                            Trsnob12 = dr["TRSNo"].ToString();
                            Toprnsb12 = dr["TOPsn"].ToString();
                            lotb12 = dr["LotCode"].ToString();
                        }

                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP21' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc12 = "", lotc12 = "", mirc12 = "", stc12 = "";
                        while (dr.Read())
                        {

                            if (!secc12.Contains(dr["Mirno"].ToString()))
                                secc12 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP21' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc12 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc12 = "", Toprnsc12 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP21' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc12 = dr["TWheight"].ToString();
                            Trsnoc12 = dr["TRSNo"].ToString();
                            Toprnsc12 = dr["TOPsn"].ToString();
                            lotc12 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP21' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD12 = "", lotD12 = "", mirD12 = "", std12 = "";
                        while (dr.Read())
                        {
                            if (!secD12.Contains(dr["Mirno"].ToString()))
                                secD12 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP21' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std12 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD12 = "", ToprnsD12 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP21' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD12 = dr["TWheight"].ToString();
                            TrsnoD12 = dr["TRSNo"].ToString();
                            ToprnsD12 = dr["TOPsn"].ToString();
                            lotD12 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        //*******************************CP-24*****************************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP24' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb13 = "", lotb13 = "", mirb13 = "", stb13 = "";
                        while (dr.Read())
                        {
                            if (!secb13.Contains(dr["Mirno"].ToString()))
                                secb13 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP24' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb13 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob13 = "", Toprnsb13 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP24' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb13 = dr["TWheight"].ToString();
                            Trsnob13 = dr["TRSNo"].ToString();
                            Toprnsb13 = dr["TOPsn"].ToString();
                            lotb13 = dr["LotCode"].ToString();
                        }

                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP24' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc13 = "", lotc13 = "", mirc13 = "", stc13 = "";
                        while (dr.Read())
                        {

                            if (!secc13.Contains(dr["Mirno"].ToString()))
                                secc13 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP24' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc13 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc13 = "", Toprnsc13 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP24' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc13 = dr["TWheight"].ToString();
                            Trsnoc13 = dr["TRSNo"].ToString();
                            Toprnsc13 = dr["TOPsn"].ToString();
                            lotc13 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP24' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD13 = "", lotD13 = "", mirD13 = "", std13 = "";
                        while (dr.Read())
                        {
                            if (!secD13.Contains(dr["Mirno"].ToString()))
                                secD13 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP24' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std13 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD13 = "", ToprnsD13 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP24' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD13 = dr["TWheight"].ToString();
                            TrsnoD13 = dr["TRSNo"].ToString();
                            ToprnsD13 = dr["TOPsn"].ToString();
                            lotD13 = dr["LotCode"].ToString();
                        }
                        dr.Close();


                        //***************************CP25*************************************
                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP25' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb14 = "", lotb14 = "", mirb14 = "", stb14 = "";
                        while (dr.Read())
                        {
                            if (!secb14.Contains(dr["Mirno"].ToString()))
                                secb14 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP25' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb14 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob14 = "", Toprnsb14 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP25' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb14 = dr["TWheight"].ToString();
                            Trsnob14 = dr["TRSNo"].ToString();
                            Toprnsb14 = dr["TOPsn"].ToString();
                            lotb14 = dr["LotCode"].ToString();
                        }

                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP25' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc14 = "", lotc14 = "", mirc14 = "", stc14 = "";
                        while (dr.Read())
                        {

                            if (!secc14.Contains(dr["Mirno"].ToString()))
                                secc14 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP25' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc14 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc14 = "", Toprnsc14 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP25' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc14 = dr["TWheight"].ToString();
                            Trsnoc14 = dr["TRSNo"].ToString();
                            Toprnsc14 = dr["TOPsn"].ToString();
                            lotc14 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP25' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD14 = "", lotD14 = "", mirD14 = "", std14 = "";
                        while (dr.Read())
                        {
                            if (!secD14.Contains(dr["Mirno"].ToString()))
                                secD14 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP25' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std13 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD14 = "", ToprnsD14 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP25' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD14 = dr["TWheight"].ToString();
                            TrsnoD14 = dr["TRSNo"].ToString();
                            ToprnsD14 = dr["TOPsn"].ToString();
                            lotD14 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        //******************clear******************************************************

                        //  *******CP19**************
                        ((Excel123.Range)wrksheet.Cells["1", "BI"]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["2", CL]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["4", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["4", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["4", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["5", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["5", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["5", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["6", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["6", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["6", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["7", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["7", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["7", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["8", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["8", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["8", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["9", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["9", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["9", CM]).Value2 = "";

                        //****************************************** ************************************************
                        //  *******CP20**************
                        ((Excel123.Range)wrksheet.Cells["10", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["10", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["10", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["11", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["11", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["11", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["12", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["12", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["12", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["13", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["13", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["13", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["14", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["14", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["14", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["15", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["15", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["15", CM]).Value2 = "";

                        //***************************************************************************************
                        //  *******CP21**************
                        ((Excel123.Range)wrksheet.Cells["16", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["16", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["16", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["17", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["17", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["17", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["18", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["18", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["18", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["19", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["19", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["19", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["20", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["20", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["20", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["21", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["21", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["21", CM]).Value2 = "";

                        //********************************************************************************************
                        //  *******CP01**************
                        ((Excel123.Range)wrksheet.Cells["22", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["22", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["22", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["23", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["23", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["23", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["24", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["24", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["24", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["25", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["25", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["25", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["26", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["26", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["26", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["27", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["27", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["27", CM]).Value2 = "";


                        //********************************************************************
                        //  *******CP02**************

                        ((Excel123.Range)wrksheet.Cells["28", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["28", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["28", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["29", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["29", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["29", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["30", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["30", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["30", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["31", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["31", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["31", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["32", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["32", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["32", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["33", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["33", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["33", CM]).Value2 = "";

                        //*************************************************************************************************************
                        //  *******CP03**************

                        ((Excel123.Range)wrksheet.Cells["34", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["34", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["34", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["35", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["35", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["35", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["36", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["36", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["36", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["37", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["37", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["37", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["38", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["38", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["38", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["39", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["39", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["39", CM]).Value2 = "";

                        //*************************CP-04************************************

                        ((Excel123.Range)wrksheet.Cells["40", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["40", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["40", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["41", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["41", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["41", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["42", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["42", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["42", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["43", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["43", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["43", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["44", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["44", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["44", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["45", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["45", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["45", CM]).Value2 = "";

                        //***********************CP-05*****************************
                        ((Excel123.Range)wrksheet.Cells["46", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["46", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["46", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["47", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["47", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["47", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["48", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["48", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["48", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["49", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["49", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["49", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["50", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["50", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["50", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["51", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["51", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["51", CM]).Value2 = "";

                        //****************************CP-06*****************************************

                        ((Excel123.Range)wrksheet.Cells["52", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["52", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["52", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["53", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["53", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["53", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["54", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["54", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["54", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["55", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["55", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["55", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["56", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["56", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["56", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["57", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["57", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["57", CM]).Value2 = "";

                        //****************************** CP-07*****************************************

                        ((Excel123.Range)wrksheet.Cells["58", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["58", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["58", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["59", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["59", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["59", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["60", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["60", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["60", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["61", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["61", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["61", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["62", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["62", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["62", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["63", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["63", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["63", CM]).Value2 = "";

                        //***********************CP-08***********************

                        ((Excel123.Range)wrksheet.Cells["64", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["64", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["64", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["65", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["65", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["65", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["66", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["66", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["66", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["67", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["67", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["67", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["68", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["68", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["68", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["69", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["69", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["69", CM]).Value2 = "";

                        //***********************************CP-09****************

                        ((Excel123.Range)wrksheet.Cells["70", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["70", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["70", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["71", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["71", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["71", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["72", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["72", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["72", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["73", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["73", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["73", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["74", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["74", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["74", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["75", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["75", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["75", CM]).Value2 = "";

                        //*******************CP-10***************************

                        ((Excel123.Range)wrksheet.Cells["76", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["76", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["76", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["77", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["77", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["77", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["78", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["78", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["78", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["79", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["79", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["79", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["80", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["80", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["80", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["81", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["81", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["81", CM]).Value2 = "";

                        //***************************CD22*******************************

                        ((Excel123.Range)wrksheet.Cells["82", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["82", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["82", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["83", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["83", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["83", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["84", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["84", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["84", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["85", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["85", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["85", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["86", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["86", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["86", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["87", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["87", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["87", CM]).Value2 = "";


                        //********************* insert data *********************
                        //***************cp19
                        ((Excel123.Range)wrksheet.Cells["1", "BI"]).Value2 = currentTime;
                        ((Excel123.Range)wrksheet.Cells["2", CL]).Value2 = scandate;

                        ((Excel123.Range)wrksheet.Cells["4", CK]).Value2 = sec;
                        ((Excel123.Range)wrksheet.Cells["4", CL]).Value2 = secC1;
                        ((Excel123.Range)wrksheet.Cells["4", CM]).Value2 = secD1;

                        ((Excel123.Range)wrksheet.Cells["5", CK]).Value2 = tot;
                        ((Excel123.Range)wrksheet.Cells["5", CL]).Value2 = mirC1;
                        ((Excel123.Range)wrksheet.Cells["5", CM]).Value2 = mirD1;


                        ((Excel123.Range)wrksheet.Cells["6", CK]).Value2 = Trsno;
                        ((Excel123.Range)wrksheet.Cells["6", CL]).Value2 = TrsnoC1;
                        ((Excel123.Range)wrksheet.Cells["6", CM]).Value2 = TrsnoD1;

                        ((Excel123.Range)wrksheet.Cells["7", CK]).Value2 = Toprns;
                        ((Excel123.Range)wrksheet.Cells["7", CL]).Value2 = ToprnsC1;
                        ((Excel123.Range)wrksheet.Cells["7", CM]).Value2 = ToprnsD1;

                        ((Excel123.Range)wrksheet.Cells["8", CK]).Value2 = st;
                        ((Excel123.Range)wrksheet.Cells["8", CL]).Value2 = stc1;
                        ((Excel123.Range)wrksheet.Cells["8", CM]).Value2 = std1;

                        ((Excel123.Range)wrksheet.Cells["9", CK]).Value2 = lot;
                        ((Excel123.Range)wrksheet.Cells["9", CL]).Value2 = lotC1;
                        ((Excel123.Range)wrksheet.Cells["9", CM]).Value2 = lotD1;

                        //**************************************************************************************
                        //************ cp20*****************

                        ((Excel123.Range)wrksheet.Cells["10", CK]).Value2 = secb2;
                        ((Excel123.Range)wrksheet.Cells["10", CL]).Value2 = secc2;
                        ((Excel123.Range)wrksheet.Cells["10", CM]).Value2 = secD2;


                        ((Excel123.Range)wrksheet.Cells["11", CK]).Value2 = mirb2;
                        ((Excel123.Range)wrksheet.Cells["11", CL]).Value2 = mirc2;
                        ((Excel123.Range)wrksheet.Cells["11", CM]).Value2 = mirD2;

                        ((Excel123.Range)wrksheet.Cells["12", CK]).Value2 = Trsnob2;
                        ((Excel123.Range)wrksheet.Cells["12", CL]).Value2 = Trsnoc2;
                        ((Excel123.Range)wrksheet.Cells["12", CM]).Value2 = TrsnoD2;

                        ((Excel123.Range)wrksheet.Cells["13", CK]).Value2 = Toprnsb2;
                        ((Excel123.Range)wrksheet.Cells["13", CL]).Value2 = Toprnsc2;
                        ((Excel123.Range)wrksheet.Cells["13", CM]).Value2 = ToprnsD2;

                        ((Excel123.Range)wrksheet.Cells["14", CK]).Value2 = stb2;
                        ((Excel123.Range)wrksheet.Cells["14", CL]).Value2 = stc2;
                        ((Excel123.Range)wrksheet.Cells["14", CM]).Value2 = std2;

                        ((Excel123.Range)wrksheet.Cells["15", CK]).Value2 = lotb2;
                        ((Excel123.Range)wrksheet.Cells["15", CL]).Value2 = lotc2;
                        ((Excel123.Range)wrksheet.Cells["15", CM]).Value2 = lotD2;

                        //**************************************************************************************
                        //**************cp21****************
                        ((Excel123.Range)wrksheet.Cells["16", CK]).Value2 = secb3;
                        ((Excel123.Range)wrksheet.Cells["16", CL]).Value2 = secc3;
                        ((Excel123.Range)wrksheet.Cells["16", CM]).Value2 = secD3;

                        ((Excel123.Range)wrksheet.Cells["17", CK]).Value2 = mirb3;
                        ((Excel123.Range)wrksheet.Cells["17", CL]).Value2 = mirc3;
                        ((Excel123.Range)wrksheet.Cells["17", CM]).Value2 = mirD3;

                        ((Excel123.Range)wrksheet.Cells["18", CK]).Value2 = Trsnob3;
                        ((Excel123.Range)wrksheet.Cells["18", CL]).Value2 = Trsnoc3;
                        ((Excel123.Range)wrksheet.Cells["18", CM]).Value2 = TrsnoD3;

                        ((Excel123.Range)wrksheet.Cells["19", CK]).Value2 = Toprnsb3;
                        ((Excel123.Range)wrksheet.Cells["19", CL]).Value2 = Toprnsc3;
                        ((Excel123.Range)wrksheet.Cells["19", CM]).Value2 = ToprnsD3;

                        ((Excel123.Range)wrksheet.Cells["20", CK]).Value2 = stb3;
                        ((Excel123.Range)wrksheet.Cells["20", CL]).Value2 = stc3;
                        ((Excel123.Range)wrksheet.Cells["20", CM]).Value2 = std3;


                        ((Excel123.Range)wrksheet.Cells["21", CK]).Value2 = lotb3;
                        ((Excel123.Range)wrksheet.Cells["21", CL]).Value2 = lotc3;
                        ((Excel123.Range)wrksheet.Cells["21", CM]).Value2 = lotD3;

                        //***********************************************************************************
                        //******************************cp01**********************
                        ((Excel123.Range)wrksheet.Cells["22", CK]).Value2 = secb4;
                        ((Excel123.Range)wrksheet.Cells["22", CL]).Value2 = secc4;
                        ((Excel123.Range)wrksheet.Cells["22", CM]).Value2 = secD4;

                        ((Excel123.Range)wrksheet.Cells["23", CK]).Value2 = mirb4;
                        ((Excel123.Range)wrksheet.Cells["23", CL]).Value2 = mirc4;
                        ((Excel123.Range)wrksheet.Cells["23", CM]).Value2 = mirD4;

                        ((Excel123.Range)wrksheet.Cells["24", CK]).Value2 = Trsnob4;
                        ((Excel123.Range)wrksheet.Cells["24", CL]).Value2 = Trsnoc4;
                        ((Excel123.Range)wrksheet.Cells["24", CM]).Value2 = TrsnoD4;

                        ((Excel123.Range)wrksheet.Cells["25", CK]).Value2 = Toprnsb4;
                        ((Excel123.Range)wrksheet.Cells["25", CL]).Value2 = Toprnsc4;
                        ((Excel123.Range)wrksheet.Cells["25", CM]).Value2 = ToprnsD4;

                        ((Excel123.Range)wrksheet.Cells["26", CK]).Value2 = stb4;
                        ((Excel123.Range)wrksheet.Cells["26", CL]).Value2 = stc4;
                        ((Excel123.Range)wrksheet.Cells["26", CM]).Value2 = std4;


                        ((Excel123.Range)wrksheet.Cells["27", CK]).Value2 = lotb4;
                        ((Excel123.Range)wrksheet.Cells["27", CL]).Value2 = lotc4;
                        ((Excel123.Range)wrksheet.Cells["27", CM]).Value2 = lotD4;

                        //*********************************************************************
                        //********************cp02**********************
                        ((Excel123.Range)wrksheet.Cells["28", CK]).Value2 = secb5;
                        ((Excel123.Range)wrksheet.Cells["28", CL]).Value2 = secc5;
                        ((Excel123.Range)wrksheet.Cells["28", CM]).Value2 = secD5;

                        ((Excel123.Range)wrksheet.Cells["29", CK]).Value2 = mirb5;
                        ((Excel123.Range)wrksheet.Cells["29", CL]).Value2 = mirc5;
                        ((Excel123.Range)wrksheet.Cells["29", CM]).Value2 = mirD5;

                        ((Excel123.Range)wrksheet.Cells["30", CK]).Value2 = Trsnob5;
                        ((Excel123.Range)wrksheet.Cells["30", CL]).Value2 = Trsnoc5;
                        ((Excel123.Range)wrksheet.Cells["30", CM]).Value2 = TrsnoD5;

                        ((Excel123.Range)wrksheet.Cells["31", CK]).Value2 = Toprnsb5;
                        ((Excel123.Range)wrksheet.Cells["31", CL]).Value2 = Toprnsc5;
                        ((Excel123.Range)wrksheet.Cells["31", CM]).Value2 = ToprnsD5;

                        ((Excel123.Range)wrksheet.Cells["32", CK]).Value2 = stb5;
                        ((Excel123.Range)wrksheet.Cells["32", CL]).Value2 = stc5;
                        ((Excel123.Range)wrksheet.Cells["32", CM]).Value2 = std5;


                        ((Excel123.Range)wrksheet.Cells["33", CK]).Value2 = lotb5;
                        ((Excel123.Range)wrksheet.Cells["33", CL]).Value2 = lotc5;
                        ((Excel123.Range)wrksheet.Cells["33", CM]).Value2 = lotD5;

                        //*********************************************************
                        //******************cp03*********************

                        ((Excel123.Range)wrksheet.Cells["34", CK]).Value2 = secb6;
                        ((Excel123.Range)wrksheet.Cells["34", CL]).Value2 = secc6;
                        ((Excel123.Range)wrksheet.Cells["34", CM]).Value2 = secD6;

                        ((Excel123.Range)wrksheet.Cells["35", CK]).Value2 = mirb6;
                        ((Excel123.Range)wrksheet.Cells["35", CL]).Value2 = mirc6;
                        ((Excel123.Range)wrksheet.Cells["35", CM]).Value2 = mirD6;

                        ((Excel123.Range)wrksheet.Cells["36", CK]).Value2 = Trsnob6;
                        ((Excel123.Range)wrksheet.Cells["36", CL]).Value2 = Trsnoc6;
                        ((Excel123.Range)wrksheet.Cells["36", CM]).Value2 = TrsnoD6;

                        ((Excel123.Range)wrksheet.Cells["37", CK]).Value2 = Toprnsb6;
                        ((Excel123.Range)wrksheet.Cells["37", CL]).Value2 = Toprnsc6;
                        ((Excel123.Range)wrksheet.Cells["37", CM]).Value2 = ToprnsD6;

                        ((Excel123.Range)wrksheet.Cells["38", CK]).Value2 = stb6;
                        ((Excel123.Range)wrksheet.Cells["38", CL]).Value2 = stc6;
                        ((Excel123.Range)wrksheet.Cells["38", CM]).Value2 = std6;


                        ((Excel123.Range)wrksheet.Cells["39", CK]).Value2 = lotb6;
                        ((Excel123.Range)wrksheet.Cells["39", CL]).Value2 = lotc6;
                        ((Excel123.Range)wrksheet.Cells["39", CM]).Value2 = lotD6;

                        //*******************CP-04*****************************

                        ((Excel123.Range)wrksheet.Cells["40", CK]).Value2 = secb7;
                        ((Excel123.Range)wrksheet.Cells["40", CL]).Value2 = secc7;
                        ((Excel123.Range)wrksheet.Cells["40", CM]).Value2 = secD7;

                        ((Excel123.Range)wrksheet.Cells["41", CK]).Value2 = mirb7;
                        ((Excel123.Range)wrksheet.Cells["41", CL]).Value2 = mirc7;
                        ((Excel123.Range)wrksheet.Cells["41", CM]).Value2 = mirD7;

                        ((Excel123.Range)wrksheet.Cells["42", CK]).Value2 = Trsnob7;
                        ((Excel123.Range)wrksheet.Cells["42", CL]).Value2 = Trsnoc7;
                        ((Excel123.Range)wrksheet.Cells["42", CM]).Value2 = TrsnoD7;

                        ((Excel123.Range)wrksheet.Cells["43", CK]).Value2 = Toprnsb7;
                        ((Excel123.Range)wrksheet.Cells["43", CL]).Value2 = Toprnsc7;
                        ((Excel123.Range)wrksheet.Cells["43", CM]).Value2 = ToprnsD7;

                        ((Excel123.Range)wrksheet.Cells["44", CK]).Value2 = stb7;
                        ((Excel123.Range)wrksheet.Cells["44", CL]).Value2 = stc7;
                        ((Excel123.Range)wrksheet.Cells["44", CM]).Value2 = std7;


                        ((Excel123.Range)wrksheet.Cells["45", CK]).Value2 = lotb7;
                        ((Excel123.Range)wrksheet.Cells["45", CL]).Value2 = lotc7;
                        ((Excel123.Range)wrksheet.Cells["45", CM]).Value2 = lotD7;

                        //**********************CP-05*************************
                        ((Excel123.Range)wrksheet.Cells["46", CK]).Value2 = secb8;
                        ((Excel123.Range)wrksheet.Cells["46", CL]).Value2 = secc8;
                        ((Excel123.Range)wrksheet.Cells["46", CM]).Value2 = secD8;

                        ((Excel123.Range)wrksheet.Cells["47", CK]).Value2 = mirb8;
                        ((Excel123.Range)wrksheet.Cells["47", CL]).Value2 = mirc8;
                        ((Excel123.Range)wrksheet.Cells["47", CM]).Value2 = mirD8;

                        ((Excel123.Range)wrksheet.Cells["48", CK]).Value2 = Trsnob8;
                        ((Excel123.Range)wrksheet.Cells["48", CL]).Value2 = Trsnoc8;
                        ((Excel123.Range)wrksheet.Cells["48", CM]).Value2 = TrsnoD8;

                        ((Excel123.Range)wrksheet.Cells["49", CK]).Value2 = Toprnsb8;
                        ((Excel123.Range)wrksheet.Cells["49", CL]).Value2 = Toprnsc8;
                        ((Excel123.Range)wrksheet.Cells["49", CM]).Value2 = ToprnsD8;

                        ((Excel123.Range)wrksheet.Cells["50", CK]).Value2 = stb8;
                        ((Excel123.Range)wrksheet.Cells["50", CL]).Value2 = stc8;
                        ((Excel123.Range)wrksheet.Cells["50", CM]).Value2 = std8;


                        ((Excel123.Range)wrksheet.Cells["51", CK]).Value2 = lotb8;
                        ((Excel123.Range)wrksheet.Cells["51", CL]).Value2 = lotc8;
                        ((Excel123.Range)wrksheet.Cells["51", CM]).Value2 = lotD8;

                        //****************************CP-06*****************************************

                        ((Excel123.Range)wrksheet.Cells["52", CK]).Value2 = secb9;
                        ((Excel123.Range)wrksheet.Cells["52", CL]).Value2 = secc9;
                        ((Excel123.Range)wrksheet.Cells["52", CM]).Value2 = secD9;

                        ((Excel123.Range)wrksheet.Cells["53", CK]).Value2 = mirb9;
                        ((Excel123.Range)wrksheet.Cells["53", CL]).Value2 = mirc9;
                        ((Excel123.Range)wrksheet.Cells["53", CM]).Value2 = mirD9;

                        ((Excel123.Range)wrksheet.Cells["54", CK]).Value2 = Trsnob9;
                        ((Excel123.Range)wrksheet.Cells["54", CL]).Value2 = Trsnoc9;
                        ((Excel123.Range)wrksheet.Cells["54", CM]).Value2 = TrsnoD9;

                        ((Excel123.Range)wrksheet.Cells["55", CK]).Value2 = Toprnsb9;
                        ((Excel123.Range)wrksheet.Cells["55", CL]).Value2 = Toprnsc9;
                        ((Excel123.Range)wrksheet.Cells["55", CM]).Value2 = ToprnsD9;

                        ((Excel123.Range)wrksheet.Cells["56", CK]).Value2 = stb9;
                        ((Excel123.Range)wrksheet.Cells["56", CL]).Value2 = stc9;
                        ((Excel123.Range)wrksheet.Cells["56", CM]).Value2 = std9;


                        ((Excel123.Range)wrksheet.Cells["57", CK]).Value2 = lotb9;
                        ((Excel123.Range)wrksheet.Cells["57", CL]).Value2 = lotc9;
                        ((Excel123.Range)wrksheet.Cells["57", CM]).Value2 = lotD9;

                        //******************CP-07*****************************************

                        ((Excel123.Range)wrksheet.Cells["58", CK]).Value2 = secb10;
                        ((Excel123.Range)wrksheet.Cells["58", CL]).Value2 = secc10;
                        ((Excel123.Range)wrksheet.Cells["58", CM]).Value2 = secD10;

                        ((Excel123.Range)wrksheet.Cells["59", CK]).Value2 = mirb10;
                        ((Excel123.Range)wrksheet.Cells["59", CL]).Value2 = mirc10;
                        ((Excel123.Range)wrksheet.Cells["59", CM]).Value2 = mirD10;

                        ((Excel123.Range)wrksheet.Cells["60", CK]).Value2 = Trsnob10;
                        ((Excel123.Range)wrksheet.Cells["60", CL]).Value2 = Trsnoc10;
                        ((Excel123.Range)wrksheet.Cells["60", CM]).Value2 = TrsnoD10;

                        ((Excel123.Range)wrksheet.Cells["61", CK]).Value2 = Toprnsb10;
                        ((Excel123.Range)wrksheet.Cells["61", CL]).Value2 = Toprnsc10;
                        ((Excel123.Range)wrksheet.Cells["61", CM]).Value2 = ToprnsD10;

                        ((Excel123.Range)wrksheet.Cells["62", CK]).Value2 = stb10;
                        ((Excel123.Range)wrksheet.Cells["62", CL]).Value2 = stc10;
                        ((Excel123.Range)wrksheet.Cells["62", CM]).Value2 = std10;


                        ((Excel123.Range)wrksheet.Cells["63", CK]).Value2 = lotb10;
                        ((Excel123.Range)wrksheet.Cells["63", CL]).Value2 = lotc10;
                        ((Excel123.Range)wrksheet.Cells["63", CM]).Value2 = lotD10;

                        //***********************CP-08******************************

                        ((Excel123.Range)wrksheet.Cells["64", CK]).Value2 = secb11;
                        ((Excel123.Range)wrksheet.Cells["64", CL]).Value2 = secc11;
                        ((Excel123.Range)wrksheet.Cells["64", CM]).Value2 = secD11;

                        ((Excel123.Range)wrksheet.Cells["65", CK]).Value2 = mirb11;
                        ((Excel123.Range)wrksheet.Cells["65", CL]).Value2 = mirc11;
                        ((Excel123.Range)wrksheet.Cells["65", CM]).Value2 = mirD11;

                        ((Excel123.Range)wrksheet.Cells["66", CK]).Value2 = Trsnob11;
                        ((Excel123.Range)wrksheet.Cells["66", CL]).Value2 = Trsnoc11;
                        ((Excel123.Range)wrksheet.Cells["66", CM]).Value2 = TrsnoD11;

                        ((Excel123.Range)wrksheet.Cells["67", CK]).Value2 = Toprnsb11;
                        ((Excel123.Range)wrksheet.Cells["67", CL]).Value2 = Toprnsc11;
                        ((Excel123.Range)wrksheet.Cells["67", CM]).Value2 = ToprnsD11;

                        ((Excel123.Range)wrksheet.Cells["68", CK]).Value2 = stb11;
                        ((Excel123.Range)wrksheet.Cells["68", CL]).Value2 = stc11;
                        ((Excel123.Range)wrksheet.Cells["68", CM]).Value2 = std11;


                        ((Excel123.Range)wrksheet.Cells["69", CK]).Value2 = lotb11;
                        ((Excel123.Range)wrksheet.Cells["69", CL]).Value2 = lotc11;
                        ((Excel123.Range)wrksheet.Cells["69", CM]).Value2 = lotD11;

                        //***********************CP-09******************************

                        ((Excel123.Range)wrksheet.Cells["70", CK]).Value2 = secb12;
                        ((Excel123.Range)wrksheet.Cells["70", CL]).Value2 = secc12;
                        ((Excel123.Range)wrksheet.Cells["70", CM]).Value2 = secD12;

                        ((Excel123.Range)wrksheet.Cells["71", CK]).Value2 = mirb12;
                        ((Excel123.Range)wrksheet.Cells["71", CL]).Value2 = mirc12;
                        ((Excel123.Range)wrksheet.Cells["71", CM]).Value2 = mirD12;

                        ((Excel123.Range)wrksheet.Cells["72", CK]).Value2 = Trsnob12;
                        ((Excel123.Range)wrksheet.Cells["72", CL]).Value2 = Trsnoc12;
                        ((Excel123.Range)wrksheet.Cells["72", CM]).Value2 = TrsnoD12;

                        ((Excel123.Range)wrksheet.Cells["73", CK]).Value2 = Toprnsb12;
                        ((Excel123.Range)wrksheet.Cells["73", CL]).Value2 = Toprnsc12;
                        ((Excel123.Range)wrksheet.Cells["73", CM]).Value2 = ToprnsD12;

                        ((Excel123.Range)wrksheet.Cells["74", CK]).Value2 = stb12;
                        ((Excel123.Range)wrksheet.Cells["74", CL]).Value2 = stc12;
                        ((Excel123.Range)wrksheet.Cells["74", CM]).Value2 = std12;


                        ((Excel123.Range)wrksheet.Cells["75", CK]).Value2 = lotb12;
                        ((Excel123.Range)wrksheet.Cells["75", CL]).Value2 = lotc12;
                        ((Excel123.Range)wrksheet.Cells["75", CM]).Value2 = lotD12;

                        //*********************CP-10*********************************

                        ((Excel123.Range)wrksheet.Cells["76", CK]).Value2 = secb13;
                        ((Excel123.Range)wrksheet.Cells["76", CL]).Value2 = secc13;
                        ((Excel123.Range)wrksheet.Cells["76", CM]).Value2 = secD13;

                        ((Excel123.Range)wrksheet.Cells["77", CK]).Value2 = mirb13;
                        ((Excel123.Range)wrksheet.Cells["77", CL]).Value2 = mirc13;
                        ((Excel123.Range)wrksheet.Cells["77", CM]).Value2 = mirD13;

                        ((Excel123.Range)wrksheet.Cells["78", CK]).Value2 = Trsnob13;
                        ((Excel123.Range)wrksheet.Cells["78", CL]).Value2 = Trsnoc13;
                        ((Excel123.Range)wrksheet.Cells["78", CM]).Value2 = TrsnoD13;

                        ((Excel123.Range)wrksheet.Cells["79", CK]).Value2 = Toprnsb13;
                        ((Excel123.Range)wrksheet.Cells["79", CL]).Value2 = Toprnsc13;
                        ((Excel123.Range)wrksheet.Cells["79", CM]).Value2 = ToprnsD13;

                        ((Excel123.Range)wrksheet.Cells["80", CK]).Value2 = stb13;
                        ((Excel123.Range)wrksheet.Cells["80", CL]).Value2 = stc13;
                        ((Excel123.Range)wrksheet.Cells["80", CM]).Value2 = std13;


                        ((Excel123.Range)wrksheet.Cells["81", CK]).Value2 = lotb13;
                        ((Excel123.Range)wrksheet.Cells["81", CL]).Value2 = lotc13;
                        ((Excel123.Range)wrksheet.Cells["81", CM]).Value2 = lotD13;


                        //**********************CD22********************************


                        ((Excel123.Range)wrksheet.Cells["82", CK]).Value2 = secb14;
                        ((Excel123.Range)wrksheet.Cells["82", CL]).Value2 = secc14;
                        ((Excel123.Range)wrksheet.Cells["82", CM]).Value2 = secD14;

                        ((Excel123.Range)wrksheet.Cells["83", CK]).Value2 = mirb14;
                        ((Excel123.Range)wrksheet.Cells["83", CL]).Value2 = mirc14;
                        ((Excel123.Range)wrksheet.Cells["83", CM]).Value2 = mirD14;

                        ((Excel123.Range)wrksheet.Cells["84", CK]).Value2 = Trsnob14;
                        ((Excel123.Range)wrksheet.Cells["84", CL]).Value2 = Trsnoc14;
                        ((Excel123.Range)wrksheet.Cells["84", CM]).Value2 = TrsnoD14;

                        ((Excel123.Range)wrksheet.Cells["85", CK]).Value2 = Toprnsb14;
                        ((Excel123.Range)wrksheet.Cells["85", CL]).Value2 = Toprnsc14;
                        ((Excel123.Range)wrksheet.Cells["85", CM]).Value2 = ToprnsD14;

                        ((Excel123.Range)wrksheet.Cells["86", CK]).Value2 = stb14;
                        ((Excel123.Range)wrksheet.Cells["86", CL]).Value2 = stc14;
                        ((Excel123.Range)wrksheet.Cells["86", CM]).Value2 = std14;


                        ((Excel123.Range)wrksheet.Cells["87", CK]).Value2 = lotb14;
                        ((Excel123.Range)wrksheet.Cells["87", CL]).Value2 = lotc14;
                        ((Excel123.Range)wrksheet.Cells["87", CM]).Value2 = lotD14;






                    }

                    #region Dubai Plant Virtual Scheduling

                    else if (lblPlantCode == "TMD1")
                    {
                        string scan_Date = scandate.Month + "/" + scandate.Day + "/" + scandate.Year + " " + scandate.TimeOfDay;

                        //*************************************CP01***************************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";

                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string sec = "", lot = "", st = "", tot = "", Trsno = "", Toprns = "";
                        while (dr.Read())
                        {
                            if (!sec.Contains(dr["Mirno"].ToString()))
                                //sec += (dr["SctDinemtion"].ToString().Remove(1, dr["SctDinemtion"].ToString().IndexOf('X')) + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                                sec += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";

                            //lot += dr["LotCode"].ToString().Remove(5);
                        }
                        dr.Close();

                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP01' and BP='" + lblPlantCode + "' and Flag_Fab is null";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();

                        while (dr.Read())
                        {
                            st += dr["OPStatus"].ToString();
                        }

                        dr.Close();
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            tot = dr["TWheight"].ToString();
                            Trsno = dr["TRSNo"].ToString();
                            Toprns = dr["TOPsn"].ToString();
                            lot = dr["LotCode"].ToString();
                        }
                        dr.Close();
                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP01' and BP='" + lblPlantCode + "' and Flag_Fab is null";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secC1 = "", lotC1 = "", mirC1 = "", stc1 = "", TrsnoC1 = "", ToprnsC1 = "";
                        while (dr.Read())
                        {
                            if (!secC1.Contains(dr["Mirno"].ToString()))
                                secC1 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotC1 += dr["LotCode"].ToString().Remove(5);
                            //mirC1+= dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();

                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();

                        while (dr.Read())
                        {

                            stc1 += dr["OPStatus"].ToString();


                        }
                        dr.Close();
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP01' and BP='" + lblPlantCode + "' and Flag_Fab is null";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirC1 = dr["TWheight"].ToString();
                            TrsnoC1 = dr["TRSNo"].ToString();
                            ToprnsC1 = dr["TOPsn"].ToString();
                            lotC1 = dr["LotCode"].ToString();
                        }
                        dr.Close();
                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD1 = "", lotD1 = "", mirD1 = "", std1 = "";
                        if (dr.Read())
                        {
                            if (!secD1.Contains(dr["Mirno"].ToString()))
                                secD1 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD1 += dr["LotCode"].ToString().Remove(5);
                            //lotD1 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();

                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();

                        while (dr.Read())
                        {

                            std1 += dr["OPStatus"].ToString();


                        }
                        dr.Close();
                        string TrsnoD1 = "", ToprnsD1 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD1 = dr["TWheight"].ToString();
                            TrsnoD1 = dr["TRSNo"].ToString();
                            ToprnsD1 = dr["TOPsn"].ToString();
                            lotD1 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        //***********************************CP02************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb2 = "", lotb2 = "", mirb2 = "", stb2 = "";
                        while (dr.Read())
                        {
                            if (!secb2.Contains(dr["Mirno"].ToString()))
                                secb2 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            // lotb2 += dr["LotCode"].ToString().Remove(5);
                            //mirb2 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();

                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb2 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob2 = "", Toprnsb2 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb2 = dr["TWheight"].ToString();
                            Trsnob2 = dr["TRSNo"].ToString();
                            Toprnsb2 = dr["TOPsn"].ToString();
                            lotb2 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc2 = "", lotc2 = "", mirc2 = "", stc2 = "";
                        while (dr.Read())
                        {
                            if (!secc2.Contains(dr["Mirno"].ToString()))
                                secc2 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc2 += dr["LotCode"].ToString().Remove(5);
                            //mirc2 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc2 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc2 = "", Toprnsc2 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc2 = dr["TWheight"].ToString();
                            Trsnoc2 = dr["TRSNo"].ToString();
                            Toprnsc2 = dr["TOPsn"].ToString();
                            lotc2 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD2 = "", lotD2 = "", mirD2 = "", std2 = "";
                        while (dr.Read())
                        {
                            if (!secD2.Contains(dr["Mirno"].ToString()))
                                secD2 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD2 += dr["LotCode"].ToString().Remove(5);
                            // mirD2 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std2 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD2 = "", ToprnsD2 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD2 = dr["TWheight"].ToString();
                            TrsnoD2 = dr["TRSNo"].ToString();
                            ToprnsD2 = dr["TOPsn"].ToString();
                            lotD2 = dr["LotCode"].ToString();
                        }
                        dr.Close();


                        //*********************************** CP03 ***********************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb3 = "", lotb3 = "", mirb3 = "", stb3 = "";
                        while (dr.Read())
                        {
                            if (!secb3.Contains(dr["Mirno"].ToString()))
                                secb3 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb3 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob3 = "", Toprnsb3 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb3 = dr["TWheight"].ToString();
                            Trsnob3 = dr["TRSNo"].ToString();
                            Toprnsb3 = dr["TOPsn"].ToString();
                            lotb3 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc3 = "", lotc3 = "", mirc3 = "", stc3 = "";
                        while (dr.Read())
                        {
                            if (!secc3.Contains(dr["Mirno"].ToString()))
                                secc3 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc3 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc3 = "", Toprnsc3 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc3 = dr["TWheight"].ToString();
                            Trsnoc3 = dr["TRSNo"].ToString();
                            Toprnsc3 = dr["TOPsn"].ToString();
                            lotc3 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD3 = "", lotD3 = "", mirD3 = "", std3 = "";
                        while (dr.Read())
                        {
                            if (!secD3.Contains(dr["Mirno"].ToString()))
                                secD3 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std3 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD3 = "", ToprnsD3 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD3 = dr["TWheight"].ToString();
                            TrsnoD3 = dr["TRSNo"].ToString();
                            ToprnsD3 = dr["TOPsn"].ToString();
                            lotD3 = dr["LotCode"].ToString();
                        }
                        dr.Close();


                        //*************************************CP04*************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb4 = "", lotb4 = "", mirb4 = "", stb4 = "";
                        while (dr.Read())
                        {
                            if (!secb4.Contains(dr["Mirno"].ToString()))
                                secb4 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb4 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob4 = "", Toprnsb4 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb4 = dr["TWheight"].ToString();
                            Trsnob4 = dr["TRSNo"].ToString();
                            Toprnsb4 = dr["TOPsn"].ToString();
                            lotb4 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc4 = "", lotc4 = "", mirc4 = "", stc4 = "";
                        while (dr.Read())
                        {
                            if (!secc4.Contains(dr["Mirno"].ToString()))
                                secc4 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc4 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc4 = "", Toprnsc4 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc4 = dr["TWheight"].ToString();
                            Trsnoc4 = dr["TRSNo"].ToString();
                            Toprnsc4 = dr["TOPsn"].ToString();
                            lotc4 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD4 = "", lotD4 = "", mirD4 = "", std4 = "";
                        while (dr.Read())
                        {
                            if (!secD4.Contains(dr["Mirno"].ToString()))
                                secD4 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std4 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD4 = "", ToprnsD4 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP04' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD4 = dr["TWheight"].ToString();
                            TrsnoD4 = dr["TRSNo"].ToString();
                            ToprnsD4 = dr["TOPsn"].ToString();
                            lotD4 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        //**********************************************CP05****************************************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP05' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb5 = "", lotb5 = "", mirb5 = "", stb5 = "";
                        while (dr.Read())
                        {
                            if (!secb5.Contains(dr["Mirno"].ToString()))
                                secb5 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP05' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb5 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob5 = "", Toprnsb5 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP05' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb5 = dr["TWheight"].ToString();
                            Trsnob5 = dr["TRSNo"].ToString();
                            Toprnsb5 = dr["TOPsn"].ToString();
                            lotb5 = dr["LotCode"].ToString();
                        }

                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP05' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc5 = "", lotc5 = "", mirc5 = "", stc5 = "";
                        while (dr.Read())
                        {

                            if (!secc5.Contains(dr["Mirno"].ToString()))
                                secc5 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP05' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc5 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc5 = "", Toprnsc5 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP05' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc5 = dr["TWheight"].ToString();
                            Trsnoc5 = dr["TRSNo"].ToString();
                            Toprnsc5 = dr["TOPsn"].ToString();
                            lotc5 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP05' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD5 = "", lotD5 = "", mirD5 = "", std5 = "";
                        while (dr.Read())
                        {
                            if (!secD5.Contains(dr["Mirno"].ToString()))
                                secD5 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP05' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std5 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD5 = "", ToprnsD5 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP05' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD5 = dr["TWheight"].ToString();
                            TrsnoD5 = dr["TRSNo"].ToString();
                            ToprnsD5 = dr["TOPsn"].ToString();
                            lotD5 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        //************************************** CP-06   ***********************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP06' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb6 = "", lotb6 = "", mirb6 = "", stb6 = "";
                        while (dr.Read())
                        {
                            if (!secb6.Contains(dr["Mirno"].ToString()))
                                secb6 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP06' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb6 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob6 = "", Toprnsb6 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP06' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb6 = dr["TWheight"].ToString();
                            Trsnob6 = dr["TRSNo"].ToString();
                            Toprnsb6 = dr["TOPsn"].ToString();
                            lotb6 = dr["LotCode"].ToString();
                        }

                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP06' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc6 = "", lotc6 = "", mirc6 = "", stc6 = "";
                        while (dr.Read())
                        {

                            if (!secc6.Contains(dr["Mirno"].ToString()))
                                secc6 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP06' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc6 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc6 = "", Toprnsc6 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP06' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc6 = dr["TWheight"].ToString();
                            Trsnoc6 = dr["TRSNo"].ToString();
                            Toprnsc6 = dr["TOPsn"].ToString();
                            lotc6 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP06' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD6 = "", lotD6 = "", mirD6 = "", std6 = "";
                        while (dr.Read())
                        {
                            if (!secD6.Contains(dr["Mirno"].ToString()))
                                secD6 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP06' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std6 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD6 = "", ToprnsD6 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP06' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD6 = dr["TWheight"].ToString();
                            TrsnoD6 = dr["TRSNo"].ToString();
                            ToprnsD6 = dr["TOPsn"].ToString();
                            lotD6 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        //**************************************************CP-07*******************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP07' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb7 = "", lotb7 = "", mirb7 = "", stb7 = "";
                        while (dr.Read())
                        {
                            if (!secb7.Contains(dr["Mirno"].ToString()))
                                secb7 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP07' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb7 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob7 = "", Toprnsb7 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP07' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb7 = dr["TWheight"].ToString();
                            Trsnob7 = dr["TRSNo"].ToString();
                            Toprnsb7 = dr["TOPsn"].ToString();
                            lotb7 = dr["LotCode"].ToString();
                        }

                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP07' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc7 = "", lotc7 = "", mirc7 = "", stc7 = "";
                        while (dr.Read())
                        {

                            if (!secc7.Contains(dr["Mirno"].ToString()))
                                secc7 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP07' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc7 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc7 = "", Toprnsc7 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP07' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc7 = dr["TWheight"].ToString();
                            Trsnoc7 = dr["TRSNo"].ToString();
                            Toprnsc7 = dr["TOPsn"].ToString();
                            lotc7 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP07' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD7 = "", lotD7 = "", mirD7 = "", std7 = "";
                        while (dr.Read())
                        {
                            if (!secD7.Contains(dr["Mirno"].ToString()))
                                secD7 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP07' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std7 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD7 = "", ToprnsD7 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP07' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD7 = dr["TWheight"].ToString();
                            TrsnoD7 = dr["TRSNo"].ToString();
                            ToprnsD7 = dr["TOPsn"].ToString();
                            lotD7 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        //****************************CD01*****************************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CD01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb8 = "", lotb8 = "", mirb8 = "", stb8 = "";
                        while (dr.Read())
                        {
                            if (!secb8.Contains(dr["Mirno"].ToString()))
                                secb8 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CD01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb8 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob8 = "", Toprnsb8 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CD01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb8 = dr["TWheight"].ToString();
                            Trsnob8 = dr["TRSNo"].ToString();
                            Toprnsb8 = dr["TOPsn"].ToString();
                            lotb8 = dr["LotCode"].ToString();
                        }

                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CD01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc8 = "", lotc8 = "", mirc8 = "", stc8 = "";
                        while (dr.Read())
                        {

                            if (!secc8.Contains(dr["Mirno"].ToString()))
                                secc8 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CD01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc8 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc8 = "", Toprnsc8 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP11' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc8 = dr["TWheight"].ToString();
                            Trsnoc8 = dr["TRSNo"].ToString();
                            Toprnsc8 = dr["TOPsn"].ToString();
                            lotc8 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CD01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD8 = "", lotD8 = "", mirD8 = "", std8 = "";
                        while (dr.Read())
                        {
                            if (!secD8.Contains(dr["Mirno"].ToString()))
                                secD8 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CD01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std8 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD8 = "", ToprnsD8 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CD01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD8 = dr["TWheight"].ToString();
                            TrsnoD8 = dr["TRSNo"].ToString();
                            ToprnsD8 = dr["TOPsn"].ToString();
                            lotD8 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        //****************************CD02*****************************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CD02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb9 = "", lotb9 = "", mirb9 = "", stb9 = "";
                        while (dr.Read())
                        {
                            if (!secb9.Contains(dr["Mirno"].ToString()))
                                secb9 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CD02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb9 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob9 = "", Toprnsb9 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CD02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb9 = dr["TWheight"].ToString();
                            Trsnob9 = dr["TRSNo"].ToString();
                            Toprnsb9 = dr["TOPsn"].ToString();
                            lotb9 = dr["LotCode"].ToString();
                        }

                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CD02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc9 = "", lotc9 = "", mirc9 = "", stc9 = "";
                        while (dr.Read())
                        {

                            if (!secc9.Contains(dr["Mirno"].ToString()))
                                secc9 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP11' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc9 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc9 = "", Toprnsc9 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CD02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc9 = dr["TWheight"].ToString();
                            Trsnoc9 = dr["TRSNo"].ToString();
                            Toprnsc9 = dr["TOPsn"].ToString();
                            lotc9 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CD02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD9 = "", lotD9 = "", mirD9 = "", std9 = "";
                        while (dr.Read())
                        {
                            if (!secD9.Contains(dr["Mirno"].ToString()))
                                secD9 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CD02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std9 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD9 = "", ToprnsD9 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CD02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD9 = dr["TWheight"].ToString();
                            TrsnoD9 = dr["TRSNo"].ToString();
                            ToprnsD9 = dr["TOPsn"].ToString();
                            lotD9 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        //*****************************kaltenbech**************************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='kaltenbech' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb10 = "", lotb10 = "", mirb10 = "", stb10 = "";
                        while (dr.Read())
                        {
                            if (!secb10.Contains(dr["Mirno"].ToString()))
                                secb10 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='kaltenbech' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb10 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob10 = "", Toprnsb10 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='kaltenbech' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb10 = dr["TWheight"].ToString();
                            Trsnob10 = dr["TRSNo"].ToString();
                            Toprnsb10 = dr["TOPsn"].ToString();
                            lotb10 = dr["LotCode"].ToString();
                        }

                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='kaltenbech' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc10 = "", lotc10 = "", mirc10 = "", stc10 = "";
                        while (dr.Read())
                        {

                            if (!secc10.Contains(dr["Mirno"].ToString()))
                                secc10 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='kaltenbech' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc10 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc10 = "", Toprnsc10 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='kaltenbech' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc10 = dr["TWheight"].ToString();
                            Trsnoc10 = dr["TRSNo"].ToString();
                            Toprnsc10 = dr["TOPsn"].ToString();
                            lotc10 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='kaltenbech' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD10 = "", lotD10 = "", mirD10 = "", std10 = "";
                        while (dr.Read())
                        {
                            if (!secD10.Contains(dr["Mirno"].ToString()))
                                secD10 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='kaltenbech' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std10 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD10 = "", ToprnsD10 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='kaltenbech' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD10 = dr["TWheight"].ToString();
                            TrsnoD10 = dr["TRSNo"].ToString();
                            ToprnsD10 = dr["TOPsn"].ToString();
                            lotD10 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        //**************************CG01************************************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CG01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb11 = "", lotb11 = "", mirb11 = "", stb11 = "";
                        while (dr.Read())
                        {
                            if (!secb11.Contains(dr["Mirno"].ToString()))
                                secb11 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CG01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb11 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob11 = "", Toprnsb11 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CG01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb11 = dr["TWheight"].ToString();
                            Trsnob11 = dr["TRSNo"].ToString();
                            Toprnsb11 = dr["TOPsn"].ToString();
                            lotb11 = dr["LotCode"].ToString();
                        }

                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CG01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc11 = "", lotc11 = "", mirc11 = "", stc11 = "";
                        while (dr.Read())
                        {

                            if (!secc11.Contains(dr["Mirno"].ToString()))
                                secc11 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CG01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc11 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc11 = "", Toprnsc11 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CG01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc11 = dr["TWheight"].ToString();
                            Trsnoc11 = dr["TRSNo"].ToString();
                            Toprnsc11 = dr["TOPsn"].ToString();
                            lotc11 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CG01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD11 = "", lotD11 = "", mirD11 = "", std11 = "";
                        while (dr.Read())
                        {
                            if (!secD11.Contains(dr["Mirno"].ToString()))
                                secD11 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CG01' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std11 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD11 = "", ToprnsD11 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP20' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD11 = dr["TWheight"].ToString();
                            TrsnoD11 = dr["TRSNo"].ToString();
                            ToprnsD11 = dr["TOPsn"].ToString();
                            lotD11 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        //*****************************CG02***********************************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CG02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb12 = "", lotb12 = "", mirb12 = "", stb12 = "";
                        while (dr.Read())
                        {
                            if (!secb12.Contains(dr["Mirno"].ToString()))
                                secb12 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CG02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb12 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob12 = "", Toprnsb12 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CG02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb12 = dr["TWheight"].ToString();
                            Trsnob12 = dr["TRSNo"].ToString();
                            Toprnsb12 = dr["TOPsn"].ToString();
                            lotb12 = dr["LotCode"].ToString();
                        }

                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CG02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc12 = "", lotc12 = "", mirc12 = "", stc12 = "";
                        while (dr.Read())
                        {

                            if (!secc12.Contains(dr["Mirno"].ToString()))
                                secc12 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CG02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc12 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc12 = "", Toprnsc12 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CG02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc12 = dr["TWheight"].ToString();
                            Trsnoc12 = dr["TRSNo"].ToString();
                            Toprnsc12 = dr["TOPsn"].ToString();
                            lotc12 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CG02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD12 = "", lotD12 = "", mirD12 = "", std12 = "";
                        while (dr.Read())
                        {
                            if (!secD12.Contains(dr["Mirno"].ToString()))
                                secD12 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CG02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std12 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD12 = "", ToprnsD12 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CG02' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD12 = dr["TWheight"].ToString();
                            TrsnoD12 = dr["TRSNo"].ToString();
                            ToprnsD12 = dr["TOPsn"].ToString();
                            lotD12 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        //*******************************CG03*****************************************

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CG03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb13 = "", lotb13 = "", mirb13 = "", stb13 = "";
                        while (dr.Read())
                        {
                            if (!secb13.Contains(dr["Mirno"].ToString()))
                                secb13 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CG03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb13 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob13 = "", Toprnsb13 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CG03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb13 = dr["TWheight"].ToString();
                            Trsnob13 = dr["TRSNo"].ToString();
                            Toprnsb13 = dr["TOPsn"].ToString();
                            lotb13 = dr["LotCode"].ToString();
                        }

                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CG03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc13 = "", lotc13 = "", mirc13 = "", stc13 = "";
                        while (dr.Read())
                        {

                            if (!secc13.Contains(dr["Mirno"].ToString()))
                                secc13 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CG03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc13 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc13 = "", Toprnsc13 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CG03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc13 = dr["TWheight"].ToString();
                            Trsnoc13 = dr["TRSNo"].ToString();
                            Toprnsc13 = dr["TOPsn"].ToString();
                            lotc13 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CG03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD13 = "", lotD13 = "", mirD13 = "", std13 = "";
                        while (dr.Read())
                        {
                            if (!secD13.Contains(dr["Mirno"].ToString()))
                                secD13 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CG03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std13 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD13 = "", ToprnsD13 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CG03' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD13 = dr["TWheight"].ToString();
                            TrsnoD13 = dr["TRSNo"].ToString();
                            ToprnsD13 = dr["TOPsn"].ToString();
                            lotD13 = dr["LotCode"].ToString();
                        }
                        dr.Close();


                        //***************************CP25*************************************
                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP25' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secb14 = "", lotb14 = "", mirb14 = "", stb14 = "";
                        while (dr.Read())
                        {
                            if (!secb14.Contains(dr["Mirno"].ToString()))
                                secb14 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotb3 += dr["LotCode"].ToString().Remove(5);
                            // mirb3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP25' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stb14 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnob14 = "", Toprnsb14 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='First' and MachineName='CP25' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirb14 = dr["TWheight"].ToString();
                            Trsnob14 = dr["TRSNo"].ToString();
                            Toprnsb14 = dr["TOPsn"].ToString();
                            lotb14 = dr["LotCode"].ToString();
                        }

                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP25' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secc14 = "", lotc14 = "", mirc14 = "", stc14 = "";
                        while (dr.Read())
                        {

                            if (!secc14.Contains(dr["Mirno"].ToString()))
                                secc14 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotc3 += dr["LotCode"].ToString().Remove(5);
                            //mirc3 += dr["TotalWt"].ToString() + " / ";
                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP25' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            stc14 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string Trsnoc14 = "", Toprnsc14 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Second' and MachineName='CP25' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirc14 = dr["TWheight"].ToString();
                            Trsnoc14 = dr["TRSNo"].ToString();
                            Toprnsc14 = dr["TOPsn"].ToString();
                            lotc14 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        queryadp = "SELECT distinct SctDinemtion ,LotCode ,Mirno from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP25' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        string secD14 = "", lotD14 = "", mirD14 = "", std14 = "";
                        while (dr.Read())
                        {
                            if (!secD14.Contains(dr["Mirno"].ToString()))
                                secD14 += (dr["SctDinemtion"].ToString() + "(" + dr["Mirno"].ToString() + ")").Replace(" ", "") + "(" + dr["LotCode"].ToString().Remove(5) + ")";
                            //lotD3 += dr["LotCode"].ToString().Remove(5);
                            //mirD3 += dr["TotalWt"].ToString() + " / ";

                        }
                        dr.Close();
                        queryadp = "SELECT distinct OPStatus from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP25' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            std13 += dr["OPStatus"].ToString();
                        }
                        dr.Close();
                        string TrsnoD14 = "", ToprnsD14 = "";
                        queryadp = "SELECT sum(TotalWt)as TWheight,sum(Tot_OPS) as TOPsn,count(RSNo) as TRSNo,sum(RunTime) as LotCode from Operations where PlanningDate='" + scan_Date + "' and PlanningShift='Third' and MachineName='CP25' and BP='" + lblPlantCode + "' and Flag_Fab is null ";
                        cmd = new SqlCommand(queryadp, Conn);
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            mirD14 = dr["TWheight"].ToString();
                            TrsnoD14 = dr["TRSNo"].ToString();
                            ToprnsD14 = dr["TOPsn"].ToString();
                            lotD14 = dr["LotCode"].ToString();
                        }
                        dr.Close();

                        //******************clear******************************************************

                        //  *******CP19**************
                        ((Excel123.Range)wrksheet.Cells["1", "BI"]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["2", CL]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["4", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["4", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["4", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["5", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["5", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["5", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["6", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["6", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["6", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["7", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["7", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["7", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["8", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["8", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["8", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["9", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["9", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["9", CM]).Value2 = "";

                        //****************************************** ************************************************
                        //  *******CP20**************
                        ((Excel123.Range)wrksheet.Cells["10", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["10", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["10", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["11", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["11", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["11", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["12", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["12", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["12", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["13", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["13", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["13", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["14", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["14", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["14", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["15", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["15", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["15", CM]).Value2 = "";

                        //***************************************************************************************
                        //  *******CP21**************
                        ((Excel123.Range)wrksheet.Cells["16", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["16", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["16", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["17", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["17", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["17", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["18", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["18", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["18", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["19", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["19", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["19", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["20", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["20", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["20", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["21", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["21", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["21", CM]).Value2 = "";

                        //********************************************************************************************
                        //  *******CP01**************
                        ((Excel123.Range)wrksheet.Cells["22", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["22", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["22", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["23", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["23", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["23", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["24", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["24", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["24", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["25", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["25", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["25", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["26", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["26", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["26", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["27", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["27", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["27", CM]).Value2 = "";


                        //********************************************************************
                        //  *******CP02**************

                        ((Excel123.Range)wrksheet.Cells["28", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["28", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["28", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["29", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["29", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["29", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["30", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["30", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["30", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["31", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["31", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["31", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["32", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["32", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["32", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["33", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["33", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["33", CM]).Value2 = "";

                        //*************************************************************************************************************
                        //  *******CP03**************

                        ((Excel123.Range)wrksheet.Cells["34", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["34", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["34", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["35", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["35", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["35", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["36", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["36", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["36", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["37", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["37", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["37", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["38", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["38", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["38", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["39", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["39", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["39", CM]).Value2 = "";

                        //*************************CP-04************************************

                        ((Excel123.Range)wrksheet.Cells["40", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["40", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["40", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["41", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["41", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["41", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["42", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["42", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["42", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["43", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["43", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["43", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["44", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["44", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["44", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["45", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["45", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["45", CM]).Value2 = "";

                        //***********************CP-05*****************************
                        ((Excel123.Range)wrksheet.Cells["46", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["46", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["46", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["47", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["47", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["47", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["48", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["48", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["48", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["49", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["49", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["49", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["50", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["50", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["50", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["51", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["51", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["51", CM]).Value2 = "";

                        //****************************CP-06*****************************************

                        ((Excel123.Range)wrksheet.Cells["52", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["52", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["52", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["53", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["53", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["53", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["54", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["54", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["54", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["55", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["55", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["55", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["56", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["56", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["56", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["57", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["57", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["57", CM]).Value2 = "";

                        //****************************** CP-07*****************************************

                        ((Excel123.Range)wrksheet.Cells["58", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["58", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["58", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["59", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["59", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["59", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["60", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["60", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["60", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["61", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["61", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["61", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["62", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["62", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["62", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["63", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["63", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["63", CM]).Value2 = "";

                        //***********************CP-08***********************

                        ((Excel123.Range)wrksheet.Cells["64", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["64", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["64", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["65", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["65", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["65", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["66", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["66", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["66", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["67", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["67", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["67", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["68", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["68", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["68", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["69", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["69", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["69", CM]).Value2 = "";

                        //***********************************CP-09****************

                        ((Excel123.Range)wrksheet.Cells["70", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["70", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["70", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["71", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["71", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["71", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["72", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["72", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["72", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["73", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["73", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["73", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["74", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["74", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["74", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["75", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["75", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["75", CM]).Value2 = "";

                        //*******************CP-10***************************

                        ((Excel123.Range)wrksheet.Cells["76", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["76", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["76", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["77", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["77", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["77", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["78", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["78", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["78", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["79", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["79", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["79", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["80", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["80", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["80", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["81", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["81", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["81", CM]).Value2 = "";

                        //***************************CD22*******************************

                        ((Excel123.Range)wrksheet.Cells["82", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["82", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["82", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["83", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["83", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["83", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["84", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["84", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["84", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["85", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["85", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["85", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["86", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["86", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["86", CM]).Value2 = "";

                        ((Excel123.Range)wrksheet.Cells["87", CK]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["87", CL]).Value2 = "";
                        ((Excel123.Range)wrksheet.Cells["87", CM]).Value2 = "";


                        //********************* insert data *********************
                        //***************cp19
                        ((Excel123.Range)wrksheet.Cells["1", "BI"]).Value2 = currentTime;
                        ((Excel123.Range)wrksheet.Cells["2", CL]).Value2 = scandate;

                        ((Excel123.Range)wrksheet.Cells["4", CK]).Value2 = sec;
                        ((Excel123.Range)wrksheet.Cells["4", CL]).Value2 = secC1;
                        ((Excel123.Range)wrksheet.Cells["4", CM]).Value2 = secD1;

                        ((Excel123.Range)wrksheet.Cells["5", CK]).Value2 = tot;
                        ((Excel123.Range)wrksheet.Cells["5", CL]).Value2 = mirC1;
                        ((Excel123.Range)wrksheet.Cells["5", CM]).Value2 = mirD1;


                        ((Excel123.Range)wrksheet.Cells["6", CK]).Value2 = Trsno;
                        ((Excel123.Range)wrksheet.Cells["6", CL]).Value2 = TrsnoC1;
                        ((Excel123.Range)wrksheet.Cells["6", CM]).Value2 = TrsnoD1;

                        ((Excel123.Range)wrksheet.Cells["7", CK]).Value2 = Toprns;
                        ((Excel123.Range)wrksheet.Cells["7", CL]).Value2 = ToprnsC1;
                        ((Excel123.Range)wrksheet.Cells["7", CM]).Value2 = ToprnsD1;

                        ((Excel123.Range)wrksheet.Cells["8", CK]).Value2 = st;
                        ((Excel123.Range)wrksheet.Cells["8", CL]).Value2 = stc1;
                        ((Excel123.Range)wrksheet.Cells["8", CM]).Value2 = std1;

                        ((Excel123.Range)wrksheet.Cells["9", CK]).Value2 = lot;
                        ((Excel123.Range)wrksheet.Cells["9", CL]).Value2 = lotC1;
                        ((Excel123.Range)wrksheet.Cells["9", CM]).Value2 = lotD1;

                        //**************************************************************************************
                        //************ cp20*****************

                        ((Excel123.Range)wrksheet.Cells["10", CK]).Value2 = secb2;
                        ((Excel123.Range)wrksheet.Cells["10", CL]).Value2 = secc2;
                        ((Excel123.Range)wrksheet.Cells["10", CM]).Value2 = secD2;


                        ((Excel123.Range)wrksheet.Cells["11", CK]).Value2 = mirb2;
                        ((Excel123.Range)wrksheet.Cells["11", CL]).Value2 = mirc2;
                        ((Excel123.Range)wrksheet.Cells["11", CM]).Value2 = mirD2;

                        ((Excel123.Range)wrksheet.Cells["12", CK]).Value2 = Trsnob2;
                        ((Excel123.Range)wrksheet.Cells["12", CL]).Value2 = Trsnoc2;
                        ((Excel123.Range)wrksheet.Cells["12", CM]).Value2 = TrsnoD2;

                        ((Excel123.Range)wrksheet.Cells["13", CK]).Value2 = Toprnsb2;
                        ((Excel123.Range)wrksheet.Cells["13", CL]).Value2 = Toprnsc2;
                        ((Excel123.Range)wrksheet.Cells["13", CM]).Value2 = ToprnsD2;

                        ((Excel123.Range)wrksheet.Cells["14", CK]).Value2 = stb2;
                        ((Excel123.Range)wrksheet.Cells["14", CL]).Value2 = stc2;
                        ((Excel123.Range)wrksheet.Cells["14", CM]).Value2 = std2;

                        ((Excel123.Range)wrksheet.Cells["15", CK]).Value2 = lotb2;
                        ((Excel123.Range)wrksheet.Cells["15", CL]).Value2 = lotc2;
                        ((Excel123.Range)wrksheet.Cells["15", CM]).Value2 = lotD2;

                        //**************************************************************************************
                        //**************cp21****************
                        ((Excel123.Range)wrksheet.Cells["16", CK]).Value2 = secb3;
                        ((Excel123.Range)wrksheet.Cells["16", CL]).Value2 = secc3;
                        ((Excel123.Range)wrksheet.Cells["16", CM]).Value2 = secD3;

                        ((Excel123.Range)wrksheet.Cells["17", CK]).Value2 = mirb3;
                        ((Excel123.Range)wrksheet.Cells["17", CL]).Value2 = mirc3;
                        ((Excel123.Range)wrksheet.Cells["17", CM]).Value2 = mirD3;

                        ((Excel123.Range)wrksheet.Cells["18", CK]).Value2 = Trsnob3;
                        ((Excel123.Range)wrksheet.Cells["18", CL]).Value2 = Trsnoc3;
                        ((Excel123.Range)wrksheet.Cells["18", CM]).Value2 = TrsnoD3;

                        ((Excel123.Range)wrksheet.Cells["19", CK]).Value2 = Toprnsb3;
                        ((Excel123.Range)wrksheet.Cells["19", CL]).Value2 = Toprnsc3;
                        ((Excel123.Range)wrksheet.Cells["19", CM]).Value2 = ToprnsD3;

                        ((Excel123.Range)wrksheet.Cells["20", CK]).Value2 = stb3;
                        ((Excel123.Range)wrksheet.Cells["20", CL]).Value2 = stc3;
                        ((Excel123.Range)wrksheet.Cells["20", CM]).Value2 = std3;


                        ((Excel123.Range)wrksheet.Cells["21", CK]).Value2 = lotb3;
                        ((Excel123.Range)wrksheet.Cells["21", CL]).Value2 = lotc3;
                        ((Excel123.Range)wrksheet.Cells["21", CM]).Value2 = lotD3;

                        //***********************************************************************************
                        //******************************cp01**********************
                        ((Excel123.Range)wrksheet.Cells["22", CK]).Value2 = secb4;
                        ((Excel123.Range)wrksheet.Cells["22", CL]).Value2 = secc4;
                        ((Excel123.Range)wrksheet.Cells["22", CM]).Value2 = secD4;

                        ((Excel123.Range)wrksheet.Cells["23", CK]).Value2 = mirb4;
                        ((Excel123.Range)wrksheet.Cells["23", CL]).Value2 = mirc4;
                        ((Excel123.Range)wrksheet.Cells["23", CM]).Value2 = mirD4;

                        ((Excel123.Range)wrksheet.Cells["24", CK]).Value2 = Trsnob4;
                        ((Excel123.Range)wrksheet.Cells["24", CL]).Value2 = Trsnoc4;
                        ((Excel123.Range)wrksheet.Cells["24", CM]).Value2 = TrsnoD4;

                        ((Excel123.Range)wrksheet.Cells["25", CK]).Value2 = Toprnsb4;
                        ((Excel123.Range)wrksheet.Cells["25", CL]).Value2 = Toprnsc4;
                        ((Excel123.Range)wrksheet.Cells["25", CM]).Value2 = ToprnsD4;

                        ((Excel123.Range)wrksheet.Cells["26", CK]).Value2 = stb4;
                        ((Excel123.Range)wrksheet.Cells["26", CL]).Value2 = stc4;
                        ((Excel123.Range)wrksheet.Cells["26", CM]).Value2 = std4;


                        ((Excel123.Range)wrksheet.Cells["27", CK]).Value2 = lotb4;
                        ((Excel123.Range)wrksheet.Cells["27", CL]).Value2 = lotc4;
                        ((Excel123.Range)wrksheet.Cells["27", CM]).Value2 = lotD4;

                        //*********************************************************************
                        //********************cp02**********************
                        ((Excel123.Range)wrksheet.Cells["28", CK]).Value2 = secb5;
                        ((Excel123.Range)wrksheet.Cells["28", CL]).Value2 = secc5;
                        ((Excel123.Range)wrksheet.Cells["28", CM]).Value2 = secD5;

                        ((Excel123.Range)wrksheet.Cells["29", CK]).Value2 = mirb5;
                        ((Excel123.Range)wrksheet.Cells["29", CL]).Value2 = mirc5;
                        ((Excel123.Range)wrksheet.Cells["29", CM]).Value2 = mirD5;

                        ((Excel123.Range)wrksheet.Cells["30", CK]).Value2 = Trsnob5;
                        ((Excel123.Range)wrksheet.Cells["30", CL]).Value2 = Trsnoc5;
                        ((Excel123.Range)wrksheet.Cells["30", CM]).Value2 = TrsnoD5;

                        ((Excel123.Range)wrksheet.Cells["31", CK]).Value2 = Toprnsb5;
                        ((Excel123.Range)wrksheet.Cells["31", CL]).Value2 = Toprnsc5;
                        ((Excel123.Range)wrksheet.Cells["31", CM]).Value2 = ToprnsD5;

                        ((Excel123.Range)wrksheet.Cells["32", CK]).Value2 = stb5;
                        ((Excel123.Range)wrksheet.Cells["32", CL]).Value2 = stc5;
                        ((Excel123.Range)wrksheet.Cells["32", CM]).Value2 = std5;


                        ((Excel123.Range)wrksheet.Cells["33", CK]).Value2 = lotb5;
                        ((Excel123.Range)wrksheet.Cells["33", CL]).Value2 = lotc5;
                        ((Excel123.Range)wrksheet.Cells["33", CM]).Value2 = lotD5;

                        //*********************************************************
                        //******************cp03*********************

                        ((Excel123.Range)wrksheet.Cells["34", CK]).Value2 = secb6;
                        ((Excel123.Range)wrksheet.Cells["34", CL]).Value2 = secc6;
                        ((Excel123.Range)wrksheet.Cells["34", CM]).Value2 = secD6;

                        ((Excel123.Range)wrksheet.Cells["35", CK]).Value2 = mirb6;
                        ((Excel123.Range)wrksheet.Cells["35", CL]).Value2 = mirc6;
                        ((Excel123.Range)wrksheet.Cells["35", CM]).Value2 = mirD6;

                        ((Excel123.Range)wrksheet.Cells["36", CK]).Value2 = Trsnob6;
                        ((Excel123.Range)wrksheet.Cells["36", CL]).Value2 = Trsnoc6;
                        ((Excel123.Range)wrksheet.Cells["36", CM]).Value2 = TrsnoD6;

                        ((Excel123.Range)wrksheet.Cells["37", CK]).Value2 = Toprnsb6;
                        ((Excel123.Range)wrksheet.Cells["37", CL]).Value2 = Toprnsc6;
                        ((Excel123.Range)wrksheet.Cells["37", CM]).Value2 = ToprnsD6;

                        ((Excel123.Range)wrksheet.Cells["38", CK]).Value2 = stb6;
                        ((Excel123.Range)wrksheet.Cells["38", CL]).Value2 = stc6;
                        ((Excel123.Range)wrksheet.Cells["38", CM]).Value2 = std6;


                        ((Excel123.Range)wrksheet.Cells["39", CK]).Value2 = lotb6;
                        ((Excel123.Range)wrksheet.Cells["39", CL]).Value2 = lotc6;
                        ((Excel123.Range)wrksheet.Cells["39", CM]).Value2 = lotD6;

                        //*******************CP-04*****************************

                        ((Excel123.Range)wrksheet.Cells["40", CK]).Value2 = secb7;
                        ((Excel123.Range)wrksheet.Cells["40", CL]).Value2 = secc7;
                        ((Excel123.Range)wrksheet.Cells["40", CM]).Value2 = secD7;

                        ((Excel123.Range)wrksheet.Cells["41", CK]).Value2 = mirb7;
                        ((Excel123.Range)wrksheet.Cells["41", CL]).Value2 = mirc7;
                        ((Excel123.Range)wrksheet.Cells["41", CM]).Value2 = mirD7;

                        ((Excel123.Range)wrksheet.Cells["42", CK]).Value2 = Trsnob7;
                        ((Excel123.Range)wrksheet.Cells["42", CL]).Value2 = Trsnoc7;
                        ((Excel123.Range)wrksheet.Cells["42", CM]).Value2 = TrsnoD7;

                        ((Excel123.Range)wrksheet.Cells["43", CK]).Value2 = Toprnsb7;
                        ((Excel123.Range)wrksheet.Cells["43", CL]).Value2 = Toprnsc7;
                        ((Excel123.Range)wrksheet.Cells["43", CM]).Value2 = ToprnsD7;

                        ((Excel123.Range)wrksheet.Cells["44", CK]).Value2 = stb7;
                        ((Excel123.Range)wrksheet.Cells["44", CL]).Value2 = stc7;
                        ((Excel123.Range)wrksheet.Cells["44", CM]).Value2 = std7;


                        ((Excel123.Range)wrksheet.Cells["45", CK]).Value2 = lotb7;
                        ((Excel123.Range)wrksheet.Cells["45", CL]).Value2 = lotc7;
                        ((Excel123.Range)wrksheet.Cells["45", CM]).Value2 = lotD7;

                        //**********************CP-05*************************
                        ((Excel123.Range)wrksheet.Cells["46", CK]).Value2 = secb8;
                        ((Excel123.Range)wrksheet.Cells["46", CL]).Value2 = secc8;
                        ((Excel123.Range)wrksheet.Cells["46", CM]).Value2 = secD8;

                        ((Excel123.Range)wrksheet.Cells["47", CK]).Value2 = mirb8;
                        ((Excel123.Range)wrksheet.Cells["47", CL]).Value2 = mirc8;
                        ((Excel123.Range)wrksheet.Cells["47", CM]).Value2 = mirD8;

                        ((Excel123.Range)wrksheet.Cells["48", CK]).Value2 = Trsnob8;
                        ((Excel123.Range)wrksheet.Cells["48", CL]).Value2 = Trsnoc8;
                        ((Excel123.Range)wrksheet.Cells["48", CM]).Value2 = TrsnoD8;

                        ((Excel123.Range)wrksheet.Cells["49", CK]).Value2 = Toprnsb8;
                        ((Excel123.Range)wrksheet.Cells["49", CL]).Value2 = Toprnsc8;
                        ((Excel123.Range)wrksheet.Cells["49", CM]).Value2 = ToprnsD8;

                        ((Excel123.Range)wrksheet.Cells["50", CK]).Value2 = stb8;
                        ((Excel123.Range)wrksheet.Cells["50", CL]).Value2 = stc8;
                        ((Excel123.Range)wrksheet.Cells["50", CM]).Value2 = std8;


                        ((Excel123.Range)wrksheet.Cells["51", CK]).Value2 = lotb8;
                        ((Excel123.Range)wrksheet.Cells["51", CL]).Value2 = lotc8;
                        ((Excel123.Range)wrksheet.Cells["51", CM]).Value2 = lotD8;

                        //****************************CP-06*****************************************

                        ((Excel123.Range)wrksheet.Cells["52", CK]).Value2 = secb9;
                        ((Excel123.Range)wrksheet.Cells["52", CL]).Value2 = secc9;
                        ((Excel123.Range)wrksheet.Cells["52", CM]).Value2 = secD9;

                        ((Excel123.Range)wrksheet.Cells["53", CK]).Value2 = mirb9;
                        ((Excel123.Range)wrksheet.Cells["53", CL]).Value2 = mirc9;
                        ((Excel123.Range)wrksheet.Cells["53", CM]).Value2 = mirD9;

                        ((Excel123.Range)wrksheet.Cells["54", CK]).Value2 = Trsnob9;
                        ((Excel123.Range)wrksheet.Cells["54", CL]).Value2 = Trsnoc9;
                        ((Excel123.Range)wrksheet.Cells["54", CM]).Value2 = TrsnoD9;

                        ((Excel123.Range)wrksheet.Cells["55", CK]).Value2 = Toprnsb9;
                        ((Excel123.Range)wrksheet.Cells["55", CL]).Value2 = Toprnsc9;
                        ((Excel123.Range)wrksheet.Cells["55", CM]).Value2 = ToprnsD9;

                        ((Excel123.Range)wrksheet.Cells["56", CK]).Value2 = stb9;
                        ((Excel123.Range)wrksheet.Cells["56", CL]).Value2 = stc9;
                        ((Excel123.Range)wrksheet.Cells["56", CM]).Value2 = std9;


                        ((Excel123.Range)wrksheet.Cells["57", CK]).Value2 = lotb9;
                        ((Excel123.Range)wrksheet.Cells["57", CL]).Value2 = lotc9;
                        ((Excel123.Range)wrksheet.Cells["57", CM]).Value2 = lotD9;

                        //******************CP-07*****************************************

                        ((Excel123.Range)wrksheet.Cells["58", CK]).Value2 = secb10;
                        ((Excel123.Range)wrksheet.Cells["58", CL]).Value2 = secc10;
                        ((Excel123.Range)wrksheet.Cells["58", CM]).Value2 = secD10;

                        ((Excel123.Range)wrksheet.Cells["59", CK]).Value2 = mirb10;
                        ((Excel123.Range)wrksheet.Cells["59", CL]).Value2 = mirc10;
                        ((Excel123.Range)wrksheet.Cells["59", CM]).Value2 = mirD10;

                        ((Excel123.Range)wrksheet.Cells["60", CK]).Value2 = Trsnob10;
                        ((Excel123.Range)wrksheet.Cells["60", CL]).Value2 = Trsnoc10;
                        ((Excel123.Range)wrksheet.Cells["60", CM]).Value2 = TrsnoD10;

                        ((Excel123.Range)wrksheet.Cells["61", CK]).Value2 = Toprnsb10;
                        ((Excel123.Range)wrksheet.Cells["61", CL]).Value2 = Toprnsc10;
                        ((Excel123.Range)wrksheet.Cells["61", CM]).Value2 = ToprnsD10;

                        ((Excel123.Range)wrksheet.Cells["62", CK]).Value2 = stb10;
                        ((Excel123.Range)wrksheet.Cells["62", CL]).Value2 = stc10;
                        ((Excel123.Range)wrksheet.Cells["62", CM]).Value2 = std10;


                        ((Excel123.Range)wrksheet.Cells["63", CK]).Value2 = lotb10;
                        ((Excel123.Range)wrksheet.Cells["63", CL]).Value2 = lotc10;
                        ((Excel123.Range)wrksheet.Cells["63", CM]).Value2 = lotD10;

                        //***********************CP-08******************************

                        ((Excel123.Range)wrksheet.Cells["64", CK]).Value2 = secb11;
                        ((Excel123.Range)wrksheet.Cells["64", CL]).Value2 = secc11;
                        ((Excel123.Range)wrksheet.Cells["64", CM]).Value2 = secD11;

                        ((Excel123.Range)wrksheet.Cells["65", CK]).Value2 = mirb11;
                        ((Excel123.Range)wrksheet.Cells["65", CL]).Value2 = mirc11;
                        ((Excel123.Range)wrksheet.Cells["65", CM]).Value2 = mirD11;

                        ((Excel123.Range)wrksheet.Cells["66", CK]).Value2 = Trsnob11;
                        ((Excel123.Range)wrksheet.Cells["66", CL]).Value2 = Trsnoc11;
                        ((Excel123.Range)wrksheet.Cells["66", CM]).Value2 = TrsnoD11;

                        ((Excel123.Range)wrksheet.Cells["67", CK]).Value2 = Toprnsb11;
                        ((Excel123.Range)wrksheet.Cells["67", CL]).Value2 = Toprnsc11;
                        ((Excel123.Range)wrksheet.Cells["67", CM]).Value2 = ToprnsD11;

                        ((Excel123.Range)wrksheet.Cells["68", CK]).Value2 = stb11;
                        ((Excel123.Range)wrksheet.Cells["68", CL]).Value2 = stc11;
                        ((Excel123.Range)wrksheet.Cells["68", CM]).Value2 = std11;


                        ((Excel123.Range)wrksheet.Cells["69", CK]).Value2 = lotb11;
                        ((Excel123.Range)wrksheet.Cells["69", CL]).Value2 = lotc11;
                        ((Excel123.Range)wrksheet.Cells["69", CM]).Value2 = lotD11;

                        //***********************CP-09******************************

                        ((Excel123.Range)wrksheet.Cells["70", CK]).Value2 = secb12;
                        ((Excel123.Range)wrksheet.Cells["70", CL]).Value2 = secc12;
                        ((Excel123.Range)wrksheet.Cells["70", CM]).Value2 = secD12;

                        ((Excel123.Range)wrksheet.Cells["71", CK]).Value2 = mirb12;
                        ((Excel123.Range)wrksheet.Cells["71", CL]).Value2 = mirc12;
                        ((Excel123.Range)wrksheet.Cells["71", CM]).Value2 = mirD12;

                        ((Excel123.Range)wrksheet.Cells["72", CK]).Value2 = Trsnob12;
                        ((Excel123.Range)wrksheet.Cells["72", CL]).Value2 = Trsnoc12;
                        ((Excel123.Range)wrksheet.Cells["72", CM]).Value2 = TrsnoD12;

                        ((Excel123.Range)wrksheet.Cells["73", CK]).Value2 = Toprnsb12;
                        ((Excel123.Range)wrksheet.Cells["73", CL]).Value2 = Toprnsc12;
                        ((Excel123.Range)wrksheet.Cells["73", CM]).Value2 = ToprnsD12;

                        ((Excel123.Range)wrksheet.Cells["74", CK]).Value2 = stb12;
                        ((Excel123.Range)wrksheet.Cells["74", CL]).Value2 = stc12;
                        ((Excel123.Range)wrksheet.Cells["74", CM]).Value2 = std12;


                        ((Excel123.Range)wrksheet.Cells["75", CK]).Value2 = lotb12;
                        ((Excel123.Range)wrksheet.Cells["75", CL]).Value2 = lotc12;
                        ((Excel123.Range)wrksheet.Cells["75", CM]).Value2 = lotD12;

                        //*********************CP-10*********************************

                        ((Excel123.Range)wrksheet.Cells["76", CK]).Value2 = secb13;
                        ((Excel123.Range)wrksheet.Cells["76", CL]).Value2 = secc13;
                        ((Excel123.Range)wrksheet.Cells["76", CM]).Value2 = secD13;

                        ((Excel123.Range)wrksheet.Cells["77", CK]).Value2 = mirb13;
                        ((Excel123.Range)wrksheet.Cells["77", CL]).Value2 = mirc13;
                        ((Excel123.Range)wrksheet.Cells["77", CM]).Value2 = mirD13;

                        ((Excel123.Range)wrksheet.Cells["78", CK]).Value2 = Trsnob13;
                        ((Excel123.Range)wrksheet.Cells["78", CL]).Value2 = Trsnoc13;
                        ((Excel123.Range)wrksheet.Cells["78", CM]).Value2 = TrsnoD13;

                        ((Excel123.Range)wrksheet.Cells["79", CK]).Value2 = Toprnsb13;
                        ((Excel123.Range)wrksheet.Cells["79", CL]).Value2 = Toprnsc13;
                        ((Excel123.Range)wrksheet.Cells["79", CM]).Value2 = ToprnsD13;

                        ((Excel123.Range)wrksheet.Cells["80", CK]).Value2 = stb13;
                        ((Excel123.Range)wrksheet.Cells["80", CL]).Value2 = stc13;
                        ((Excel123.Range)wrksheet.Cells["80", CM]).Value2 = std13;


                        ((Excel123.Range)wrksheet.Cells["81", CK]).Value2 = lotb13;
                        ((Excel123.Range)wrksheet.Cells["81", CL]).Value2 = lotc13;
                        ((Excel123.Range)wrksheet.Cells["81", CM]).Value2 = lotD13;


                        //**********************CD22********************************


                        ((Excel123.Range)wrksheet.Cells["82", CK]).Value2 = secb14;
                        ((Excel123.Range)wrksheet.Cells["82", CL]).Value2 = secc14;
                        ((Excel123.Range)wrksheet.Cells["82", CM]).Value2 = secD14;

                        ((Excel123.Range)wrksheet.Cells["83", CK]).Value2 = mirb14;
                        ((Excel123.Range)wrksheet.Cells["83", CL]).Value2 = mirc14;
                        ((Excel123.Range)wrksheet.Cells["83", CM]).Value2 = mirD14;

                        ((Excel123.Range)wrksheet.Cells["84", CK]).Value2 = Trsnob14;
                        ((Excel123.Range)wrksheet.Cells["84", CL]).Value2 = Trsnoc14;
                        ((Excel123.Range)wrksheet.Cells["84", CM]).Value2 = TrsnoD14;

                        ((Excel123.Range)wrksheet.Cells["85", CK]).Value2 = Toprnsb14;
                        ((Excel123.Range)wrksheet.Cells["85", CL]).Value2 = Toprnsc14;
                        ((Excel123.Range)wrksheet.Cells["85", CM]).Value2 = ToprnsD14;

                        ((Excel123.Range)wrksheet.Cells["86", CK]).Value2 = stb14;
                        ((Excel123.Range)wrksheet.Cells["86", CL]).Value2 = stc14;
                        ((Excel123.Range)wrksheet.Cells["86", CM]).Value2 = std14;


                        ((Excel123.Range)wrksheet.Cells["87", CK]).Value2 = lotb14;
                        ((Excel123.Range)wrksheet.Cells["87", CL]).Value2 = lotc14;
                        ((Excel123.Range)wrksheet.Cells["87", CM]).Value2 = lotD14;
 
                    }
                    #endregion
                }

            }
            catch (Exception ex)
            {
                return Ok(ex.ToString());
            }
            finally
            {

                workbook.Save();
                workbook.Close(0, 0, 0);
                excelApp.Quit();
                Conn.Close();
                System.Diagnostics.Process.Start(exc);
            }
            return Ok("Ok");
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
