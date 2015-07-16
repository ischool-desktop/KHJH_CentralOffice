using System;
using System.Collections.Generic;
using System.Xml.Linq;
using DesktopLib;
using FISCA;
using FISCA.DSA;
using FISCA.UDT;

namespace KHJHCentralOffice
{
    public partial class NetworkItem : DetailContentImproved
    {
        private static string LogApplication = "http://dsa.ischool.com.tw/counter/shared";

        private static string LogContract = "__ischool_log_query";

        //private Chart MainChart { get; set; }

        private Connection Conn { get; set; }

        private School SchoolData { get; set; }

        public NetworkItem()
        {
            InitializeComponent();
            Group = "使用量";
        }

        protected override void OnInitializeAsync()
        {
            try
            {
                Conn = new Connection();
                Conn.Connect(LogApplication, LogContract, "logquery", "logquery1234");
            }
            catch (Exception ex)
            {
                RTOut.WriteError(ex);
            }
        }

        protected override void OnPrimaryKeyChangedAsync()
        {
            AccessHelper access = new AccessHelper();
            List<School> schools = access.Select<School>(string.Format("uid='{0}'", PrimaryKey));

            if (schools.Count > 0)
                SchoolData = schools[0];
            else
                throw new Exception("查不到學校資料：" + PrimaryKey);

            HashSet<string> tables = GetTables();

//            Chart mc = new Chart();
//            mc.ChartArea.AxisY2.IsMajorGridVisible = false;
//            mc.ChartArea.AxisY.IsMajorGridVisible = false;
//            mc.ChartArea.AxisX.IsMajorGridVisible = false;
//            mc.ChartArea.AxisX2.IsMajorGridVisible = false;
//            mc.ChartArea.AxisX.DefaultLabel.Format = "dd";
//            mc.ChartArea.AxisY2.AxisLabels.IsDataPointNameVisible = false;

//            Series serCount = new Series() { ChartType = ChartType.Line };
//            serCount.Name = "取樣數";
//            serCount.DefaultDataPoint.IsLabelVisible = true;
//            serCount.IsPrimaryAxisY = false;

//            Series serTraffic = new Series() { ChartType = ChartType.Bar };
//            serTraffic.Name = "流量(MB)";
//            serTraffic.IsPrimaryAxisY = true;

//            List<DateTime> times = new List<DateTime>();
//            for (int i = 8; i >= 0; i--)
//            {
//                times.Add(DateTime.Now.AddDays(-i));
//            }

//            foreach (DateTime dt in times)
//            {
//                string tableName = string.Format("log_client_{0:0000}_{1:00}_{2:00}", dt.Year, dt.Month, dt.Day);

//                if (!tables.Contains(tableName))
//                    continue;

//                string cmd = @"select count(*) count,sum(bts+btr) traffic,trunc(stddev(spend)) spend,trunc(stddev(server_spend)) sspend,trunc(stddev(spend-server_spend)) network_spend
//                from {0}
//                where app in ('{1}')
//                group by app";

//                cmd = string.Format(cmd, tableName, SchoolData.DSNS);

//                XElement records = Query(cmd);
//                foreach (XElement record in records.Elements("Record"))
//                {
//                    int count = record.ElementInt("count", -1);
//                    int totalSpend = record.ElementInt("spend", -1);
//                    int serverSpend = record.ElementInt("sspend", -1);
//                    int networkSpend = record.ElementInt("network_spend", -1);
//                    int traffic = record.ElementInt("traffic", -1);

//                    serTraffic.DataPoints.Add(new DataPoint(dt, traffic / 1024 / 1024));
//                    serCount.DataPoints.Add(new DataPoint(dt, count));
//                }
//            }
//            mc.SeriesCollection.Add(serTraffic);
//            mc.SeriesCollection.Add(serCount);

//            //Series s0 = new Series();
//            //s0.IsPrimaryAxisY = true;
//            ////mc.ChartArea.AxisY.DefaultLabel.Color = s0.DefaultDataPoint.Color = s0.DefaultDataPoint.BorderColor = Color.Green;
//            //s0.ChartType = ChartType.Line;
//            //for (int i = 1; i < 8; i++)
//            //    s0.DataPoints.Add(new DataPoint(i, i * i));
//            //mc.SeriesCollection.Add(s0);

//            //Series s1 = new Series();
//            //s1.IsPrimaryAxisY = false;
//            ////mc.ChartArea.AxisY2.DefaultLabel.Color = s1.DefaultDataPoint.Color = Color.Blue;
//            //s1.ChartType = ChartType.Bar;
//            //for (int i = 1; i < 8; i++)
//            //    s1.DataPoints.Add(new DataPoint(i, i * i * i));
//            //mc.SeriesCollection.Add(s1);

//            MainChart = mc;
        }

        private HashSet<string> GetTables()
        {
            string cmd = @"select table_name
                from information_schema.tables
                where table_schema='public'";

            HashSet<string> result = new HashSet<string>();
            XElement tables = Query(cmd);
            foreach (XElement table in tables.Elements("Record"))
            {
                result.Add(table.ElementText("table_name"));
            }

            return result;
        }

        protected override void OnPrimaryKeyChangedComplete(Exception error)
        {
            if (error != null)
                throw error;

            //MainChart.Width = DisplayArea.Width;
            //MainChart.Height = DisplayArea.Height;
            //DisplayArea.Image = MainChart.GetChartImage();
        }

        private void NetworkItem_Load(object sender, EventArgs e)
        {
            InitDetailContent();
        }

        private XElement Query(string cmd)
        {
            XElement req = new XElement("Request");
            req.Add(new XElement("Command", cmd));

            Envelope rsp = Conn.SendRequest("Query", new Envelope(new XStringHolder(req)));

            return XElement.Parse(rsp.BodyContent.XmlString);
        }
    }
}
