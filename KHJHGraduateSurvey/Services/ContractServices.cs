using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Xml.Linq;
using FISCA.Authentication;
using FISCA.DSAClient;

namespace JH_KH_GraduateSurvey
{
    /// <summary>
    /// 直接存取Contract服務
    /// </summary>
    public class ContractService
    {
        private static string DefaultContractName = "centraloffice";
        private static Dictionary<string, Connection> mConnections = null;
        private static List<Timer> mTimers = null;

        /// <summary>
        /// 關閉所有連線，OK
        /// </summary>
        public static void CloseConnection()
        {
            mConnections = null;
            mTimers = null;
        }

        /// <summary>
        /// 初始化連線，OK
        /// </summary>
        public static void InitialConnection()
        {
            GetConnection(FISCA.Authentication.DSAServices.AccessPoint);

            GetConnection(DSAServices.GreeningAccessPoint, "user");
        }

        /// <summary>
        /// 用Passport連線至AccessPoint，並傳回對應的連線資訊。OK
        /// </summary>
        /// <param name="AccessPoint"></param>
        /// <returns></returns>
        public static Tuple<Connection, string> GetConnection(string AccessPoint)
        {
            return GetConnection(AccessPoint, DefaultContractName);
        }

        /// <summary>
        /// 用Passport連線至AccessPoint，並傳回對應的連線資訊；並且會放到連線區中，之後可重覆使用。OK
        /// </summary>
        /// <param name="AccessPoint"></param>
        /// <returns></returns>
        public static Tuple<Connection, string> GetConnection(string AccessPoint, string ContractName)
        {
            if (mConnections == null)
                mConnections = new Dictionary<string, Connection>();

            if (mTimers == null)
                mTimers = new List<Timer>();

            if (mConnections.ContainsKey(AccessPoint))
                return new Tuple<Connection, string>(mConnections[AccessPoint], string.Empty);

            try
            {
                #region Connection一開始會用Passport連線，之後強制使用Session連線，並定期送Request確保Session不會過期
                Connection vConnection = new Connection();
                vConnection.EnableSession = true;
                vConnection.Connect(AccessPoint, ContractName, FISCA.Authentication.DSAServices.AccessPoint, FISCA.Authentication.DSAServices.AccessPoint);

                //Timer mTimer = new Timer(x =>
                //{
                //    vConnection.SendRequest("DS.Base.Connect", new Envelope());
                //}, null, 9000 * 60, 9000 * 60);

                //mTimers.Add(mTimer);

                mConnections.Add(AccessPoint, vConnection);
                #endregion

                return new Tuple<Connection, string>(vConnection, string.Empty);
            }
            catch (Exception ve)
            {
                return new Tuple<Connection, string>(null, "無法連線至『" + AccessPoint + "』主機" + System.Environment.NewLine + "訊息：『" + ve.Message + "』");
            }
        }

        /// <summary>
        /// 送出文件，OK
        /// </summary>
        /// <param name="Connection"></param>
        /// <param name="ServiceName"></param>
        /// <param name="RequestElement"></param>
        /// <returns></returns>
        private static XElement SendRequest(Connection Conn, string ServiceName, XElement RequestElement)
        {
            try
            {
                Envelope Request = new Envelope();

                Request.Body = new XmlStringHolder(RequestElement.ToString());

                Envelope Response = Conn.CallService(ServiceName, Request);

                XElement Element = XElement.Load(new StringReader(Response.Body.XmlString));

                return Element;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 取得學校資訊
        /// </summary>
        /// <param name="Connection"></param>
        /// <param name="DSNS"></param>
        /// <returns></returns>
        public static XElement GetSchool(
            Connection Connection,
            string DSNS)
        {
            try
            {
                XElement Request = new XElement("Request");

                //<Request>
                //    <Field>
                //        <All></All>
                //    </Field>
                //    <Condition></Condition>
                //</Request>

                XElement elmCondition = new XElement("Condition");
                elmCondition.Add(new XElement("Dsns", DSNS));

                Request.Add(elmCondition);

                XElement Response = SendRequest(Connection, "_.SelectSchool", Request);

                return Response;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 上傳畢業生進路
        /// </summary>
        /// <param name="Connection"></param>
        /// <param name="SchoolID">學校系統編號</param>
        /// <param name="SurveyYear">學年度</param>
        /// <param name="mergeKeyValue">值</param>
        /// <returns></returns>
        public static XElement UploadApproach(
            Connection Connection,
            string SchoolID,
            string SurveyYear,
            Dictionary<string, object> mergeKeyValue)
        {
            try
             {
                XElement Request = new XElement("Request");

                //<Request>
                //    <Field>
                //        <!--以下為非必要欄位-->
                //        <SchoolID>1730</SchoolID>
                //        <SurveyYear>102</SurveyYear>
                //        <Content>xxx</Content>
                //    </Field>
                //</Request>

                XElement elmContent = new XElement("Content");
                XElement elmField = new XElement("Field");
                elmField.Add(new XElement("SchoolID", SchoolID));
                elmField.Add(new XElement("SurveyYear", SurveyYear));

                foreach(KeyValuePair<string, object> kv in mergeKeyValue)
                {
                    if (kv.Key.ToLower() == "UnApproachStudents".ToLower())
                        elmContent.Add(mergeKeyValue["UnApproachStudents"]);
                    else
                        elmContent.Add(new XElement(kv.Key.Replace("/", "-"), kv.Value));
                }
                #region 升學或就業情形
                //elmContent.Add(new XElement("A1",mergeKeyValue["A1"]));
                //elmContent.Add(new XElement("B1", mergeKeyValue["B1"]));
                //elmContent.Add(new XElement("C1", mergeKeyValue["C1"]));
                //elmContent.Add(new XElement("D1", mergeKeyValue["D1"]));

                //elmContent.Add(new XElement("B1-A1", mergeKeyValue["B1/A1"]));
                //elmContent.Add(new XElement("C1-A1", mergeKeyValue["C1/A1"]));
                //elmContent.Add(new XElement("D1-A1", mergeKeyValue["D1/A1"]));
                #endregion

                #region 就讀學校
                //elmContent.Add(new XElement("A2", mergeKeyValue["A2"]));
                //elmContent.Add(new XElement("B2", mergeKeyValue["B2"]));
                //elmContent.Add(new XElement("C2", mergeKeyValue["C2"]));
                //elmContent.Add(new XElement("D2", mergeKeyValue["D2"]));

                //elmContent.Add(new XElement("E2", mergeKeyValue["E2"]));
                //elmContent.Add(new XElement("F2", mergeKeyValue["F2"]));
                //elmContent.Add(new XElement("G2", mergeKeyValue["G2"]));
                //elmContent.Add(new XElement("H2", mergeKeyValue["H2"]));
                //elmContent.Add(new XElement("I2", mergeKeyValue["I2"]));

                //elmContent.Add(new XElement("B2-A2", mergeKeyValue["B2/A2"]));
                //elmContent.Add(new XElement("C2-A2", mergeKeyValue["C2/A2"]));
                //elmContent.Add(new XElement("D2-A2", mergeKeyValue["D2/A2"]));
                //elmContent.Add(new XElement("E2-A2", mergeKeyValue["E2/A2"]));
                //elmContent.Add(new XElement("F2-A2", mergeKeyValue["F2/A2"]));

                //elmContent.Add(new XElement("G2-A2", mergeKeyValue["G2/A2"]));
                //elmContent.Add(new XElement("H2-A2", mergeKeyValue["H2/A2"]));
                //elmContent.Add(new XElement("I2-A2", mergeKeyValue["I2/A2"]));
                #endregion

                #region 學制別
                //elmContent.Add(new XElement("A3", mergeKeyValue["A3"]));

                //elmContent.Add(new XElement("B3", mergeKeyValue["B3"]));
                //elmContent.Add(new XElement("C3", mergeKeyValue["C3"]));
                //elmContent.Add(new XElement("D3", mergeKeyValue["D3"]));
                //elmContent.Add(new XElement("E3", mergeKeyValue["E3"]));
                //elmContent.Add(new XElement("F3", mergeKeyValue["F3"]));
                //elmContent.Add(new XElement("G3", mergeKeyValue["G3"]));
                //elmContent.Add(new XElement("H3", mergeKeyValue["H3"]));
                //elmContent.Add(new XElement("I3", mergeKeyValue["I3"]));
                //elmContent.Add(new XElement("J3", mergeKeyValue["J3"]));


                //elmContent.Add(new XElement("B3-A3", mergeKeyValue["B3/A3"]));
                //elmContent.Add(new XElement("C3-A3", mergeKeyValue["C3/A3"]));
                //elmContent.Add(new XElement("D3-A3", mergeKeyValue["D3/A3"]));
                //elmContent.Add(new XElement("E3-A3", mergeKeyValue["E3/A3"]));
                //elmContent.Add(new XElement("F3-A3", mergeKeyValue["F3/A3"]));
                //elmContent.Add(new XElement("G3-A3", mergeKeyValue["G3/A3"]));
                //elmContent.Add(new XElement("H3-A3", mergeKeyValue["H3/A3"]));
                //elmContent.Add(new XElement("I3-A3", mergeKeyValue["I3/A3"]));
                //elmContent.Add(new XElement("J3-A3", mergeKeyValue["J3/A3"]));
                #endregion

                #region 入學方式
                //elmContent.Add(new XElement("A4", mergeKeyValue["A4"]));
                //elmContent.Add(new XElement("B4", mergeKeyValue["B4"]));
                //elmContent.Add(new XElement("C4", mergeKeyValue["C4"]));
                //elmContent.Add(new XElement("D4", mergeKeyValue["D4"]));
                //elmContent.Add(new XElement("E4", mergeKeyValue["E4"]));

                //elmContent.Add(new XElement("F4", mergeKeyValue["F4"]));
                //elmContent.Add(new XElement("G4", mergeKeyValue["G4"]));
                //elmContent.Add(new XElement("H4", mergeKeyValue["H4"]));
                //elmContent.Add(new XElement("I4", mergeKeyValue["I4"]));
                //elmContent.Add(new XElement("J4", mergeKeyValue["J4"]));

                //elmContent.Add(new XElement("K4", mergeKeyValue["K4"]));
                //elmContent.Add(new XElement("L4", mergeKeyValue["L4"]));
                //elmContent.Add(new XElement("M4", mergeKeyValue["M4"]));
                //elmContent.Add(new XElement("N4", mergeKeyValue["N4"]));
                //elmContent.Add(new XElement("O4", mergeKeyValue["O4"]));
                //elmContent.Add(new XElement("P4", mergeKeyValue["P4"]));
                //elmContent.Add(new XElement("Q4", mergeKeyValue["Q4"]));
                //elmContent.Add(new XElement("R4", mergeKeyValue["R4"]));
                //elmContent.Add(new XElement("S4", mergeKeyValue["S4"]));


                //elmContent.Add(new XElement("B4-A4", mergeKeyValue["B4/A4"]));
                //elmContent.Add(new XElement("C4-A4", mergeKeyValue["C4/A4"]));
                //elmContent.Add(new XElement("D4-A4", mergeKeyValue["D4/A4"]));
                //elmContent.Add(new XElement("E4-A4", mergeKeyValue["E4/A4"]));
                //elmContent.Add(new XElement("F4-A4", mergeKeyValue["F4/A4"]));
                //elmContent.Add(new XElement("G4-A4", mergeKeyValue["G4/A4"]));
                //elmContent.Add(new XElement("H4-A4", mergeKeyValue["H4/A4"]));
                //elmContent.Add(new XElement("I4-A4", mergeKeyValue["I4/A4"]));
                //elmContent.Add(new XElement("J4-A4", mergeKeyValue["J4/A4"]));

                //elmContent.Add(new XElement("K4-A4", mergeKeyValue["K4/A4"]));
                //elmContent.Add(new XElement("L4-A4", mergeKeyValue["L4/A4"]));
                //elmContent.Add(new XElement("M4-A4", mergeKeyValue["M4/A4"]));
                //elmContent.Add(new XElement("N4-A4", mergeKeyValue["N4/A4"]));
                //elmContent.Add(new XElement("O4-A4", mergeKeyValue["O4/A4"]));
                //elmContent.Add(new XElement("P4-A4", mergeKeyValue["P4/A4"]));
                //elmContent.Add(new XElement("Q4-A4", mergeKeyValue["Q4/A4"]));
                //elmContent.Add(new XElement("R4-A4", mergeKeyValue["R4/A4"]));
                //elmContent.Add(new XElement("S4-A4", mergeKeyValue["S4/A4"]));
                #endregion

                #region  未升學未就業動向
                //if (mergeKeyValue.ContainsKey("UnApproachStudents"))
                //    elmContent.Add(mergeKeyValue["UnApproachStudents"]);
                //elmContent.Add(new XElement("B5", mergeKeyValue["B5"]));
                //elmContent.Add(new XElement("C5", mergeKeyValue["C5"]));
                //elmContent.Add(new XElement("D5", mergeKeyValue["D5"]));
                //elmContent.Add(new XElement("E5", mergeKeyValue["E5"]));
                //elmContent.Add(new XElement("F5", mergeKeyValue["F5"]));
                //elmContent.Add(new XElement("G5", mergeKeyValue["G5"]));

                //elmContent.Add(new XElement("B5-A5", mergeKeyValue["B5/A5"]));
                //elmContent.Add(new XElement("C5-A5", mergeKeyValue["C5/A5"]));
                //elmContent.Add(new XElement("D5-A5", mergeKeyValue["D5/A5"]));

                //elmContent.Add(new XElement("E5-A5", mergeKeyValue["E5/A5"]));
                //elmContent.Add(new XElement("F5-A5", mergeKeyValue["F5/A5"]));
                //elmContent.Add(new XElement("G5-A5", mergeKeyValue["G5/A5"]));
                #endregion

                string strContent = elmContent.ToString();

                elmField.Add(new XElement("Content", strContent));

                Request.Add(elmField);

                XElement Response = SendRequest(Connection, "_.UploadApproachStatistics", Request);

                return Response;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /// <summary>
        /// 取得開放時間
        /// </summary>
        /// <param name="Connection"></param>
        /// <returns></returns>
        public static XElement GetSurveyYears(Connection Connection)
        {
            try
            {
                XElement Request = new XElement("Request");
                XElement Response = SendRequest(Connection, "_.GetSurveyYears", Request);

                return Response;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /// <summary>
        /// 取得開放時間
        /// </summary>
        /// <param name="Connection"></param>
        /// <returns></returns>
        public static XElement GetOpenDate(Connection Connection)
        {
            try
            {
                XElement Request = new XElement("Request");
                XElement Response = SendRequest(Connection, "_.GetOpenDate", Request);

                return Response;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /// <summary>
        /// 取得開放時間
        /// </summary>
        /// <param name="Connection"></param>
        /// <param name="SchoolYear"></param>
        /// <returns></returns>
        public static XElement GetOpenDateTime(
            Connection Connection,
            string SchoolYear)
        {
            try
            {
                XElement Request = new XElement("Request");

                //<Request>
                //    <Field>
                //        <All></All>
                //    </Field>
                //    <Condition></Condition>
                //</Request>

                XElement elmCondition = new XElement("Condition");
                elmCondition.Add(new XElement("SurveyYear", SchoolYear));

                Request.Add(elmCondition);

                XElement Response = SendRequest(Connection, "_.SelectOpenDate", Request);

                return Response;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}