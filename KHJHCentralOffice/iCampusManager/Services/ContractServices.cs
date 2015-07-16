using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Xml.Linq;
using FISCA.Authentication;
using FISCA.DSAClient;

namespace KHJHCentralOffice
{
    /// <summary>
    /// 直接存取Contract服務
    /// </summary>
    public class CentralOffice
    {
        private static string DefaultContractName = "CalendarService";
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
                vConnection.Connect(AccessPoint, ContractName, FISCA.Authentication.DSAServices.PassportToken);

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
    }
}