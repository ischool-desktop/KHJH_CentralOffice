using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DesktopLib;
using System.Xml.Linq;
using FISCA.DSA;
using FISCA;

namespace KHJHCentralOffice
{
    public partial class StudentExchangeReset : DetailContentImproved
    {
        private XElement Data { get; set; }

        public StudentExchangeReset()
        {
            InitializeComponent();
            Group = "學生轉出狀態重置";
            dgvTransfers.AutoGenerateColumns = false;
        }

        private void StudentExchangeReset_Load(object sender, EventArgs e)
        {
            InitDetailContent();
        }

        protected override void OnPrimaryKeyChangedAsync()
        {
            ConnectionHelper ch = ConnectionHelper.GetConnection(PrimaryKey);

            string sql = "select uid,studentname \"name\",transfertoken token,status,transfertarget target from $st_transferout order by last_update desc";

            XElement req = new XElement("Request",
                new XElement("SQL", sql));

            Envelope rsp = ch.CallService("UDTService.DML.Query", new Envelope(new XStringHolder(req)));

            Data = XElement.Parse(rsp.BodyContent.XmlString);

        }

        protected override void OnPrimaryKeyChangedComplete(Exception error)
        {
            dgvTransfers.Rows.Clear();
            if (error == null)
            {
                Dictionary<int, string> indexToFIeldName = GetIndexMapping(Data);

                foreach (XElement record in Data.Elements("Record"))
                {
                    Dictionary<string, string> values = new Dictionary<string, string>();
                    foreach (XElement column in record.Elements("Column"))
                    {
                        int index = int.Parse(column.AttributeText("Index"));
                        string value = (column.Value);

                        values.Add(indexToFIeldName[index], value);
                    }

                    DataGridViewRow row = new DataGridViewRow();
                    row.Tag = values["uid"];
                    row.CreateCells(dgvTransfers, values["name"], values["token"], values["status"], values["target"]);
                    dgvTransfers.Rows.Add(row);
                }
            }
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            try
            {
                ConnectionHelper ch = ConnectionHelper.GetConnection(PrimaryKey);

                DataGridViewRow row = dgvTransfers.SelectedRows[0];
                string uid = row.Tag.ToString();

                string sql = "update $st_transferout set status='1', transfertarget=null,accepttoken=null where uid='" + uid + "';";

                XElement req = new XElement("Request",
                    new XElement("Command", sql));

                Envelope rsp = ch.CallService("UDTService.DML.Command", new Envelope(new XStringHolder(req)));

                OnPrimaryKeyChanged(EventArgs.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private static Dictionary<int, string> GetIndexMapping(XElement xElement)
        {
            XElement metadata = xElement.Element("Metadata");
            Dictionary<int, string> result = new Dictionary<int, string>();

            int offset = 0;
            foreach (XElement column in metadata.Elements("Column"))
            {
                string index = column.AttributeText("Index");
                string value = column.AttributeText("Field");

                result.Add(int.Parse(index), value);

                offset++;
            }

            return result;
        }
    }
}
