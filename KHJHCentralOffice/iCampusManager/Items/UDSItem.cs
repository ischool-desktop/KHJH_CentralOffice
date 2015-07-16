using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using FISCA.DSA;
using FISCA;
using System.Xml.XPath;
using DesktopLib;

namespace KHJHCentralOffice
{
    public partial class UDSItem : DetailContentImproved
    {
        private XElement UDS { get; set; }

        public UDSItem()
        {
            InitializeComponent();
            Group = "UDS";
        }

        protected override void OnPrimaryKeyChangedAsync()
        {
            ConnectionHelper conn = ConnectionHelper.GetConnection(PrimaryKey);

            Envelope rsp = conn.CallService("UDSManagerService.ListContracts", new Envelope());

            UDS = XElement.Parse(rsp.BodyContent.XmlString);
        }

        protected override void OnPrimaryKeyChangedComplete(Exception error)
        {
            if (error != null)
            {
                RTOut.WriteError(error);
                throw error;
            }

            dgvUDS.Rows.Clear();

            List<string> udsNames = new List<string>();

            foreach (XElement each in UDS.Elements("Contract"))
            {
                udsNames.Add(each.AttributeText("Name"));
            }

            udsNames.Sort();

            foreach (string each in udsNames)
            {
                DataGridViewRow row = new DataGridViewRow();

                row.CreateCells(dgvUDS, each);

                dgvUDS.Rows.Add(row);
            }
        }

        private void btnInstall_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Filter = "Contract 定議 (*.csml;*.cml)|*.csml;*.cml";

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    ConnectionHelper conn = ConnectionHelper.GetConnection(PrimaryKey);

                    XElement udsdef = XElement.Load(dialog.FileName);

                    if (dialog.FileName.EndsWith(".csml", StringComparison.InvariantCultureIgnoreCase))
                        conn.CallService("UDSManagerService.ImportContracts", new Envelope(new XStringHolder(udsdef.ToString())));
                    else if (dialog.FileName.EndsWith(".cml", StringComparison.InvariantCultureIgnoreCase))
                        conn.CallService("UDSManagerService.ImportContract", new Envelope(new XStringHolder(udsdef.ToString())));
                    else
                        throw new ArgumentException("不支援的格式！");

                    MessageBox.Show("安裝/更新 UDS 完成。");
                    OnPrimaryKeyChanged(EventArgs.Empty);
                }
                catch (Exception ex)
                {
                    RTOut.WriteError(ex);
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dr = MessageBox.Show("確定要刪除選擇的 UDS？", "Campus", MessageBoxButtons.YesNo);
                if (dr == DialogResult.No)
                    return;

                XElement req = new XElement("Request");
                foreach (DataGridViewRow row in dgvUDS.SelectedRows)
                {
                    string name = row.Cells["chName"].Value + "";
                    req.Add(new XElement("ContractName", name));
                }

                if (req.Elements().Count() > 0)
                {
                    ConnectionHelper conn = ConnectionHelper.GetConnection(PrimaryKey);
                    conn.CallService("UDSManagerService.DeleteContracts", new Envelope(new XHelper(req)));
                    OnPrimaryKeyChanged(EventArgs.Empty);

                    MessageBox.Show("UDS Contract 已刪除！");
                }
            }
            catch (Exception ex)
            {
                RTOut.WriteError(ex);
                MessageBox.Show(ex.Message);
            }
        }

        private void dgvUDS_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            DataGridViewRow row = dgvUDS.Rows[e.RowIndex];

            string name = row.Cells["chName"].Value + "";

            XElement contract = UDS.XPathSelectElement(string.Format("Contract[@Name='{0}']", name));

            SimpleDefContent udscontent = new SimpleDefContent();
            udscontent.Text = name;
            udscontent.SetXmlContent(contract.ToString());
            udscontent.ShowDialog();
        }

        private void UDSItem_Load(object sender, EventArgs e)
        {
            InitDetailContent();
        }
    }
}
