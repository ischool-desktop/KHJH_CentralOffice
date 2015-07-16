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
    public partial class UDTItem : DetailContentImproved
    {
        private XElement UDT { get; set; }

        public UDTItem()
        {
            InitializeComponent();
            Group = "UDT";
        }

        protected override void OnPrimaryKeyChangedAsync()
        {
            ConnectionHelper conn = ConnectionHelper.GetConnection(PrimaryKey);

            Envelope rsp = conn.CallService("UDTService.DDL.ListTableNames", new Envelope());

            UDT = XElement.Parse(rsp.BodyContent.XmlString);
        }

        protected override void OnPrimaryKeyChangedComplete(Exception error)
        {
            if (error != null)
            {
                RTOut.WriteError(error);
                throw error;
            }

            dgvUDT.Rows.Clear();

            List<string> udtNames = new List<string>();

            foreach (XElement each in UDT.Elements("TableName"))
                udtNames.Add(each.Value);

            udtNames.Sort();

            foreach (string each in udtNames)
            {
                DataGridViewRow row = new DataGridViewRow();

                row.CreateCells(dgvUDT, each);

                dgvUDT.Rows.Add(row);
            }
        }

        private void btnInstall_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Filter = "UDT 定議 (*.tsml;*.tml)|*.tsml;*.tml";

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    ConnectionHelper conn = ConnectionHelper.GetConnection(PrimaryKey);

                    XElement udsdef = XElement.Load(dialog.FileName);

                    if (dialog.FileName.EndsWith(".tsml", StringComparison.InvariantCultureIgnoreCase))
                        conn.CallService("UDTService.DDL.SetTables", new Envelope(new XStringHolder(udsdef.ToString())));
                    else if (dialog.FileName.EndsWith(".tml", StringComparison.InvariantCultureIgnoreCase))
                        conn.CallService("UDTService.DDL.SetTable", new Envelope(new XStringHolder(udsdef.ToString())));
                    else
                        throw new ArgumentException("不支援的格式！");

                    MessageBox.Show("安裝/更新 UDT 完成。");
                    OnPrimaryKeyChanged(EventArgs.Empty);
                }
                catch (Exception ex)
                {
                    RTOut.WriteError(ex);
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void dgvUDT_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            DataGridViewRow row = dgvUDT.Rows[e.RowIndex];

            string name = row.Cells["chName"].Value + "";

            UDTTableDefContent udtcontent = new UDTTableDefContent();
            udtcontent.Text = name;
            udtcontent.SetTableName(name);
            udtcontent.SetUID(PrimaryKey);
            udtcontent.ShowDialog();
        }

        private void UDTItem_Load(object sender, EventArgs e)
        {
            InitDetailContent();
        }
    }
}
