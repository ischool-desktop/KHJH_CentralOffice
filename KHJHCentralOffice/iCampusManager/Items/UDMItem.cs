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
using DesktopLib;

namespace KHJHCentralOffice
{
    public partial class UDMItem : DetailContentImproved
    {
        private XElement UDM { get; set; }

        public UDMItem()
        {
            InitializeComponent();
            Group = "UDM";
            dgvUDM.AutoGenerateColumns = false;
        }

        protected override void OnPrimaryKeyChangedAsync()
        {
            ConnectionHelper conn = ConnectionHelper.GetConnection(PrimaryKey);

            Envelope rsp = conn.CallService("UDMService.ListModules", new Envelope());

            UDM = XElement.Parse(rsp.BodyContent.XmlString);
        }

        protected override void OnPrimaryKeyChangedComplete(Exception error)
        {
            if (error != null)
            {
                RTOut.WriteError(error);
                throw error;
            }

            dgvUDM.Rows.Clear();

            List<UDMGridRow> udms = new List<UDMGridRow>();
            foreach (XElement each in UDM.Elements("Module"))
                udms.Add(new UDMGridRow(each));

            udms.Sort((x, y) =>
            {
                return x.Name.CompareTo(y.Name);
            });

            dgvUDM.DataSource = new BindingList<UDMGridRow>(udms);
        }

        private void btnInstall_Click(object sender, EventArgs e)
        {
            try
            {
                InputBox box = new InputBox();
                if (box.ShowDialog() != DialogResult.OK)
                    return;

                string url = box.InputString;
                XElement req = new XElement("Request", new XElement("URL", url));

                ConnectionHelper conn = ConnectionHelper.GetConnection(PrimaryKey);
                conn.CallService("UDMService.ForceRegistry", new Envelope(new XStringHolder(req)));

                MessageBox.Show("註冊 UDM 完成。");
                OnPrimaryKeyChanged(EventArgs.Empty);
            }
            catch (Exception ex)
            {
                RTOut.WriteError(ex);
                MessageBox.Show(ex.Message);
            }
        }

        private void btnCheckUpdate_Click(object sender, EventArgs e)
        {
            IEnumerable<UDMGridRow> udms = GetSelectedUDM();

            MultiTaskingRunner runner = new MultiTaskingRunner();

            foreach (UDMGridRow udm in udms)
            {
                runner.AddTask(udm.Name, x =>
                {
                    UDMGridRow s = x as UDMGridRow;
                    CheckUDMNewVersion(s);
                }, udm, new System.Threading.CancellationTokenSource());
            }

            runner.AllTaskCompleted += delegate
            {
            };

            runner.ExecuteTasks();
        }

        private void CheckUDMNewVersion(UDMGridRow s)
        {
            ConnectionHelper conn = ConnectionHelper.GetConnection(PrimaryKey);

            XElement req = new XElement("Request", new XElement("ModuleName", s.Name));

            Envelope rsp = conn.CallService("UDMService.CheckUpdate", new Envelope(new XStringHolder(req)));
            XElement rspxml = XElement.Parse(rsp.BodyContent.XmlString);
            bool hasUpdate = rspxml.ElementBool("HasUpdate", false);

            if (hasUpdate)
                s.Version = s.RawData.ElementText("Version") + "*";
        }

        private IEnumerable<UDMGridRow> GetSelectedUDM()
        {
            foreach (DataGridViewRow row in dgvUDM.Rows)
                yield return row.DataBoundItem as UDMGridRow;
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgvUDM.SelectedRows.Count <= 0)
                    return;

                DialogResult dr = MessageBox.Show("如果 UDM 包含 UDT 會一同刪除！您確定？", "Campus", MessageBoxButtons.YesNo);
                if (dr == DialogResult.No)
                    return;

                DataGridViewRow row = dgvUDM.SelectedRows[0];
                string name = (row.DataBoundItem as UDMGridRow).Name;
                XElement req = new XElement("Request", new XElement("ModuleName", name));

                ConnectionHelper conn = ConnectionHelper.GetConnection(PrimaryKey);
                conn.CallService("UDMService.RemoveModule", new Envelope(new XStringHolder(req)));

                MessageBox.Show("刪除 UDM 完成。");
                OnPrimaryKeyChanged(EventArgs.Empty);
            }
            catch (Exception ex)
            {
                RTOut.WriteError(ex);
                MessageBox.Show(ex.Message);
            }
        }

        private void dgvUDM_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            DataGridViewRow row = dgvUDM.Rows[e.RowIndex];

            string xml = (row.DataBoundItem as UDMGridRow).RawData.ToString();

            SimpleDefContent udscontent = new SimpleDefContent();
            udscontent.Text = row.Cells["chName"].Value + "";
            udscontent.SetXmlContent(xml);
            udscontent.ShowDialog();

        }

        class UDMGridRow : INotifyPropertyChanged
        {
            public UDMGridRow(XElement data)
            {
                RawData = data;

                Name = data.ElementText("Name");
                Provider = data.ElementText("Provider");
                Version = data.ElementText("Version");

                Url = Program.TrimModuleServerUrl(data.ElementText("URL"));
            }

            public string Name { get; private set; }

            public string Provider { get; private set; }

            private string _version = string.Empty;
            public string Version
            {
                get { return _version; }
                set
                {
                    _version = value;
                    if (PropertyChanged != null)
                        PropertyChanged(this, new PropertyChangedEventArgs("Version"));
                }
            }

            public string Url { get; private set; }

            public XElement RawData { get; set; }

            #region INotifyPropertyChanged 成員

            public event PropertyChangedEventHandler PropertyChanged;

            #endregion
        }

        private void buttonItem1_Click(object sender, EventArgs e)
        {
            if (dgvUDM.SelectedRows.Count < 0)
                return;

            DataGridViewRow row = dgvUDM.SelectedRows[0];
            UDMGridRow udmrow = (row.DataBoundItem as UDMGridRow);

            XElement req = new XElement("Request", new XElement("ModuleName", udmrow.Name));

            try
            {
                ConnectionHelper conn = ConnectionHelper.GetConnection(PrimaryKey);
                conn.CallService("UDMService.UpdateModule", new Envelope(new XStringHolder(req)));

                MessageBox.Show("更新完成。");
                OnPrimaryKeyChanged(EventArgs.Empty);
            }
            catch (Exception ex)
            {
                RTOut.WriteError(ex);
                MessageBox.Show(ex.Message);
            }
        }

        private void UDMItem_Load(object sender, EventArgs e)
        {
            InitDetailContent();
        }
    }
}
