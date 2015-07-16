using System;
using System.Collections.Generic;
using System.ComponentModel;
using FISCA.DSA;
using System.Xml.Linq;
using FISCA;
using System.Threading.Tasks;
using FISCA.Deployment;
using System.Threading;
using System.Windows.Forms;
using DesktopLib;

namespace KHJHCentralOffice
{
    public partial class ModuleItem : DetailContentImproved
    {
        private XElement Plugin { get; set; }

        public ModuleItem()
        {
            InitializeComponent();
            Group = "Desktop Plugin";
            dgvPlugin.AutoGenerateColumns = false;
        }

        protected override void OnPrimaryKeyChangedAsync()
        {
            ConnectionHelper conn = ConnectionHelper.GetConnection(PrimaryKey);

            Envelope rsp = conn.CallService("SelectDesktopModule", new Envelope());

            Plugin = XElement.Parse(rsp.BodyContent.XmlString);
        }

        protected override void OnPrimaryKeyChangedComplete(Exception error)
        {
            if (error != null)
            {
                RTOut.WriteError(error);
                throw error;
            }

            dgvPlugin.Rows.Clear();

            List<string> pluginUrls = new List<string>();

            foreach (XElement each in Plugin.Elements("Module"))
                pluginUrls.Add(each.ElementText("ModuleUrl"));

            pluginUrls.Sort();

            List<PluginGridRow> rows = new List<PluginGridRow>();
            foreach (string each in pluginUrls)
                rows.Add(new PluginGridRow(each));

            dgvPlugin.DataSource = new BindingList<PluginGridRow>(rows);
        }

        private void btnInstall_Click(object sender, EventArgs e)
        {
            try
            {
                InputBox box = new InputBox();
                if (box.ShowDialog() != DialogResult.OK)
                    return;

                string url = box.InputString;
                XElement req = new XElement("Request",
                    new XElement("Module",
                        new XElement("Field",
                            new XElement("ModuleUrl", url),
                            new XElement("Type", "FiscaAEModule"),
                            new XElement("Memo", "集中安裝"),
                            new XElement("Config",
                                new XElement("Content",
                                    new XElement("VersionOption", "Stable"))))));

                ConnectionHelper conn = ConnectionHelper.GetConnection(PrimaryKey);
                conn.CallService("InsertDesktopModule", new Envelope(new XStringHolder(req)));

                MessageBox.Show("安裝完成。");
                OnPrimaryKeyChanged(EventArgs.Empty);
            }
            catch (Exception ex)
            {
                RTOut.WriteError(ex);
                MessageBox.Show(ex.Message);
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgvPlugin.SelectedRows.Count <= 0)
                    return;

                DialogResult dr = MessageBox.Show("確定要刪除選擇的項目？", "Campus", MessageBoxButtons.YesNo);
                if (dr == DialogResult.No)
                    return;

                PluginGridRow pg = dgvPlugin.SelectedRows[0].DataBoundItem as PluginGridRow;
                XElement req = new XElement("Request",
                    new XElement("Module",
                        new XElement("Condition",
                            new XElement("ModuleUrl", pg.Url))));

                ConnectionHelper conn = ConnectionHelper.GetConnection(PrimaryKey);
                conn.CallService("DeleteDesktopModule", new Envelope(new XHelper(req)));
                OnPrimaryKeyChanged(EventArgs.Empty);

                MessageBox.Show("項目已刪除！");
            }
            catch (Exception ex)
            {
                RTOut.WriteError(ex);
                MessageBox.Show(ex.Message);
            }
        }

        internal class PluginGridRow : INotifyPropertyChanged
        {
            public PluginGridRow(string url)
            {
                Title = "讀取中...";
                Url = url;
                UrlShort = Program.TrimModuleServerUrl(url);

                Task task = Task.Factory.StartNew(() =>
                {
                    Manifest = new AppManifest();
                    Manifest.Load(url);
                }, CancellationToken.None, TaskCreationOptions.None, TaskScheduler.Default);

                task.ContinueWith(x =>
                {
                    if (x.Exception != null)
                        Title = "讀取錯誤!";
                    else
                        Title = Manifest.Title;
                }, TaskScheduler.FromCurrentSynchronizationContext());
            }

            private string _title = string.Empty;
            public string Title
            {
                get { return _title; }
                set
                {
                    _title = value;

                    if (PropertyChanged != null)
                        PropertyChanged(this, new PropertyChangedEventArgs("Title"));
                }
            }

            public string UrlShort { get; set; }

            public string Url { get; private set; }

            public AppManifest Manifest { get; private set; }

            #region INotifyPropertyChanged 成員

            public event PropertyChangedEventHandler PropertyChanged;

            #endregion
        }

        private void ModuleItem_Load(object sender, EventArgs e)
        {
            InitDetailContent();
        }

        private void dgvPlugin_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            DataGridViewRow row = dgvPlugin.Rows[e.RowIndex];

            string xml = (row.DataBoundItem as PluginGridRow).Manifest.ToXml().ToString();

            SimpleDefContent udscontent = new SimpleDefContent();
            udscontent.Text = (row.DataBoundItem as PluginGridRow).Url;
            udscontent.SetXmlContent(xml);
            udscontent.ShowDialog();
        }
    }
}
