using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Xml.XPath;
using DesktopLib;
using DevComponents.DotNetBar;
using FISCA;
using FISCA.DSA;

namespace KHJHCentralOffice
{
    public partial class WebGadgetItem : DetailContentImproved
    {
        public static string Web2Url = "https://web2.ischool.com.tw/deployment";

        private XElement PackagesData { get; set; }

        private GadgetPackageRecord SelectedPackage { get; set; }

        public WebGadgetItem()
        {
            InitializeComponent();
            dgvGadgets.AutoGenerateColumns = false;
            Group = "Web2 Gadget";
            PackagesData = new XElement("Response");
            SelectedPackage = null;
        }

        protected override void OnPrimaryKeyChangedAsync()
        {
            ConnectionHelper conn = ConnectionHelper.GetConnection(PrimaryKey);
            Envelope rsp = conn.CallService("SelectWebPackage", new Envelope());
            XElement rspxml = XElement.Parse(rsp.BodyContent.XmlString);
            PackagesData = rspxml;
        }

        protected override void OnPrimaryKeyChangedComplete(Exception error)
        {
            if (error != null)
            {
                RTOut.WriteError(error);
                throw error;
            }

            XElement target = PackagesData.XPathSelectElement("Package[TargetType='Guest' and RefTagID='0']");
            tabGuest.Tag = CreateGadgetPackageRecord(target);

            target = PackagesData.XPathSelectElement("Package[TargetType='Teacher'  and RefTagID='0']");
            tabTeacher.Tag = CreateGadgetPackageRecord(target);

            target = PackagesData.XPathSelectElement("Package[TargetType='Student' and RefTagID='0']");
            tabStudent.Tag = CreateGadgetPackageRecord(target);

            target = PackagesData.XPathSelectElement("Package[TargetType='Parent' and RefTagID='0']");
            tabParent.Tag = CreateGadgetPackageRecord(target);

            target = PackagesData.XPathSelectElement("Package[TargetType='Administrator' and RefTagID='0']");
            tabAdmin.Tag = CreateGadgetPackageRecord(target);

            foreach (SuperTabItem item in tabs.Tabs)
                item.Visible = !(item.Tag == null);
            tabs.RecalcLayout();

            PuplateTabData(tabs.SelectedTab);
        }

        private static GadgetPackageRecord CreateGadgetPackageRecord(XElement target)
        {
            if (target == null)
                return null;
            else
                return new GadgetPackageRecord(target);
        }

        private void tabs_SelectedTabChanged(object sender, SuperTabStripSelectedTabChangedEventArgs e)
        {
            PuplateTabData(e.NewValue as SuperTabItem);
        }

        private void btnAddGadget_Click(object sender, EventArgs e)
        {
            InputBox box = new InputBox();
            DialogResult dr = box.ShowDialog();

            if (dr == DialogResult.OK)
            {
                try
                {
                    SelectedPackage.AddGadget(box.InputString);
                    SaveGadget(SelectedPackage);
                    OnPrimaryKeyChanged(EventArgs.Empty);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void btnDeleteGadget_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dr = MessageBox.Show("確定要刪除選擇的小工具？", "Campus", MessageBoxButtons.YesNo);

                if (dr == DialogResult.No)
                    return;

                foreach (DataGridViewRow row in dgvGadgets.SelectedRows)
                {
                    GadgetGridRow gg = row.DataBoundItem as GadgetGridRow;
                    SelectedPackage.RemoveGadget(gg.DeployPath);
                }

                SaveGadget(SelectedPackage);
                OnPrimaryKeyChanged(EventArgs.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnAdjustOrder_Click(object sender, EventArgs e)
        {
            try
            {
                List<GadgetGridRow> rows = new List<GadgetGridRow>(dgvGadgets.DataSource as BindingList<GadgetGridRow>);

                GadgetOrderAdjustForm adjust = new GadgetOrderAdjustForm(SelectedPackage, rows);
                DialogResult dr = adjust.ShowDialog();

                if (dr == DialogResult.OK)
                {
                    SaveGadget(SelectedPackage);
                    OnPrimaryKeyChanged(EventArgs.Empty);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void PuplateTabData(SuperTabItem superTabItem)
        {
            dgvGadgets.DataSource = null;
            SelectedPackage = superTabItem.Tag as GadgetPackageRecord;

            if (SelectedPackage != null)
            {
                List<GadgetGridRow> rows = new List<GadgetGridRow>();
                foreach (XElement gadget in SelectedPackage.EachGadget())
                    rows.Add(new GadgetGridRow(gadget.AttributeText("deployPath")));

                dgvGadgets.DataSource = new BindingList<GadgetGridRow>(rows);
            }
        }

        private static XElement GetGadgetDescription(string deployPath)
        {
            string url = string.Format(Web2Url + "/{0}/description.xml", deployPath);

            return XElement.Load(url);
        }

        private void SaveGadget(GadgetPackageRecord record)
        {
            try
            {
                ConnectionHelper ch = ConnectionHelper.GetConnection(PrimaryKey);

                XElement req = new XElement("Request",
                    new XElement("Package",
                        new XElement("Field", record.Definition),
                        new XElement("Condition",
                            new XElement("Uid", record.UID))));

                ch.CallService("UpdateWebPackage", new Envelope(new XStringHolder(req)));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgvGadgets_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            DataGridViewRow row = dgvGadgets.Rows[e.RowIndex];
            GadgetGridRow grow = row.DataBoundItem as GadgetGridRow;
            MessageBox.Show(grow.DeployPath);
        }

        #region GadgetGridRow Class
        /// <summary>
        /// 用於 DataGridView 的 Value Object。
        /// </summary>
        internal class GadgetGridRow : INotifyPropertyChanged
        {
            public GadgetGridRow(string deployPath)
            {
                Title = "讀取中...";
                DeployPath = deployPath;
                Description = DeployPath;

                Task task = Task.Factory.StartNew(() =>
                {
                    try
                    {
                        DeployDescription = new GadgetDeployDescription(DeployPath);

                        Title = DeployDescription.Title;
                        Description = DeployDescription.Description;
                    }
                    catch (Exception ex)
                    {
                        Title = "讀取錯誤...";
                        Description = ex.Message;
                    }
                }, System.Threading.CancellationToken.None, TaskCreationOptions.None, TaskScheduler.Default);
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

            private string _description = string.Empty;
            public string Description
            {
                get { return _description; }
                set
                {
                    _description = value;
                    if (PropertyChanged != null)
                        PropertyChanged(this, new PropertyChangedEventArgs("Description"));
                }
            }

            public string DeployPath { get; private set; }

            public GadgetDeployDescription DeployDescription { get; set; }

            #region INotifyPropertyChanged 成員

            public event PropertyChangedEventHandler PropertyChanged;

            #endregion
        }
        #endregion

        #region GadgetDeployDescription
        /// <summary>
        /// 代表 Gadget 在 web2 上的 description.xml 內容。
        /// </summary>
        internal class GadgetDeployDescription
        {
            public GadgetDeployDescription(string deployPath)
            {
                XElement descxml = GetGadgetDescription(deployPath);
                XElement prefs = descxml.Element("ModulePrefs");

                DeployPath = deployPath;
                Title = prefs.AttributeText("title");
                Description = prefs.AttributeText("description");
                IconPath = prefs.ElementText("Icon");
            }

            public string DeployPath { get; private set; }

            public string Title { get; private set; }

            public string Description { get; private set; }

            public string IconPath { get; private set; }
        }
        #endregion

        #region GadgetPackageRecord
        /// <summary>
        /// 資料庫中的 Package 資料。
        /// </summary>
        internal class GadgetPackageRecord
        {
            public GadgetPackageRecord(XElement data)
            {
                UID = data.ElementText("Uid");
                Definition = data.Element("Definition");
            }

            public string UID { get; set; }

            public XElement Definition { get; set; }

            public IEnumerable<XElement> EachGadget()
            {
                foreach (XElement each in Definition.XPathSelectElements("Package/Gadgets/Gadget"))
                    yield return each;
            }

            public bool GadgetExists(string deployPath)
            {
                foreach (XElement each in EachGadget())
                {
                    string dp = each.AttributeText("deployPath");
                    if (dp == deployPath)
                        return true;
                }
                return false;
            }

            public XElement GetGadgetXml(string deployPath)
            {
                string xpath = string.Format("Package/Gadgets/Gadget[@deployPath='{0}']", deployPath);

                return Definition.XPathSelectElement(xpath);
            }

            public void AddGadget(string deployPath)
            {
                GadgetDeployDescription gdd = new GadgetDeployDescription(deployPath);

                string desc = string.Format("{0}({1})", gdd.Title, gdd.Description);

                XElement gadget = new XElement("Gadget");
                gadget.SetAttributeValue("deployPath", deployPath);
                gadget.SetAttributeValue("Description", desc);

                Definition.XPathSelectElement("Package/Gadgets").Add(gadget);
            }

            public void RemoveGadget(string deployPath)
            {
                string xpath = string.Format("Package/Gadgets/Gadget[@deployPath='{0}']", deployPath);
                XElement gadget = Definition.XPathSelectElement(xpath);

                if (gadget != null)
                    gadget.Remove();
            }

            public void ReplaceGadgets(XElement[] gadgets)
            {
                XElement old = Definition.XPathSelectElement("Package/Gadgets");
                old.ReplaceAll(gadgets);
            }
        }
        #endregion

        private void WebGadgetItem_Load(object sender, EventArgs e)
        {
            InitDetailContent();
        }
    }
}