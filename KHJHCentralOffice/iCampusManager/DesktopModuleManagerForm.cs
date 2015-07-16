using DesktopLib;
using FISCA;
using FISCA.DSA;
using FISCA.Presentation;
using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;

namespace KHJHCentralOffice
{
    public partial class DesktopModuleManagerForm : BaseForm
    {
        private List<ConnectionHelper> Connections;

        private Dictionary<string, ModulesOfSchool> ModuleConfigs = new Dictionary<string, ModulesOfSchool>();

        public DesktopModuleManagerForm()
        {
            InitializeComponent();
        }

        private void DesktopModuleManagerForm_Load(object sender, EventArgs e)
        {
            try
            {
                InitConnections();
                LoadAllModuleConfig();
                GroupByToDataGrid();
            }
            catch { Close(); } //爆了就關吧。
        }

        private void GroupByToDataGrid()
        {
            List<ModuleRow> modules = GroupByModules();
            modules.Sort((x, y) =>
            {
                return x.Url.CompareTo(y.Url);
            });
            dgvModules.DataSource = new SortableBindingList<ModuleRow>(modules);
        }

        private List<ModuleRow> GroupByModules()
        {
            Dictionary<string, ModuleRow> result = new Dictionary<string, ModuleRow>();

            foreach (ModulesOfSchool each in ModuleConfigs.Values)
            {
                foreach (string url in each.Urls)
                {
                    if (!result.ContainsKey(url))
                        result.Add(url, new ModuleRow(url));

                    result[url].IncreaseRefCount();
                }
            }

            return result.Values.ToList();
        }

        private void LoadAllModuleConfig()
        {
            MultiTaskingRunner runner = new MultiTaskingRunner();

            ModuleConfigs = new Dictionary<string, ModulesOfSchool>();
            foreach (ConnectionHelper conn in Connections)
            {
                string name = Program.GlobalSchoolCache[conn.UID].Title;
                runner.AddTask(string.Format("{0}({1})", name, conn.UID), (x) =>
                {
                    LoadModuleConfig(x as ConnectionHelper);
                }, conn, new System.Threading.CancellationTokenSource());
            }
            runner.ExecuteTasks();
        }

        private void LoadModuleConfig(ConnectionHelper conn)
        {
            Envelope rsp = conn.CallService("SelectDesktopModule", new Envelope());
            XElement conf = XElement.Parse(rsp.BodyContent.XmlString);

            lock (ModuleConfigs)
            {
                ModuleConfigs[conn.UID] = new ModulesOfSchool(conf, conn.UID);
            }
        }

        private void InitConnections()
        {
            List<string> selected = Program.MainPanel.SelectedSource;
            List<ConnectionHelper> conns = new List<ConnectionHelper>();

            foreach (string uid in selected)
            {
                conns.Add(ConnectionHelper.GetConnection(uid));
            }

            Connections = conns;
        }

        class ModuleRow
        {
            public ModuleRow(string url)
            {
                Url = url;
            }

            public string Url { get; private set; }

            public int ReferenceCount { get; private set; }

            public void IncreaseRefCount()
            {
                ReferenceCount++;
            }

            public void DecreaseRefCount()
            {
                ReferenceCount--;
            }
        }

        class ModulesOfSchool
        {
            private XElement Modules;

            public ModulesOfSchool(XElement modData, string uid)
            {
                Modules = modData;
                UID = uid;
            }

            public string UID { get; private set; }

            public HashSet<string> Urls
            {
                get
                {
                    HashSet<string> hs = new HashSet<string>();

                    foreach (XElement each in Modules.Elements("Module"))
                    {
                        string murl = each.Element("ModuleUrl").Value;
                        hs.Add(murl);
                    }

                    return hs;
                }
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            MultiTaskingRunner runner = new MultiTaskingRunner();

            InputBox box = new InputBox();
            box.ShowDialog();

            string modUrl = box.InputString;
            foreach (ConnectionHelper conn in Connections)
            {
                string name = Program.GlobalSchoolCache[conn.UID].Title;
                runner.AddTask(string.Format("{0}({1})", name, conn.UID), (x) =>
                {
                    object[] obj = x as object[];
                    ConnectionHelper c = obj[0] as ConnectionHelper;
                    string xurl = (string)obj[1];
                    AddDesktopModule(c, xurl);
                }, new object[] { conn, modUrl }, new System.Threading.CancellationTokenSource());
            }
            runner.ExecuteTasks();

            try
            {
                InitConnections();
                LoadAllModuleConfig();
                GroupByToDataGrid();
            }
            catch { Close(); } //爆了就關吧。
        }

        private void AddDesktopModule(ConnectionHelper conn, string url)
        {
            XElement req = new XElement("Request",
                new XElement("Module",
                    new XElement("Field",
                        new XElement("ModuleUrl", url),
                        new XElement("Type", "FiscaAEModule"),
                        new XElement("Memo", "集中安裝"),
                        new XElement("Config",
                            new XElement("Content",
                                new XElement("VersionOption", "Stable"))))));

            conn.CallService("InsertDesktopModule", new Envelope(new FISCA.XStringHolder(req)));
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (dgvModules.SelectedRows.Count <= 0)
                return;

            DialogResult dr = MessageBox.Show("確定要刪除選擇的模組？", "ischool", MessageBoxButtons.YesNo);
            if (dr == System.Windows.Forms.DialogResult.No)
                return;

            MultiTaskingRunner runner = new MultiTaskingRunner();
            ModuleRow mr = dgvModules.SelectedRows[0].DataBoundItem as ModuleRow;
            string modUrl = mr.Url;

            foreach (ConnectionHelper conn in Connections)
            {
                string name = Program.GlobalSchoolCache[conn.UID].Title;
                runner.AddTask(string.Format("{0}({1})", name, conn.UID), (x) =>
                {
                    object[] obj = x as object[];
                    ConnectionHelper c = obj[0] as ConnectionHelper;
                    string xurl = (string)obj[1];
                    RemoveDesktopModule(c, xurl);
                }, new object[] { conn, modUrl }, new System.Threading.CancellationTokenSource());
            }
            runner.ExecuteTasks();

            try
            {
                InitConnections();
                LoadAllModuleConfig();
                GroupByToDataGrid();
            }
            catch { Close(); } //爆了就關吧。
        }

        private void RemoveDesktopModule(ConnectionHelper conn, string xurl)
        {
            XElement req = new XElement("Request",
                new XElement("Module",
                  new XElement("Condition",
                      new XElement("ModuleUrl", xurl))));

            conn.CallService("DeleteDesktopModule", new Envelope(new XHelper(req)));
        }

    }
}
