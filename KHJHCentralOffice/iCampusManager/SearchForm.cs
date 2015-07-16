using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using FISCA.DSA;
using FISCA.Presentation.Controls;
using FISCA;
using DesktopLib;

namespace KHJHCentralOffice
{
    public partial class SearchForm : BaseForm
    {
        public SearchForm()
        {
            InitializeComponent();
        }

        private void btnPlugin_Click(object sender, EventArgs e)
        {
            Search(MatchDesktop);
        }

        private void btnGadget_Click(object sender, EventArgs e)
        {
            Search(MatchGadget);
        }

        private void btnUDM_Click(object sender, EventArgs e)
        {
            Search(MatchUDM);
        }

        private void Search(Action<List<string>, string, object> patternMatcher)
        {
            MultiTaskingRunner runner = new MultiTaskingRunner();
            List<string> matchs = new List<string>();
            string pattern = txtPattern.Text;

            foreach (string uid in Program.MainPanel.SelectedSource)
            {
                string name = Program.GlobalSchoolCache[uid].Title;
                ConnectionHelper ch = ConnectionHelper.GetConnection(uid);
                runner.AddTask(string.Format("{0}({1})", name, uid), (x) =>
                {
                    patternMatcher(matchs, pattern, x);
                }, new object[] { uid, ch }, new CancellationTokenSource());
            }

            runner.ExecuteTasks();
            Program.MainPanel.AddToTemp(matchs);
            DialogResult = DialogResult.OK;
        }

        private static void MatchDesktop(List<string> matchs, string pattern, object x)
        {
            object[] args = x as object[];
            string uidx = (string)args[0];
            ConnectionHelper chx = args[1] as ConnectionHelper;

            Envelope rsp = chx.CallService("SelectDesktopModule", new Envelope());
            if (rsp.BodyContent.XmlString.IndexOf(pattern) >= 0)
            {
                lock (matchs)
                {
                    matchs.Add(uidx);
                }
            }
        }

        private static void MatchGadget(List<string> matchs, string pattern, object x)
        {
            object[] args = x as object[];
            string uidx = (string)args[0];
            ConnectionHelper chx = args[1] as ConnectionHelper;

            Envelope rsp = chx.CallService("SelectWebPackage", new Envelope());
            if (rsp.BodyContent.XmlString.IndexOf(pattern) >= 0)
            {
                lock (matchs)
                {
                    matchs.Add(uidx);
                }
            }
        }

        private static void MatchUDM(List<string> matchs, string pattern, object x)
        {
            object[] args = x as object[];
            string uidx = (string)args[0];
            ConnectionHelper chx = args[1] as ConnectionHelper;

            Envelope rsp = chx.CallService("UDMService.ListModules", new Envelope());
            if (rsp.BodyContent.XmlString.IndexOf(pattern) >= 0)
            {
                lock (matchs)
                {
                    matchs.Add(uidx);
                }
            }
        }
    }
}
