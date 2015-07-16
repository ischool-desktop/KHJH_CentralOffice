using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using FISCA.Presentation.Controls;

namespace KHJHCentralOffice
{
    internal partial class GadgetOrderAdjustForm : BaseForm
    {
        public WebGadgetItem.GadgetPackageRecord Record { get; set; }

        private BindingList<WebGadgetItem.GadgetGridRow> Rows { get; set; }

        public GadgetOrderAdjustForm(
            WebGadgetItem.GadgetPackageRecord record,
            List<WebGadgetItem.GadgetGridRow> rows)
        {
            InitializeComponent();
            Record = record;
            Rows = new BindingList<WebGadgetItem.GadgetGridRow>(rows);

            dgvGadgets.AutoGenerateColumns = false;
            dgvGadgets.DataSource = rows;
        }

        private void PlaceGadgetXmlInRow()
        {
            foreach (DataGridViewRow row in dgvGadgets.Rows)
            {
                WebGadgetItem.GadgetGridRow gg = row.DataBoundItem as WebGadgetItem.GadgetGridRow;
                XElement gxml = Record.GetGadgetXml(gg.DeployPath);
                row.Tag = gxml;
            }
        }

        private void btnUp_Click(object sender, EventArgs e)
        {
            if (dgvGadgets.SelectedRows.Count <= 0)
                return;

            Dictionary<WebGadgetItem.GadgetGridRow, int> ggs = new Dictionary<WebGadgetItem.GadgetGridRow, int>();
            List<WebGadgetItem.GadgetGridRow> selecteds = new List<WebGadgetItem.GadgetGridRow>();
            int fixedIndex = -1;
            foreach (DataGridViewRow grow in dgvGadgets.Rows)
            {
                if (!grow.Selected) //需要用這種方式取得 Selected，不然順序會亂掉。
                    continue;

                selecteds.Add(grow.DataBoundItem as WebGadgetItem.GadgetGridRow);

                int newIndex = grow.Index - 1;

                if (newIndex == fixedIndex)
                {
                    fixedIndex++;
                    continue;
                }

                WebGadgetItem.GadgetGridRow gg = grow.DataBoundItem as WebGadgetItem.GadgetGridRow;
                ggs.Add(gg, newIndex);
            }

            foreach (WebGadgetItem.GadgetGridRow gg in ggs.Keys)
            {
                if (ggs[gg] < 0) continue;

                Rows.Remove(gg);
                Rows.Insert(ggs[gg], gg);
            }

            foreach (DataGridViewRow grow in dgvGadgets.Rows)
                grow.Selected = selecteds.Contains(grow.DataBoundItem as WebGadgetItem.GadgetGridRow);

            dgvGadgets.Refresh();
        }

        private void btnDown_Click(object sender, EventArgs e)
        {
            if (dgvGadgets.SelectedRows.Count <= 0)
                return;

            Dictionary<WebGadgetItem.GadgetGridRow, int> ggs = new Dictionary<WebGadgetItem.GadgetGridRow, int>();
            List<WebGadgetItem.GadgetGridRow> selecteds = new List<WebGadgetItem.GadgetGridRow>();

            List<DataGridViewRow> dgvRows = new List<DataGridViewRow>();
            foreach (DataGridViewRow each in dgvGadgets.Rows)
                dgvRows.Add(each);
            dgvRows.Reverse();

            int fixedIndex = dgvGadgets.Rows.Count;
            foreach (DataGridViewRow grow in dgvRows)
            {
                if (!grow.Selected) //需要用這種方式取得 Selected，不然順序會亂掉。
                    continue;

                selecteds.Add(grow.DataBoundItem as WebGadgetItem.GadgetGridRow);

                int newIndex = grow.Index + 1;

                if (newIndex == fixedIndex)
                {
                    fixedIndex--;
                    continue;
                }

                WebGadgetItem.GadgetGridRow gg = grow.DataBoundItem as WebGadgetItem.GadgetGridRow;
                ggs.Add(gg, newIndex);
            }

            foreach (WebGadgetItem.GadgetGridRow gg in ggs.Keys)
            {
                if (ggs[gg] >= dgvGadgets.Rows.Count) continue;

                Rows.Remove(gg);
                Rows.Insert(ggs[gg], gg);
            }

            foreach (DataGridViewRow grow in dgvGadgets.Rows)
                grow.Selected = selecteds.Contains(grow.DataBoundItem as WebGadgetItem.GadgetGridRow);

            dgvGadgets.Refresh();
        }

        private void btnConfirm_Click(object sender, EventArgs e)
        {
            List<XElement> gadgets = new List<XElement>();

            PlaceGadgetXmlInRow();
            foreach (DataGridViewRow row in dgvGadgets.Rows)
            {
                XElement gadget = row.Tag as XElement;
                if (gadget != null)
                    gadgets.Add(gadget);
            }
            Record.ReplaceGadgets(gadgets.ToArray());
        }
    }
}
