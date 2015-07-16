﻿using FISCA.DSA;
using FISCA.Presentation.Controls;
using KHJHCentralOffice;
using System;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Xml.XPath;

namespace VirtualAccountFound
{
    public partial class VAFinder : BaseForm
    {
        const int BalanceAccountLength = 15;

        public VAFinder()
        {
            InitializeComponent();
        }

        private void btnQuery_Click(object sender, EventArgs e)
        {
            try
            {
                string cmd = @"
                            select amount,paidamount,paydate,billxmldata from $paymenthistory.accountsreceivable 
                            where uid in 
	                            (
	                            select cast(refpaymenthistoryuid as bigint)
	                            from $paymenttransaction.accountsreceivable 
	                            where balanceaccount = '@BalanceAccount'
	                            )
                            ";

                string account = GetAccount();
                cmd = cmd.Replace("@BalanceAccount", account);

                ConnectionHelper conn = ConnectionHelper.GetConnection(Program.MainPanel.SelectedSource.First());

                XElement req = new XElement("Request",
                    new XElement("SQL", cmd));

                Envelope rsp = conn.CallService("UDTService.DML.Query", new Envelope(new FISCA.XStringHolder(req.ToString())));

                dgvMerge.Rows.Clear();
                dgvPayItem.Rows.Clear();

                lblPayInfo.Text = "繳費資訊";

                XElement record = XElement.Parse(rsp.BodyContent.XmlString);
                record = record.Element("Record");
                if (record != null)
                {
                    string amount = record.XPathSelectElement("Column[@Index='1']").Value;   //amount
                    string paidamount = record.XPathSelectElement("Column[@Index='2']").Value;   //paidamount
                    string paydate = record.XPathSelectElement("Column[@Index='3']").Value;  //paydate
                    string billxml = record.XPathSelectElement("Column[@Index='4']").Value;  //billxmldata

                    lblPayInfo.Text = string.Format("繳費資訊, 應繳：{0}, 已繳：{1}, 繳費日期：{2}", amount, paidamount, paydate);

                    XElement objxml = XElement.Parse(billxml);

                    foreach (XElement each in objxml.Element("Extensions").Elements("Extension"))
                    {
                        string[] fieldinfos = each.Attribute("Name").Value.Split(new string[] { "::" },
                            StringSplitOptions.RemoveEmptyEntries);

                        string type = fieldinfos[0].ToLower();
                        string name = string.Empty;
                        string value = each.Value;

                        if (fieldinfos.Length <= 1) //只有一個元素時都算是 MergeField。
                        {
                            type = "MergeField".ToLower();
                            name = fieldinfos[0];
                        }
                        else
                            name = fieldinfos[1];

                        if (type == "MergeField".ToLower())
                        {
                            DataGridViewRow row = new DataGridViewRow();
                            row.CreateCells(dgvMerge, name, value);
                            dgvMerge.Rows.Add(row);
                        }
                        else if (type == "PayItem".ToLower())
                        {
                            DataGridViewRow row = new DataGridViewRow();
                            row.CreateCells(dgvMerge, name, value);
                            dgvPayItem.Rows.Add(row);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("找不到資料。");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private string GetAccount()
        {
            string account = txtVirualAccount.Text;

            account = account.Substring(0, BalanceAccountLength);

            return account;
        }
    }
}
