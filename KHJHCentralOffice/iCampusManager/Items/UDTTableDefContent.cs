using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using FISCA;
using FISCA.DSA;
using FISCA.Presentation.Controls;

namespace KHJHCentralOffice
{
    public partial class UDTTableDefContent : BaseForm
    {
        private string TableName { get; set; }

        private string UID { get; set; }

        public UDTTableDefContent()
        {
            InitializeComponent();
        }

        public void SetTableName(string name)
        {
            TableName = name;
        }

        internal void SetUID(string pKey)
        {
            UID = pKey;
        }

        private void UDTTableDefContent_Load(object sender, EventArgs e)
        {
            try
            {
                ConnectionHelper ch = ConnectionHelper.GetConnection(UID);

                XElement tn = new XElement("Request");
                tn.Add(new XElement("TableName", TableName));

                Envelope rsp = ch.CallService("UDTService.DDL.GetTableInfo", new Envelope(new XHelper(tn)));

                XElement rspxml = XElement.Parse(rsp.BodyContent.XmlString);
                editor.Text = rspxml.ToString();
            }
            catch (Exception ex)
            {
                RTOut.WriteError(ex);
                MessageBox.Show(ex.Message);
            }
        }
    }
}
