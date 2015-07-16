using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;

namespace KHJHCentralOffice
{
    public partial class SimpleDefContent : BaseForm
    {
        public SimpleDefContent()
        {
            InitializeComponent();
        }

        private void UDSContractDefContent_Load(object sender, EventArgs e)
        {

        }

        public void SetXmlContent(string xml)
        {
            editor.Text = xml;
        }
    }
}
