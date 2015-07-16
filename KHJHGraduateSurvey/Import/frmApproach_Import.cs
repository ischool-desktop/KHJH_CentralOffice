using System;
using FISCA.Presentation.Controls;
using FISCA.DSAClient;
using System.Xml.Linq;

namespace JH_KH_GraduateSurvey.Import
{
    public partial class frmApproach_Import : BaseForm
    {
        public int SchoolYear { set; get; }
        public DateTime StartDate { set; get; }
        public DateTime EndDate { set; get; }

        public frmApproach_Import()
        {
            InitializeComponent();

            this.Load += FrmApproach_Import_Load;
        }

        private void FrmApproach_Import_Load(object sender, EventArgs e)
        {
            Connection conn = new Connection();
            try
            {
                conn.EnableSession = false;
                conn.Connect(
                    "j.kh.edu.tw",
                    "centraloffice",
                    FISCA.Authentication.DSAServices.AccessPoint,
                    FISCA.Authentication.DSAServices.AccessPoint
                );
            }
            catch (Exception ex)
            {
                this.lblMessage.Text = ex.Message;
                return;
            }
            //var result = ContractService.GetOpenDate(Connection); FISCA.Authentication.DSAServices.AccessPoint
            //< data >
            //  < uid > 7448 </ uid >
            //  < last_update > 2015 - 07 - 09 17:57:22.974804 </ last_update >
            //  < school_year > 103 </ school_year >
            //  < start_date > 2015 - 07 - 01 00:00:00 </ start_date >
            //  < end_date > 2015 - 07 - 31 00:00:00 </ end_date >
            //</ data >
            //result.Element("school_year").Value;
            XElement result;
            try
            {
                result = ContractService.GetOpenDate(conn);
            }
            catch(Exception ex)
            {
                this.lblMessage.Text = ex.Message;
                return;
            }
            this.SchoolYear = int.Parse(result.Element("school_year").Value);
            this.StartDate = DateTime.Parse(result.Element("start_date").Value);
            this.EndDate = DateTime.Parse(result.Element("end_date").Value);
            this.lblMessage.Text = string.Format("目前填報年度：{0}\n填報期間：{1}~{2}", this.SchoolYear, this.StartDate.ToShortDateString(),this.EndDate.ToShortDateString());
            this.btnNext.Visible = true;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }
    }
}
