using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
//using FISCA.Data;
using FISCA.DSAClient;
using FISCA.Presentation.Controls;
using FISCA.UDT;
using K12.Data;
using JH_KH_GraduateSurvey.Report;

namespace JH_KH_GraduateSurvey
{
    public partial class Approach_Upload : BaseForm
    {
        private AccessHelper Access;
        private int CurrentSurveyYear;
        //private QueryHelper Query;

        public Approach_Upload()
        {
            InitializeComponent();

            this.Load += new EventHandler(Form_Load);

            Access = new AccessHelper();
            //Query = new QueryHelper();

            this.InitSchoolYear();
        }

        private void Form_Load(object sender, EventArgs e)
        {
            this.circularProgress.Visible = false;
            this.circularProgress.IsRunning = false;
        }

        private void InitSchoolYear()
        {
            this.btnPrint.Enabled = false;
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
            catch (Exception ex)
            {
                this.lblMessage.Text = ex.Message;
                return;
            }
            CurrentSurveyYear = int.Parse(result.Element("school_year").Value);
            DateTime StartDate = DateTime.Parse(result.Element("start_date").Value);
            DateTime EndDate = DateTime.Parse(result.Element("end_date").Value);
            this.lblMessage.Text = string.Format("目前填報年度：{0}\n填報期間：{1}~{2}", CurrentSurveyYear, StartDate.ToShortDateString(), EndDate.ToShortDateString());
            this.btnPrint.Enabled = true;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            this.btnPrint.Enabled = false;
            this.circularProgress.Visible = true;
            this.circularProgress.IsRunning = true;

            Task<Dictionary<string, object>> task = Accessor.ApproachStatistics.Execute(this.CurrentSurveyYear);
            task.ContinueWith((x) =>
            {
                if (x.Exception != null)
                {
                    MessageBox.Show(x.Exception.InnerException.Message);
                    goto TheEnd;
                }
                else
                {
                    List<string> keys = new List<string>();
                    List<object> values = new List<object>();
                    Dictionary<string, object> mergeKeyValue = x.Result;
                    if (MessageBox.Show("您是否確認上傳" + mergeKeyValue["筆數"] + "筆記錄？", "確認上傳？", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                    {
                        Upload(mergeKeyValue);
                    }
                }
            TheEnd:
                this.btnPrint.Enabled = true;
                this.circularProgress.Visible = false;
                this.circularProgress.IsRunning = false;

            }, System.Threading.CancellationToken.None, TaskContinuationOptions.None, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private void Upload(Dictionary<string, object> mergeKeyValue)
        {
            Connection Connection = new FISCA.DSAClient.Connection();
            try
            {
                Connection.EnableSession = false;
                Connection.Connect(
                "j.kh.edu.tw",
                "centraloffice",
                FISCA.Authentication.DSAServices.AccessPoint,
                FISCA.Authentication.DSAServices.AccessPoint);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }            

            #region 取得學校系統編號
            XElement elmSchool = ContractService
                .GetSchool(Connection, FISCA.Authentication.DSAServices.AccessPoint)
                .Element("School");

            if (elmSchool == null)
            {
                MessageBox.Show("學校不在局端清單中，無法上傳！");
                return;
            }

            //取得學校在局端的系統編號
            string SchoolID = elmSchool.Element("Uid").Value;
            #endregion

            #region 上傳統計資料
            try
            {
                ContractService.UploadApproach(Connection,
                    SchoolID,
                    "" + CurrentSurveyYear,
                    mergeKeyValue);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            #endregion
            
            MessageBox.Show("上傳成功！");
        }
    }
}