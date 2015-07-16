using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Aspose.Words;
using FISCA.Data;
using FISCA.DSAClient;
using FISCA.Presentation.Controls;
using FISCA.UDT;

namespace JH_KH_GraduateSurvey.Report
{
    public partial class Approach_Report : BaseForm
    {
        private AccessHelper Access;
        private QueryHelper Query;
        
        public Approach_Report()
        {
            InitializeComponent();

            this.Load += new EventHandler(Form_Load);

            Access = new AccessHelper();
            Query = new QueryHelper();

            this.InitSchoolYear();
        }

        private void Form_Load(object sender, EventArgs e)
        {
            this.circularProgress.Visible = false;
            this.circularProgress.IsRunning = false;
        }

        private void InitSchoolYear()
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
                MessageBox.Show(ex.Message);
                return;
            }
            //var result = ContractService.GetOpenDate(Connection); FISCA.Authentication.DSAServices.AccessPoint
            //< Response >
            //< data >
            //  < uid > 7448 </ uid >
            //  < last_update > 2015 - 07 - 09 17:57:22.974804 </ last_update >
            //  < school_year > 103 </ school_year >
            //  < start_date > 2015 - 07 - 01 00:00:00 </ start_date >
            //  < end_date > 2015 - 07 - 31 00:00:00 </ end_date >
            //</ data >
            //< data >
            //  < uid > 7448 </ uid >
            //  < last_update > 2015 - 07 - 09 17:57:22.974804 </ last_update >
            //  < school_year > 103 </ school_year >
            //  < start_date > 2015 - 07 - 01 00:00:00 </ start_date >
            //  < end_date > 2015 - 07 - 31 00:00:00 </ end_date >
            //</ data >
            //</Response>
            XElement result;
            try
            {
                result = ContractService.GetSurveyYears(conn);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            this.nudSchoolYear.Minimum = result.Descendants("school_year").Min(x => decimal.Parse(x.Value));
            this.nudSchoolYear.Maximum = result.Descendants("school_year").Max(x => decimal.Parse(x.Value));
            this.nudSchoolYear.Value = this.nudSchoolYear.Maximum;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //報表產生完成後，儲存並且開啟
        private void Completed(string inputReportName, Document inputDoc)
        {
            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = inputReportName + DateTime.Now.ToString("yyyy-MM-dd_HH_mm_ss") + ".doc";
            sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            sd.AddExtension = true;
            if (sd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    inputDoc.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);
                    System.Diagnostics.Process.Start(sd.FileName);
                }
                catch
                {
                    MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            string survey_year = this.nudSchoolYear.Value + "";
            this.btnPrint.Enabled = false;
            this.circularProgress.Visible = true;
            this.circularProgress.IsRunning = true;

            Task<Dictionary<string, object>> task = Accessor.ApproachStatistics.Execute(int.Parse(survey_year));
            task.ContinueWith((x) =>
            {                
                this.btnPrint.Enabled = true;
                this.circularProgress.Visible = false;
                this.circularProgress.IsRunning = false;

                if (x.Exception != null)
                    MessageBox.Show(x.Exception.InnerException.Message);
                else
                {
                    MemoryStream template = Accessor.ApproachReportTemplate.Execute(int.Parse(survey_year));
                    Document doc = new Document();
                    Document dataDoc = new Document(template, "", LoadFormat.Doc, "");
                    dataDoc.MailMerge.RemoveEmptyParagraphs = true;
                    doc.Sections.Clear();
                    List<string> keys = new List<string>();
                    List<object> values = new List<object>();
                    Dictionary<string, object> mergeKeyValue = x.Result;
                    foreach (string key in mergeKeyValue.Keys)
                    {
                        keys.Add(key);
                        values.Add(mergeKeyValue[key]);
                    }
                    dataDoc.MailMerge.Execute(keys.ToArray(), values.ToArray());
                    dataDoc.MailMerge.DeleteFields();
                    doc.Sections.Add(doc.ImportNode(dataDoc.Sections[0], true));
                    Completed(survey_year + "學年度國中畢業學生進路調查填報表格", doc);
                }
            }, System.Threading.CancellationToken.None, TaskContinuationOptions.None, TaskScheduler.FromCurrentSynchronizationContext());
        }
    }
}