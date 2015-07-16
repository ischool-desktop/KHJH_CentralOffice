using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Words;
using FISCA.Presentation.Controls;

namespace KHJHCentralOffice
{
    public partial class Approach_Report : BaseForm
    {
        private byte[] template;
        private string title;
        private Accessor.ApproachReportTemplate.ReportType ReportType;
        
        public Approach_Report(string title, Accessor.ApproachReportTemplate.ReportType ReportType)
        {
            InitializeComponent();

            this.Load += new EventHandler(Form_Load);
            this.title = title;
            this.TitleText = title;
            this.ReportType = ReportType;

            this.InitSchoolYear();
        }

        private void Form_Load(object sender, EventArgs e)
        {
            this.circularProgress.Visible = false;
            this.circularProgress.IsRunning = false;
        }

        private void InitSchoolYear()
        {
            List<OpenTimeSetting> OpenTimes = Utility.AccessHelper.Select<OpenTimeSetting>();
            if (OpenTimes.Count > 0)
            {
                this.nudSchoolYear.Minimum = OpenTimes.Select(x => x.SurveyYear).Min();
                this.nudSchoolYear.Maximum = OpenTimes.Select(x => x.SurveyYear).Max();

                this.nudSchoolYear.Value = this.nudSchoolYear.Maximum;
            }
            else
            {
                this.nudSchoolYear.Minimum = this.nudSchoolYear.Maximum = this.nudSchoolYear.Value = decimal.Parse((DateTime.Today.Year - 1912).ToString());
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void MailMerge_MergeField(object sender, Aspose.Words.Reporting.MergeFieldEventArgs e)
        {
            #region 科目成績
             
            #endregion
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

            Task<Dictionary<string, object>> task = Accessor.ApproachReport.Execute(int.Parse(survey_year));
            task.ContinueWith((x) =>
            {
                this.btnPrint.Enabled = true;
                this.circularProgress.Visible = false;
                this.circularProgress.IsRunning = false;

                if (x.Exception != null)
                    MessageBox.Show(x.Exception.InnerException.Message);
                else
                {
                    MemoryStream template = Accessor.ApproachReportTemplate.Execute(int.Parse(survey_year), this.ReportType);
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
                    Completed(survey_year + this.TitleText, doc);
                }
            }, System.Threading.CancellationToken.None, TaskContinuationOptions.None, TaskScheduler.FromCurrentSynchronizationContext());
        }
    }
}