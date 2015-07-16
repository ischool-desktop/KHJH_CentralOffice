using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Cells;
using Aspose.Words;
using FISCA.Presentation.Controls;

namespace KHJHCentralOffice
{
    public partial class UnApproach_Check : BaseForm
    {
        private byte[] template;
        private string title;

        public UnApproach_Check()
        {
            InitializeComponent();

            this.Load += new EventHandler(Form_Load);
            this.InitSchoolYear();
        }

        private void Form_Load(object sender, EventArgs e)
        {
        }

        private void InitSchoolYear()
        {
            this.nudSchoolYear.Value = decimal.Parse((DateTime.Today.Year - 1911).ToString());
            this.nudSchoolYear.Value -= 1;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //報表產生完成後，儲存並且開啟
        private void Completed(string inputReportName, Workbook inputDoc)
        {
            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = inputReportName + DateTime.Now.ToString("yyyy-MM-dd_HH_mm_ss") + ".xls";
            sd.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
            sd.AddExtension = true;
            if (sd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    inputDoc.Save(sd.FileName);
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

            Task<List<School>> task = Task<List<School>>.Factory.StartNew(() =>
            {
                List<School> Schools = Utility.AccessHelper.Select<School>();
                List<ApproachStatistics> Approachs = Utility.AccessHelper
                    .Select<ApproachStatistics>("survey_year="+survey_year);

                List<School> result = new List<School>();

                foreach (School vSchool in Schools)
                {
                    if (Approachs.Find(x => ("" + x.RefSchoolID).Equals(vSchool.UID))==null)
                    {
                        result.Add(vSchool);
                    }
                }

                return result;
            });
            task.ContinueWith((x) =>
            {
                this.btnPrint.Enabled = true;

                if (x.Exception != null)
                    MessageBox.Show(x.Exception.InnerException.Message);
                else
                {
                    List<School> result = x.Result as List<School>;

                    grdSchool.Rows.Clear();

                    foreach (School vSchool in result)
                    {
                        if (!string.IsNullOrEmpty(vSchool.Title))
                            grdSchool.Rows.Add(vSchool.Title);
                    }
                }

            }, System.Threading.CancellationToken.None, TaskContinuationOptions.None, TaskScheduler.FromCurrentSynchronizationContext());
        }
    }
}