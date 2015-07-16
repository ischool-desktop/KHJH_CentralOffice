using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Words;
using FISCA.Presentation.Controls;
using Aspose.Cells;
using System.Data;
using System.Linq;

namespace KHJHCentralOffice
{
    public partial class Approach_Export : BaseForm
    {
		public Approach_Export()
        {
            InitializeComponent();

            this.Load += new EventHandler(Form_Load);

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

        private void btnPrint_Click(object sender, EventArgs e)
        {
            string survey_year = this.nudSchoolYear.Value + "";
            this.btnPrint.Enabled = false;
            this.circularProgress.Visible = true;
            this.circularProgress.IsRunning = true;

            Task<Workbook> task = Accessor.ApproachExport.Execute(int.Parse(survey_year));
            task.ContinueWith((x) =>
			{
				this.btnPrint.Enabled = true;
				this.circularProgress.Visible = false;
				this.circularProgress.IsRunning = false;
				if (x.Exception != null)
				{
					MessageBox.Show(x.Exception.InnerException.Message);
					return;
				}

				SaveFileDialog sd = new SaveFileDialog();
				sd.Title = "另存新檔";
				sd.FileName = "匯出" + survey_year + "學年度畢業學生進路統計分析資料" + DateTime.Now.ToString(" yyyy-MM-dd_HH_mm_ss") + ".xls";
				sd.Filter = "Excel 2003 相容檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
				if (sd.ShowDialog() == DialogResult.OK)
				{
					try
					{
						x.Result.Save(sd.FileName, FileFormatType.Excel2003);
						System.Diagnostics.Process.Start(sd.FileName);
					}
					catch
					{
						MessageBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
						return;
					}
				}
            }, System.Threading.CancellationToken.None, TaskContinuationOptions.None, TaskScheduler.FromCurrentSynchronizationContext());
        }
    }
}