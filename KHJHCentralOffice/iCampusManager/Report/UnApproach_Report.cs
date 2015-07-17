using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Words;
using FISCA.Presentation.Controls;
using Aspose.Cells;
using System.Xml.Linq;

namespace KHJHCentralOffice
{
    public partial class UnApproach_Report : BaseForm
    {
        private byte[] template;

        public UnApproach_Report(string title, byte[] template)
        {
            InitializeComponent();

            this.Load += new EventHandler(Form_Load);
            this.template = template;
            this.TitleText = title;
            this.Text = title;

            this.InitSchoolYear();
        }

        private void Form_Load(object sender, EventArgs e)
        {
            this.circularProgress.Visible = false;
            this.circularProgress.IsRunning = false;
        }

        private void InitSchoolYear()
        {
            this.nudSchoolYear.Value = decimal.Parse((DateTime.Today.Year - 1912).ToString());
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
            this.circularProgress.Visible = true;
            this.circularProgress.IsRunning = true;

            Task<Workbook> task = Task<Workbook>.Factory.StartNew(() =>
            {
                MemoryStream template = new MemoryStream(this.template);

                Workbook book = new Workbook();
                book.Open(template);

                List<ApproachStatistics> Records = Utility.AccessHelper
                    .Select<ApproachStatistics>("survey_year=" + survey_year);

                if (Records.Count == 0)
                    throw new Exception("本年度無填報資料。");

                List<School> Schools = Utility.AccessHelper.Select<School>();

				int RowIndex = 1;
                foreach (ApproachStatistics record in Records)
                {
                    string SchoolName = string.Empty;
                    XElement elmContent = XElement.Load(new StringReader(record.Content));

                    School vSchool = Schools.Find(x => x.UID.Equals("" + record.RefSchoolID));

                    if (vSchool != null)
                        SchoolName = vSchool.Title;

                    if (elmContent.Element("UnApproachStudents") != null)
                    {
                        foreach (XElement elmStudent in elmContent.Element("UnApproachStudents").Elements("Student"))
                        {
                            book.Worksheets[0].Cells[RowIndex, 0].PutValue(SchoolName);
                            book.Worksheets[0].Cells[RowIndex, 1].PutValue(elmStudent.Element("姓名").Value);

                            if (elmStudent.Element("班級") != null)
                                book.Worksheets[0].Cells[RowIndex, 2].PutValue(elmStudent.Element("班級").Value);
                            else
                                book.Worksheets[0].Cells[RowIndex, 2].PutValue("");
                            book.Worksheets[0].Cells[RowIndex, 3].PutValue(elmStudent.Element("座號").Value);
                            book.Worksheets[0].Cells[RowIndex, 4].PutValue(elmStudent.Element("未升學未就業動向").Value);
							book.Worksheets[0].Cells[RowIndex, 5].PutValue(string.IsNullOrWhiteSpace(elmStudent.Element("是否需要教育部協助").Value) ? "否" : elmStudent.Element("是否需要教育部協助").Value);
                            book.Worksheets[0].Cells[RowIndex, 6].PutValue(elmStudent.Element("備註").Value);
							book.Worksheets[0].Cells[RowIndex, 6].Style.IsTextWrapped = true;
							book.Worksheets[0].Cells[RowIndex, 0].Style.ShrinkToFit = true;
							book.Worksheets[0].Cells[RowIndex, 6].Style.HorizontalAlignment = TextAlignmentType.Left;
							this.DrawBorder(book.Worksheets[0].Cells[RowIndex, 0]);
							this.DrawBorder(book.Worksheets[0].Cells[RowIndex, 1]);
							this.DrawBorder(book.Worksheets[0].Cells[RowIndex, 2]);
							this.DrawBorder(book.Worksheets[0].Cells[RowIndex, 3]);
							this.DrawBorder(book.Worksheets[0].Cells[RowIndex, 4]);
							this.DrawBorder(book.Worksheets[0].Cells[RowIndex, 5]);
                            this.DrawBorder(book.Worksheets[0].Cells[RowIndex, 6]);
							book.Worksheets[0].AutoFitRow(RowIndex);
							book.Worksheets[0].Cells.SetRowHeight(RowIndex, book.Worksheets[0].Cells.GetRowHeight(RowIndex) * 10 / 7);
                            RowIndex++;
                        }
                    }
                }

                return book;
            });
            task.ContinueWith((x) =>
            {
                this.btnPrint.Enabled = true;
                this.circularProgress.Visible = false;
                this.circularProgress.IsRunning = false;

                if (x.Exception != null)
                    MessageBox.Show(x.Exception.InnerException.Message);
                else
                    Completed(this.TitleText, x.Result);
            }, System.Threading.CancellationToken.None, TaskContinuationOptions.None, TaskScheduler.FromCurrentSynchronizationContext());
        }

		private void DrawBorder(Aspose.Cells.Cell cell)
		{
			Aspose.Cells.Style style = cell.Style;

			style.Borders[Aspose.Cells.BorderType.TopBorder].LineStyle = CellBorderType.Thin;
			style.Borders[Aspose.Cells.BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
			style.Borders[Aspose.Cells.BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
			style.Borders[Aspose.Cells.BorderType.RightBorder].LineStyle = CellBorderType.Thin;

			cell.Style = style;
		}
    }
}