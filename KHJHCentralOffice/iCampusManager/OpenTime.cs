using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace KHJHCentralOffice
{
    /// <summary>
    /// 設定開放時間
    /// </summary>
    public partial class OpenTime : FISCA.Presentation.Controls.BaseForm
    {
        private List<OpenTimeSetting> OpenTimeSettings = null;

        public OpenTime()
        {
            InitializeComponent();
        }

        private void btnExit_Click(object sender, System.EventArgs e)
        {
            this.Close();
        }

        private void OpenTime_Load(object sender, System.EventArgs e)
        {
            grdOpenDate.Rows.Clear();

             OpenTimeSettings = Utility.AccessHelper
                .Select<OpenTimeSetting>();

            foreach (OpenTimeSetting vSetting in OpenTimeSettings)
            {
                grdOpenDate.Rows.Add(
                    vSetting.SurveyYear, 
                    vSetting.StartDate.ToShortDateString(), 
                    vSetting.EndDate.ToShortDateString());
            }
        }

        private void btnSave_Click(object sender, System.EventArgs e)
        {
            bool HasError = false;

            List<int> SurveyYears = new List<int>();

            foreach (DataGridViewRow Row in grdOpenDate.Rows)
            {
                if (!Row.IsNewRow)
                {
                    foreach (DataGridViewCell Cell in Row.Cells)
                    {
                        Cell.ErrorText = string.Empty;

                        //進行學年度檢查
                        if (Cell.ColumnIndex == 0)
                        {
                            int survey_Year;

                            if (int.TryParse("" + Cell.Value, out survey_Year))
                            {
                                if (SurveyYears.Contains(survey_Year))
                                {
                                    Cell.ErrorText = "調查學年度重覆！";
                                    HasError = true;
                                }
                                else
                                    SurveyYears.Add(survey_Year);
                            }
                            else
                            {
                                Cell.ErrorText = "調查學年度請輸入數字！";
                                HasError = true;
                            }
                        }

                        #region 進行日期格式檢查
                        if (Cell.ColumnIndex == 1 || Cell.ColumnIndex == 2)
                        {
                            DateTime vDateTime;

                            if (!DateTime.TryParse("" + Cell.Value, out vDateTime))
                            {
                                Cell.ErrorText = "日期格式錯誤";
                                HasError = true;
                            }
                            else
                                Cell.Value = vDateTime.ToShortDateString();
                        }
                        #endregion
                    }

                    if (string.IsNullOrEmpty(Row.Cells[1].ErrorText) &&
                       string.IsNullOrEmpty(Row.Cells[2].ErrorText))
                    {
                        DateTime dteStart = DateTime.Parse(""+Row.Cells[1].Value);
                        DateTime dteEnd = DateTime.Parse("" + Row.Cells[2].Value);

                        if (dteEnd <= dteStart)
                        {
                            Row.Cells[1].ErrorText = "結束日期必須大於開始日期！";
                            Row.Cells[2].ErrorText = "結束日期必須大於開始日期！";
                            HasError = true;
                        }
                    }
                }
            }

            if (HasError)
            {
                MessageBox.Show("輸入資料有誤，請檢查後再儲存！");
                return;
            }

            try
            {
                #region 先將全部資料刪除
                List<OpenTimeSetting> DeleteRecords = Utility
                    .AccessHelper.Select<OpenTimeSetting>();

                DeleteRecords.ForEach(x => x.Deleted = true);
                DeleteRecords.SaveAll();
                #endregion

                #region 新增資料
                OpenTimeSettings.Clear();

                foreach (DataGridViewRow Row in grdOpenDate.Rows)
                {
                    if (!Row.IsNewRow)
                    {
                        string SurveyYear = "" + Row.Cells[0].Value;
                        string StartDateTime = "" + Row.Cells[1].Value;
                        string EndDateTime = "" + Row.Cells[2].Value;

                        OpenTimeSetting vSetting = new OpenTimeSetting();

                        vSetting.SurveyYear = int.Parse(SurveyYear);
                        vSetting.StartDate = DateTime.Parse(StartDateTime);
                        vSetting.EndDate = DateTime.Parse(EndDateTime);

                        OpenTimeSettings.Add(vSetting);
                    }
                }

                Utility.AccessHelper.SaveAll(OpenTimeSettings);

                MessageBox.Show("儲存成功！");
                #endregion
            }
            catch (Exception ve)
            {
                MessageBox.Show(ve.Message);
            }
        }

        private void grdOpenDate_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}