using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Campus.Windows;
using EMBA.Validator;
using FISCA.Data;
using FISCA.DSAClient;
using FISCA.Permission;
using FISCA.UDT;

namespace JH_KH_GraduateSurvey.DetailContent
{
    [FeatureCode("ischool.jh_kh.detail_content.graduate_survey_approach", "畢業學生進路")]
    public partial class Approach_DetailContent : FISCA.Presentation.DetailContent
    {
        //  驗證資料物件
        private ErrorProvider _Errors;

        //  背景載入 UDT 資料物件
        private BackgroundWorker _BGWLoadData;
        private BackgroundWorker _BGWSaveData;

        //  監控 UI 資料變更
        private ChangeListener _Listener;

        //  正在下載的資料之主鍵，用於檢查是否下載他人資料，若 _RunningKey != PrimaryKey 就再下載乙次
        private string _RunningKey;

        private AccessHelper Access;
        private QueryHelper Query;
        private bool form_loaded;

        //  填報資料
        private Dictionary<decimal, IEnumerable<string>> dicSurveyFields;
        private decimal CurrentSchoolYear;
        private decimal SurveyYear;

        public Approach_DetailContent()
        {
            InitializeComponent();

            Access = new AccessHelper();
            Query = new QueryHelper();
            dicSurveyFields = new Dictionary<decimal, IEnumerable<string>>();

            this.Group = "畢業學生進路";
            _RunningKey = "";

            this.Load += new EventHandler(Form_Load);
            this.form_loaded = false;
            _Errors = new ErrorProvider();
            _Listener = new ChangeListener();
            _Listener.Add(new DataGridViewSource(this.dgvData));
            _Listener.Add(new TextBoxSource(this.txtMemo));
            _Listener.Add(new NumericUpDownSource(this.txtSurveyYear));
            _Listener.StatusChanged += new EventHandler<ChangeEventArgs>(Listener_StatusChanged);

            this.dgvData.CellEnter += new DataGridViewCellEventHandler(dgvData_CellEnter);
            this.dgvData.CurrentCellDirtyStateChanged += new EventHandler(dgvData_CurrentCellDirtyStateChanged);
            this.dgvData.DataError += new DataGridViewDataErrorEventHandler(dgvData_DataError);
            this.dgvData.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dgvData_ColumnHeaderMouseClick);
            this.dgvData.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dgvData_RowHeaderMouseClick);
            this.dgvData.MouseClick += new System.Windows.Forms.MouseEventHandler(this.dgvData_MouseClick);

            _BGWLoadData = new BackgroundWorker();
            _BGWLoadData.DoWork += new DoWorkEventHandler(_BGWLoadData_DoWork);
            _BGWLoadData.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BGWLoadData_RunWorkerCompleted);

            _BGWSaveData = new BackgroundWorker();
            _BGWSaveData.DoWork += new DoWorkEventHandler(_BGWSaveData_DoWork);
            _BGWSaveData.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BGWSaveData_RunWorkerCompleted);
        }

        private void Form_Load(object sender, EventArgs e)
        {
            this.InitSchoolYear();
            this.form_loaded = true;

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
            this.SurveyYear = int.Parse(result.Element("school_year").Value);
            DateTime StartDate = DateTime.Parse(result.Element("start_date").Value);
            DateTime EndDate = DateTime.Parse(result.Element("end_date").Value);
            this.lblMessage.Text = string.Format("目前填報年度：{0}\n填報期間：{1}~{2}", this.SurveyYear, StartDate.ToShortDateString(), EndDate.ToShortDateString());
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
            this.txtSurveyYear.Minimum = result.Descendants("school_year").Min(x => decimal.Parse(x.Value));
            this.txtSurveyYear.Maximum = result.Descendants("school_year").Max(x => decimal.Parse(x.Value));
            this.txtSurveyYear.Value = this.txtSurveyYear.Maximum;
            this.CurrentSchoolYear = this.txtSurveyYear.Value;
        }

        private void dgvData_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            this.dgvData.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void dgvData_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.RowIndex >= 0 && e.ColumnIndex == 1 && dgvData.SelectedCells.Count == 1)
            //{
            //    dgvData.BeginEdit(true);  //Raise EditingControlShowing Event !
            //    if (dgvData.CurrentCell != null && dgvData.CurrentCell.GetType().ToString() == "System.Windows.Forms.DataGridViewComboBoxCell")
            //        (dgvData.EditingControl as ComboBox).DroppedDown = true;  //自動拉下清單
            //}
            dgvData.BeginEdit(true);
        }

        private void dgvData_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            dgvData.CurrentCell = null;
            dgvData.Rows[e.RowIndex].Selected = true;
        }

        private void dgvData_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            dgvData.CurrentCell = null;
        }

        private void dgvData_MouseClick(object sender, MouseEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;
            DataGridView.HitTestInfo hit = dgv.HitTest(e.X, e.Y);

            if (hit.Type == DataGridViewHitTestType.TopLeftHeader)
            {
                dgvData.CurrentCell = null;
                dgvData.SelectAll();
            }
        }

        private void dgvData_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.Cancel = true;
        }

        private void Listener_StatusChanged(object sender, ChangeEventArgs e)
        {
            if (UserAcl.Current[typeof(Approach_DetailContent)].Editable)
                SaveButtonVisible = e.Status == ValueStatus.Dirty;
            else
                this.SaveButtonVisible = false;

            CancelButtonVisible = e.Status == ValueStatus.Dirty;
        }

        private void _BGWLoadData_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                this.RefreshUI(null);
                //this.lblMessage.Text = e.Error.Message;
                this.Loading = false;
                return;
            }

            if (_RunningKey != PrimaryKey)
            {
                this.Loading = true;
                this._RunningKey = PrimaryKey;
                this._BGWLoadData.RunWorkerAsync();
            }
            else
            {
                this.RefreshUI(e.Result);
            }
        }

        private void _BGWLoadData_DoWork(object sender, DoWorkEventArgs e)
        {
            string SQL = string.Format(@"select 
                q1 as 升學與就業情形,
                q2 as 升學：就讀學校情形, 
                q3 as 升學：學制別, 
                q4 as 升學：入學方式, 
                q5 as 未升學未就業：動向, 
                q6 as 是否需要教育部協助,
                memo as 備註,
                survey_year as 填報學年度 
                from $ischool.jh_kh.graduate_survey_approach where ref_student_id={0} order by last_update_time DESC", this._RunningKey);

            DataTable dataTable = Query.Select(SQL);

            e.Result = dataTable;
        }

        //  檢視不同資料項目即呼叫此方法，PrimaryKey 為資料項目的 Key 值。
        protected override void OnPrimaryKeyChanged(EventArgs e)
        {
            if (!this._BGWLoadData.IsBusy)
            {
                this.Loading = true;
                this._RunningKey = PrimaryKey;
                this._BGWLoadData.RunWorkerAsync();
            }
        }

        //  更新資料項目內 UI 的資料
        private void RefreshUI(object result)
        {
            _Listener.SuspendListen();

            this.dgvData.EndEdit();
            this.dgvData.Rows.Clear();
            this.txtSurveyYear.Value = this.CurrentSchoolYear;
            this.txtMemo.Text = string.Empty;
            this.ClearError();

            if (result == null)
            {
                ResetOverrideButton();
                return;
            }
            else
            {
                //this.lblMessage.Text = string.Empty;
                this.txtMemo.Text = string.Empty;
            }

            DataTable dataTable = result as DataTable;
            if (dataTable == null)
            {
                ResetOverrideButton();
                //this.lblMessage.Text = "異常發生。";
                return;
            }
            DataRow row = dataTable.NewRow();
            if (dataTable.Rows.Count == 0)
            {
                dataTable.ImportRow(row);
            }
            else
            {
                row = dataTable.Rows[0];
                this.txtSurveyYear.Text = row["填報學年度"] + "";
                this.txtMemo.Text = row["備註"] + "";
            }

            foreach (DataColumn column in dataTable.Columns)
            {
                if (column.ColumnName == "填報學年度" || column.ColumnName == "備註")
                    continue;

                List<object> source = new List<object>();

                source.Add(column.ColumnName + "");
                source.Add(row[column] + "");

                int idx = this.dgvData.Rows.Add(source.ToArray());
            }
            this.dgvData.CurrentCell = null;
            this.Loading = false;
            ResetOverrideButton();
        }

        protected override void OnCancelButtonClick(EventArgs e)
        {
            if (!_BGWLoadData.IsBusy)
            {
                this.ClearError();
                this._BGWLoadData.RunWorkerAsync();
            }
        }

        private void ClearError()
        {
            this._Errors.SetError(this.txtSurveyYear, string.Empty);
            this._Errors.SetError(this.txtMemo, string.Empty);
            this.dgvData.EndEdit();
            this._Errors.SetError(this.dgvData, string.Empty);
            foreach (DataGridViewRow row in this.dgvData.Rows)
            {
                if (row.IsNewRow)
                    continue;

                row.Cells[1].ErrorText = string.Empty;
            }
        }

        private bool Is_Validated()
        {
            bool is_validated = true;

            this.ClearError();

            string string_Q1 = this.dgvData.Rows[0].Cells[1].Value + "";
            string string_Q2 = this.dgvData.Rows[1].Cells[1].Value + "";
            string string_Q3 = this.dgvData.Rows[2].Cells[1].Value + "";
            string string_Q4 = this.dgvData.Rows[3].Cells[1].Value + "";
            string string_Q5 = this.dgvData.Rows[4].Cells[1].Value + "";
            string string_Q6 = this.dgvData.Rows[5].Cells[1].Value + "";

            int intPrimaryKey = int.Parse(PrimaryKey);
            Dictionary<int, Dictionary<string, string>> Data = new Dictionary<int, Dictionary<string, string>>();
            Data.Add(intPrimaryKey, new Dictionary<string, string>());
            Data[intPrimaryKey].Add("升學與就業情形", string_Q1);
            Data[intPrimaryKey].Add("升學：就讀學校情形", string_Q2);
            Data[intPrimaryKey].Add("升學：學制別", string_Q3);
            Data[intPrimaryKey].Add("升學：入學方式", string_Q4);
            Data[intPrimaryKey].Add("未升學未就業：動向", string_Q5);
            Data[intPrimaryKey].Add("是否需要教育部協助", string_Q6);
            Data[intPrimaryKey].Add("備註", txtMemo.Text.Trim());

            Dictionary<int, List<MessageItem>> dicMessages = Accessor.ApproachValidate.Execute((int)this.CurrentSchoolYear, Data);

            List<string> ErrorMessages = new List<string>();
            foreach (int key in dicMessages.Keys)
            {
                if (dicMessages[key].Count > 0)
                {
                    ErrorMessages.AddRange(dicMessages[key].Select(x => x.Message));
                }
            }
            if (ErrorMessages.Count > 0)
            {
                this._Errors.SetError(this.dgvData, string.Join("\n", ErrorMessages));
                is_validated = false;
            }
            else
            {
                this._Errors.SetError(this.dgvData, string.Empty);
            }
            return is_validated;
        }

        protected override void OnSaveButtonClick(EventArgs e)
        {
            this.dgvData.EndEdit();
            this.dgvData.CurrentCell = null;
            if (!Is_Validated())
            {
                MessageBox.Show("請先修正錯誤。");
                return;
            }

            UDT.Approach approach = new UDT.Approach();

            string string_Q1 = this.dgvData.Rows[0].Cells[1].Value + "";
            string string_Q2 = this.dgvData.Rows[1].Cells[1].Value + "";
            string string_Q3 = this.dgvData.Rows[2].Cells[1].Value + "";
            string string_Q4 = this.dgvData.Rows[3].Cells[1].Value + "";
            string string_Q5 = this.dgvData.Rows[4].Cells[1].Value + "";
            string string_Q6 = this.dgvData.Rows[5].Cells[1].Value + "";

            //int int_QQ;

            #region 儲存UDT資料

            //approach.StudentID = int.Parse(this.PrimaryKey);
            //approach.SurveyYear = int.Parse(this.txtSurveyYear.Text.Trim());
            //approach.LastUpdateTime = DateTime.Now;
            //         approach.Q6 = string_Q6;
            //         approach.Memo = txtMemo.Text.Trim();

            //         approach.Q1 = int.Parse(this.dgvData.Rows[0].Cells[1].Value + "");
            //         if (int.TryParse(string_Q2, out int_QQ))
            //             approach.Q2 = int_QQ;
            //         else
            //             approach.Q2 = null;
            //         if (int.TryParse(string_Q3, out int_QQ))
            //             approach.Q3 = int_QQ;
            //         else
            //             approach.Q3 = null;
            //         if (int.TryParse(string_Q4, out int_QQ))
            //             approach.Q4 = int_QQ;
            //         else
            //             approach.Q4 = null;
            //         if (int.TryParse(string_Q5, out int_QQ))
            //             approach.Q5 = int_QQ;
            //         else
            //             approach.Q5 = null;


            Dictionary<string, Dictionary<string, string>> Data = new Dictionary<string, Dictionary<string, string>>();
            Data.Add(PrimaryKey, new Dictionary<string, string>());
            Data[PrimaryKey].Add("升學與就業情形", string_Q1);
            Data[PrimaryKey].Add("升學：就讀學校情形", string_Q2);
            Data[PrimaryKey].Add("升學：學制別", string_Q3);
            Data[PrimaryKey].Add("升學：入學方式", string_Q4);
            Data[PrimaryKey].Add("未升學未就業：動向", string_Q5);
            Data[PrimaryKey].Add("是否需要教育部協助", string_Q6);
            Data[PrimaryKey].Add("備註", txtMemo.Text.Trim());
            #endregion
            this._BGWSaveData.RunWorkerAsync(Data);
        }

        private void _BGWSaveData_DoWork(object sender, DoWorkEventArgs e)
        {
            Dictionary<string, Dictionary<string, string>> Data = e.Argument as Dictionary<string, Dictionary<string, string>>;
            Accessor.ApproachSave.Execute((int)CurrentSchoolYear, Data);
            //SaveUDT(Data);
        }

        private void _BGWSaveData_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this._BGWLoadData.RunWorkerAsync();
        }

        //private void SaveUDT(UDT.Approach approach)
        //{
        //    List<UDT.Approach> records = Access.Select<UDT.Approach>(string.Format("ref_student_id={0}", this._RunningKey));

        //    if (records.Count > 0)
        //        records.ForEach(x => x.Deleted = true);

        //    records.Add(approach);
        //    records.SaveAll();
        //}

        private void ResetOverrideButton()
        {
            SaveButtonVisible = false;
            CancelButtonVisible = false;
            this.ClearError();

            _Listener.Reset();
            _Listener.ResumeListen();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("敬請您再次確認此筆填報資料為誤填，確應刪除，否則請您按「取消」鈕，停止「刪除」。", "危險動作", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                List<UDT.Approach> record = this.Access.Select<UDT.Approach>(string.Format("ref_student_id={0}", PrimaryKey));
                if (record.Count == 0)
                {
                    MessageBox.Show("無填報資料可刪除。");
                }
                else
                {
                    record.ForEach(x => x.Deleted = true);
                    record.SaveAll();
                    this._BGWLoadData.RunWorkerAsync();
                    MessageBox.Show("填報資料已刪除。");
                }
            }
            else
            {
                MessageBox.Show("已取消，填報資料未刪除。");
            }
        }

        private void txtSurveyYear_ValueChanged(object sender, EventArgs e)
        {
            this.CurrentSchoolYear = this.txtSurveyYear.Value;
        }
    }
}
