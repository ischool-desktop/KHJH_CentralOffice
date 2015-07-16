namespace KHJHCentralOffice
{
    partial class GraduateSurveyVagrant
    {
        /// <summary> 
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 元件設計工具產生的程式碼

        /// <summary> 
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
        /// 修改這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.grdVagrant = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.colSchoolYear = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colInJob = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colInSchool = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colPrepareSchool = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colPrepareJob = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colInTraining = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colInHome = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colNoPlan = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colDisAppearance = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colOther = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.grdVagrant)).BeginInit();
            this.SuspendLayout();
            // 
            // grdVagrant
            // 
            this.grdVagrant.AllowUserToAddRows = false;
            this.grdVagrant.AllowUserToDeleteRows = false;
            this.grdVagrant.AllowUserToResizeColumns = false;
            this.grdVagrant.BackgroundColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.grdVagrant.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grdVagrant.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colSchoolYear,
            this.colInJob,
            this.colInSchool,
            this.colPrepareSchool,
            this.colPrepareJob,
            this.colInTraining,
            this.colInHome,
            this.colNoPlan,
            this.colDisAppearance,
            this.colOther});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.grdVagrant.DefaultCellStyle = dataGridViewCellStyle1;
            this.grdVagrant.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.grdVagrant.Location = new System.Drawing.Point(13, 8);
            this.grdVagrant.Name = "grdVagrant";
            this.grdVagrant.ReadOnly = true;
            this.grdVagrant.RowTemplate.Height = 24;
            this.grdVagrant.Size = new System.Drawing.Size(524, 169);
            this.grdVagrant.TabIndex = 0;
            // 
            // colSchoolYear
            // 
            this.colSchoolYear.FillWeight = 80F;
            this.colSchoolYear.HeaderText = "學年度";
            this.colSchoolYear.Name = "colSchoolYear";
            this.colSchoolYear.ReadOnly = true;
            this.colSchoolYear.Width = 80;
            // 
            // colInJob
            // 
            this.colInJob.FillWeight = 80F;
            this.colInJob.HeaderText = "已就業";
            this.colInJob.Name = "colInJob";
            this.colInJob.ReadOnly = true;
            this.colInJob.Width = 80;
            // 
            // colInSchool
            // 
            this.colInSchool.HeaderText = "已就學";
            this.colInSchool.Name = "colInSchool";
            this.colInSchool.ReadOnly = true;
            // 
            // colPrepareSchool
            // 
            this.colPrepareSchool.HeaderText = "準備升學";
            this.colPrepareSchool.Name = "colPrepareSchool";
            this.colPrepareSchool.ReadOnly = true;
            // 
            // colPrepareJob
            // 
            this.colPrepareJob.HeaderText = "找工作";
            this.colPrepareJob.MinimumWidth = 80;
            this.colPrepareJob.Name = "colPrepareJob";
            this.colPrepareJob.ReadOnly = true;
            this.colPrepareJob.Width = 80;
            // 
            // colInTraining
            // 
            this.colInTraining.HeaderText = "參加職訓";
            this.colInTraining.Name = "colInTraining";
            this.colInTraining.ReadOnly = true;
            // 
            // colInHome
            // 
            this.colInHome.HeaderText = "家務勞動";
            this.colInHome.Name = "colInHome";
            this.colInHome.ReadOnly = true;
            // 
            // colNoPlan
            // 
            this.colNoPlan.HeaderText = "尚未規劃";
            this.colNoPlan.Name = "colNoPlan";
            this.colNoPlan.ReadOnly = true;
            // 
            // colDisAppearance
            // 
            this.colDisAppearance.FillWeight = 80F;
            this.colDisAppearance.HeaderText = "失聯";
            this.colDisAppearance.Name = "colDisAppearance";
            this.colDisAppearance.ReadOnly = true;
            this.colDisAppearance.Width = 80;
            // 
            // colOther
            // 
            this.colOther.FillWeight = 80F;
            this.colOther.HeaderText = "其他";
            this.colOther.Name = "colOther";
            this.colOther.ReadOnly = true;
            this.colOther.Width = 80;
            // 
            // GraduateSurveyVagrant
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.grdVagrant);
            this.Name = "GraduateSurveyVagrant";
            this.Size = new System.Drawing.Size(550, 190);
            this.Load += new System.EventHandler(this.BasicInfoItem_Load);
            ((System.ComponentModel.ISupportInitialize)(this.grdVagrant)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.Controls.DataGridViewX grdVagrant;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSchoolYear;
        private System.Windows.Forms.DataGridViewTextBoxColumn colInJob;
        private System.Windows.Forms.DataGridViewTextBoxColumn colInSchool;
        private System.Windows.Forms.DataGridViewTextBoxColumn colPrepareSchool;
        private System.Windows.Forms.DataGridViewTextBoxColumn colPrepareJob;
        private System.Windows.Forms.DataGridViewTextBoxColumn colInTraining;
        private System.Windows.Forms.DataGridViewTextBoxColumn colInHome;
        private System.Windows.Forms.DataGridViewTextBoxColumn colNoPlan;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDisAppearance;
        private System.Windows.Forms.DataGridViewTextBoxColumn colOther;

    }
}
