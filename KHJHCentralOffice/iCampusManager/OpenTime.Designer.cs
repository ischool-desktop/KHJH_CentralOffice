namespace KHJHCentralOffice
{
    partial class OpenTime
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.btnSave = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.grdOpenDate = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.colSurveyYear = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colStartDate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colEndDate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.grdOpenDate)).BeginInit();
            this.SuspendLayout();
            // 
            // btnSave
            // 
            this.btnSave.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSave.BackColor = System.Drawing.Color.Transparent;
            this.btnSave.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSave.Location = new System.Drawing.Point(223, 153);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnSave.TabIndex = 0;
            this.btnSave.Text = "儲存";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(314, 153);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 23);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 2;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // grdOpenDate
            // 
            this.grdOpenDate.BackgroundColor = System.Drawing.Color.White;
            this.grdOpenDate.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grdOpenDate.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colSurveyYear,
            this.colStartDate,
            this.colEndDate});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.grdOpenDate.DefaultCellStyle = dataGridViewCellStyle1;
            this.grdOpenDate.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.grdOpenDate.Location = new System.Drawing.Point(14, 15);
            this.grdOpenDate.MultiSelect = false;
            this.grdOpenDate.Name = "grdOpenDate";
            this.grdOpenDate.RowTemplate.Height = 24;
            this.grdOpenDate.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.grdOpenDate.Size = new System.Drawing.Size(376, 122);
            this.grdOpenDate.TabIndex = 3;
            this.grdOpenDate.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.grdOpenDate_CellContentClick);
            // 
            // colSurveyYear
            // 
            this.colSurveyYear.HeaderText = "調查年度";
            this.colSurveyYear.Name = "colSurveyYear";
            this.colSurveyYear.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            // 
            // colStartDate
            // 
            this.colStartDate.HeaderText = "開始日期";
            this.colStartDate.Name = "colStartDate";
            // 
            // colEndDate
            // 
            this.colEndDate.HeaderText = "結束日期";
            this.colEndDate.Name = "colEndDate";
            // 
            // OpenTime
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(402, 188);
            this.Controls.Add(this.grdOpenDate);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnSave);
            this.MaximumSize = new System.Drawing.Size(410, 215);
            this.MinimumSize = new System.Drawing.Size(410, 215);
            this.Name = "OpenTime";
            this.Text = "";
            this.TitleText = "開放時間";
            this.Load += new System.EventHandler(this.OpenTime_Load);
            ((System.ComponentModel.ISupportInitialize)(this.grdOpenDate)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.ButtonX btnSave;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.DotNetBar.Controls.DataGridViewX grdOpenDate;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSurveyYear;
        private System.Windows.Forms.DataGridViewTextBoxColumn colStartDate;
        private System.Windows.Forms.DataGridViewTextBoxColumn colEndDate;
    }
}