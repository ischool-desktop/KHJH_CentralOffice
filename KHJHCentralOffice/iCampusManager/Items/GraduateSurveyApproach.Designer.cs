namespace KHJHCentralOffice
{
    partial class GraduateSurveyApproach
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.cmbSurveyYear = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.grdContent = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.colField = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.grdContent)).BeginInit();
            this.SuspendLayout();
            // 
            // cmbSurveyYear
            // 
            this.cmbSurveyYear.DisplayMember = "Text";
            this.cmbSurveyYear.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cmbSurveyYear.FormattingEnabled = true;
            this.cmbSurveyYear.ItemHeight = 19;
            this.cmbSurveyYear.Location = new System.Drawing.Point(23, 9);
            this.cmbSurveyYear.Name = "cmbSurveyYear";
            this.cmbSurveyYear.Size = new System.Drawing.Size(121, 25);
            this.cmbSurveyYear.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cmbSurveyYear.TabIndex = 0;
            this.cmbSurveyYear.SelectedIndexChanged += new System.EventHandler(this.cmbSurveyYear_SelectedIndexChanged);
            // 
            // grdContent
            // 
            this.grdContent.AllowUserToAddRows = false;
            this.grdContent.AllowUserToDeleteRows = false;
            this.grdContent.BackgroundColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.grdContent.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.grdContent.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grdContent.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colField,
            this.colValue});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.grdContent.DefaultCellStyle = dataGridViewCellStyle2;
            this.grdContent.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.grdContent.Location = new System.Drawing.Point(23, 52);
            this.grdContent.Name = "grdContent";
            this.grdContent.ReadOnly = true;
            this.grdContent.RowHeadersVisible = false;
            this.grdContent.RowTemplate.Height = 24;
            this.grdContent.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.grdContent.Size = new System.Drawing.Size(499, 415);
            this.grdContent.TabIndex = 1;
            // 
            // colField
            // 
            this.colField.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.colField.HeaderText = "欄位";
            this.colField.Name = "colField";
            this.colField.ReadOnly = true;
            this.colField.Width = 59;
            // 
            // colValue
            // 
            this.colValue.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.colValue.HeaderText = "值";
            this.colValue.Name = "colValue";
            this.colValue.ReadOnly = true;
            // 
            // GraduateSurveyApproach
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.grdContent);
            this.Controls.Add(this.cmbSurveyYear);
            this.Name = "GraduateSurveyApproach";
            this.Size = new System.Drawing.Size(550, 495);
            this.Load += new System.EventHandler(this.BasicInfoItem_Load);
            ((System.ComponentModel.ISupportInitialize)(this.grdContent)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.Controls.ComboBoxEx cmbSurveyYear;
        private DevComponents.DotNetBar.Controls.DataGridViewX grdContent;
        private System.Windows.Forms.DataGridViewTextBoxColumn colField;
        private System.Windows.Forms.DataGridViewTextBoxColumn colValue;

    }
}
