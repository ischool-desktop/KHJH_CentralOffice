namespace KHJHCentralOffice
{
    partial class UnApproach_Check
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.nudSchoolYear = new System.Windows.Forms.NumericUpDown();
            this.lblSchoolYear = new DevComponents.DotNetBar.LabelX();
            this.btnPrint = new DevComponents.DotNetBar.ButtonX();
            this.grdSchool = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.colSchoolName = new DevComponents.DotNetBar.Controls.DataGridViewLabelXColumn();
            ((System.ComponentModel.ISupportInitialize)(this.nudSchoolYear)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.grdSchool)).BeginInit();
            this.SuspendLayout();
            // 
            // nudSchoolYear
            // 
            this.nudSchoolYear.Location = new System.Drawing.Point(90, 12);
            this.nudSchoolYear.Maximum = new decimal(new int[] {
            999,
            0,
            0,
            0});
            this.nudSchoolYear.Name = "nudSchoolYear";
            this.nudSchoolYear.Size = new System.Drawing.Size(66, 25);
            this.nudSchoolYear.TabIndex = 38;
            this.nudSchoolYear.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // lblSchoolYear
            // 
            this.lblSchoolYear.AutoSize = true;
            this.lblSchoolYear.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblSchoolYear.BackgroundStyle.Class = "";
            this.lblSchoolYear.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblSchoolYear.Location = new System.Drawing.Point(13, 14);
            this.lblSchoolYear.Name = "lblSchoolYear";
            this.lblSchoolYear.Size = new System.Drawing.Size(74, 21);
            this.lblSchoolYear.TabIndex = 37;
            this.lblSchoolYear.Text = "填報學年度";
            // 
            // btnPrint
            // 
            this.btnPrint.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnPrint.AutoExpandOnClick = true;
            this.btnPrint.BackColor = System.Drawing.Color.Transparent;
            this.btnPrint.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnPrint.Location = new System.Drawing.Point(173, 12);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(80, 25);
            this.btnPrint.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnPrint.TabIndex = 83;
            this.btnPrint.Text = "檢  查";
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // grdSchool
            // 
            this.grdSchool.AllowUserToAddRows = false;
            this.grdSchool.AllowUserToDeleteRows = false;
            this.grdSchool.AllowUserToResizeColumns = false;
            this.grdSchool.AllowUserToResizeRows = false;
            this.grdSchool.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.grdSchool.BackgroundColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.grdSchool.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.grdSchool.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grdSchool.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colSchoolName});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.grdSchool.DefaultCellStyle = dataGridViewCellStyle2;
            this.grdSchool.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.grdSchool.Location = new System.Drawing.Point(12, 54);
            this.grdSchool.Name = "grdSchool";
            this.grdSchool.ReadOnly = true;
            this.grdSchool.RowHeadersVisible = false;
            this.grdSchool.RowTemplate.Height = 24;
            this.grdSchool.Size = new System.Drawing.Size(250, 337);
            this.grdSchool.TabIndex = 85;
            // 
            // colSchoolName
            // 
            this.colSchoolName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.colSchoolName.HeaderText = "學校名稱";
            this.colSchoolName.Name = "colSchoolName";
            this.colSchoolName.ReadOnly = true;
            this.colSchoolName.TextAlignment = System.Drawing.StringAlignment.Center;
            // 
            // UnApproach_Check
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(277, 403);
            this.Controls.Add(this.grdSchool);
            this.Controls.Add(this.btnPrint);
            this.Controls.Add(this.nudSchoolYear);
            this.Controls.Add(this.lblSchoolYear);
            this.MaximumSize = new System.Drawing.Size(285, 430);
            this.MinimumSize = new System.Drawing.Size(285, 430);
            this.Name = "UnApproach_Check";
            this.Text = "";
            this.TitleText = "未上傳國中檢查";
            ((System.ComponentModel.ISupportInitialize)(this.nudSchoolYear)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.grdSchool)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.NumericUpDown nudSchoolYear;
        public DevComponents.DotNetBar.LabelX lblSchoolYear;
        private DevComponents.DotNetBar.ButtonX btnPrint;
        private DevComponents.DotNetBar.Controls.DataGridViewX grdSchool;
        private DevComponents.DotNetBar.Controls.DataGridViewLabelXColumn colSchoolName;
    }
}