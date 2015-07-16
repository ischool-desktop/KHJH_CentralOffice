namespace KHJHCentralOffice
{
    partial class UDTItem
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
            this.btnInstall = new DevComponents.DotNetBar.ButtonX();
            this.dgvUDT = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.chName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgvUDT)).BeginInit();
            this.SuspendLayout();
            // 
            // btnInstall
            // 
            this.btnInstall.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnInstall.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnInstall.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnInstall.Location = new System.Drawing.Point(462, 278);
            this.btnInstall.Name = "btnInstall";
            this.btnInstall.Size = new System.Drawing.Size(75, 23);
            this.btnInstall.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnInstall.TabIndex = 4;
            this.btnInstall.Text = "安裝/更新";
            this.btnInstall.Click += new System.EventHandler(this.btnInstall_Click);
            // 
            // dgvUDT
            // 
            this.dgvUDT.AllowUserToAddRows = false;
            this.dgvUDT.AllowUserToDeleteRows = false;
            this.dgvUDT.AllowUserToResizeRows = false;
            this.dgvUDT.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.dgvUDT.BackgroundColor = System.Drawing.Color.White;
            this.dgvUDT.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvUDT.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.chName});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvUDT.DefaultCellStyle = dataGridViewCellStyle1;
            this.dgvUDT.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgvUDT.Location = new System.Drawing.Point(13, 8);
            this.dgvUDT.Name = "dgvUDT";
            this.dgvUDT.ReadOnly = true;
            this.dgvUDT.RowHeadersVisible = false;
            this.dgvUDT.RowHeadersWidth = 25;
            this.dgvUDT.RowTemplate.Height = 24;
            this.dgvUDT.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvUDT.Size = new System.Drawing.Size(524, 264);
            this.dgvUDT.TabIndex = 3;
            this.dgvUDT.CellMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dgvUDT_CellMouseDoubleClick);
            // 
            // chName
            // 
            this.chName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.chName.HeaderText = "名稱";
            this.chName.Name = "chName";
            this.chName.ReadOnly = true;
            // 
            // UDTItem
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnInstall);
            this.Controls.Add(this.dgvUDT);
            this.Name = "UDTItem";
            this.Size = new System.Drawing.Size(550, 310);
            this.Load += new System.EventHandler(this.UDTItem_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvUDT)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.ButtonX btnInstall;
        private DevComponents.DotNetBar.Controls.DataGridViewX dgvUDT;
        private System.Windows.Forms.DataGridViewTextBoxColumn chName;
    }
}
