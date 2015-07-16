namespace KHJHCentralOffice
{
    partial class ModuleItem
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
            this.btnDelete = new DevComponents.DotNetBar.ButtonX();
            this.btnInstall = new DevComponents.DotNetBar.ButtonX();
            this.dgvPlugin = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.chName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chUrl = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgvPlugin)).BeginInit();
            this.SuspendLayout();
            // 
            // btnDelete
            // 
            this.btnDelete.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnDelete.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnDelete.Location = new System.Drawing.Point(13, 279);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(75, 23);
            this.btnDelete.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnDelete.TabIndex = 5;
            this.btnDelete.Text = "刪除";
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnInstall
            // 
            this.btnInstall.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnInstall.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnInstall.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnInstall.Location = new System.Drawing.Point(462, 279);
            this.btnInstall.Name = "btnInstall";
            this.btnInstall.Size = new System.Drawing.Size(75, 23);
            this.btnInstall.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnInstall.TabIndex = 4;
            this.btnInstall.Text = "新增";
            this.btnInstall.Click += new System.EventHandler(this.btnInstall_Click);
            // 
            // dgvPlugin
            // 
            this.dgvPlugin.AllowUserToAddRows = false;
            this.dgvPlugin.AllowUserToDeleteRows = false;
            this.dgvPlugin.AllowUserToResizeRows = false;
            this.dgvPlugin.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.dgvPlugin.BackgroundColor = System.Drawing.Color.White;
            this.dgvPlugin.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvPlugin.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.chName,
            this.chUrl});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvPlugin.DefaultCellStyle = dataGridViewCellStyle1;
            this.dgvPlugin.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgvPlugin.Location = new System.Drawing.Point(13, 9);
            this.dgvPlugin.MultiSelect = false;
            this.dgvPlugin.Name = "dgvPlugin";
            this.dgvPlugin.ReadOnly = true;
            this.dgvPlugin.RowHeadersVisible = false;
            this.dgvPlugin.RowHeadersWidth = 25;
            this.dgvPlugin.RowTemplate.Height = 24;
            this.dgvPlugin.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvPlugin.Size = new System.Drawing.Size(524, 264);
            this.dgvPlugin.TabIndex = 3;
            this.dgvPlugin.CellContentDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvPlugin_CellContentDoubleClick);
            // 
            // chName
            // 
            this.chName.DataPropertyName = "Title";
            this.chName.HeaderText = "模組名稱";
            this.chName.Name = "chName";
            this.chName.ReadOnly = true;
            this.chName.Width = 200;
            // 
            // chUrl
            // 
            this.chUrl.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.chUrl.DataPropertyName = "UrlShort";
            this.chUrl.HeaderText = "位置";
            this.chUrl.Name = "chUrl";
            this.chUrl.ReadOnly = true;
            // 
            // ModuleItem
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnInstall);
            this.Controls.Add(this.dgvPlugin);
            this.Name = "ModuleItem";
            this.Size = new System.Drawing.Size(550, 310);
            this.Load += new System.EventHandler(this.ModuleItem_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvPlugin)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.ButtonX btnDelete;
        private DevComponents.DotNetBar.ButtonX btnInstall;
        private DevComponents.DotNetBar.Controls.DataGridViewX dgvPlugin;
        private System.Windows.Forms.DataGridViewTextBoxColumn chName;
        private System.Windows.Forms.DataGridViewTextBoxColumn chUrl;
    }
}
