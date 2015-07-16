namespace KHJHCentralOffice
{
    partial class UDMManagerForm 
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.baseXmlSyntaxLanguage1 = new DesktopLib.BaseXmlSyntaxLanguage();
            this.dgvModules = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.btnDelete = new DevComponents.DotNetBar.ButtonX();
            this.btnAdd = new DevComponents.DotNetBar.ButtonX();
            this.chName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chUrl = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chUDTContains = new DevComponents.DotNetBar.Controls.DataGridViewCheckBoxXColumn();
            this.chRefCount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnUpdate = new DevComponents.DotNetBar.ButtonX();
            ((System.ComponentModel.ISupportInitialize)(this.dgvModules)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvModules
            // 
            this.dgvModules.AllowUserToAddRows = false;
            this.dgvModules.AllowUserToDeleteRows = false;
            this.dgvModules.AllowUserToOrderColumns = true;
            this.dgvModules.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvModules.BackgroundColor = System.Drawing.Color.White;
            this.dgvModules.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvModules.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.chName,
            this.chUrl,
            this.chUDTContains,
            this.chRefCount});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvModules.DefaultCellStyle = dataGridViewCellStyle2;
            this.dgvModules.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgvModules.Location = new System.Drawing.Point(12, 12);
            this.dgvModules.Name = "dgvModules";
            this.dgvModules.ReadOnly = true;
            this.dgvModules.RowTemplate.Height = 24;
            this.dgvModules.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvModules.Size = new System.Drawing.Size(758, 386);
            this.dgvModules.TabIndex = 0;
            // 
            // btnDelete
            // 
            this.btnDelete.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnDelete.BackColor = System.Drawing.Color.Transparent;
            this.btnDelete.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnDelete.Location = new System.Drawing.Point(695, 410);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(75, 23);
            this.btnDelete.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnDelete.TabIndex = 1;
            this.btnDelete.Text = "刪除模組";
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnAdd.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnAdd.BackColor = System.Drawing.Color.Transparent;
            this.btnAdd.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnAdd.Location = new System.Drawing.Point(614, 410);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnAdd.TabIndex = 1;
            this.btnAdd.Text = "新增模組";
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // chName
            // 
            this.chName.DataPropertyName = "Name";
            this.chName.HeaderText = "Name";
            this.chName.Name = "chName";
            this.chName.ReadOnly = true;
            this.chName.Width = 150;
            // 
            // chUrl
            // 
            this.chUrl.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.chUrl.DataPropertyName = "Url";
            this.chUrl.HeaderText = "Url";
            this.chUrl.Name = "chUrl";
            this.chUrl.ReadOnly = true;
            // 
            // chUDTContains
            // 
            this.chUDTContains.Checked = true;
            this.chUDTContains.CheckState = System.Windows.Forms.CheckState.Indeterminate;
            this.chUDTContains.CheckValue = "N";
            this.chUDTContains.DataPropertyName = "UDTContains";
            this.chUDTContains.HeaderText = "UDT";
            this.chUDTContains.Name = "chUDTContains";
            this.chUDTContains.ReadOnly = true;
            this.chUDTContains.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.chUDTContains.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.chUDTContains.Width = 60;
            // 
            // chRefCount
            // 
            this.chRefCount.DataPropertyName = "ReferenceCount";
            this.chRefCount.HeaderText = "Count";
            this.chRefCount.Name = "chRefCount";
            this.chRefCount.ReadOnly = true;
            this.chRefCount.Width = 80;
            // 
            // btnUpdate
            // 
            this.btnUpdate.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnUpdate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnUpdate.BackColor = System.Drawing.Color.Transparent;
            this.btnUpdate.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnUpdate.Location = new System.Drawing.Point(12, 410);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(75, 23);
            this.btnUpdate.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnUpdate.TabIndex = 3;
            this.btnUpdate.Text = "更新模組";
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // UDMManagerForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(782, 439);
            this.Controls.Add(this.btnUpdate);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.dgvModules);
            this.DoubleBuffered = true;
            this.Name = "UDMManagerForm";
            this.Text = "UDM 模組管理";
            this.Load += new System.EventHandler(this.DesktopModuleManagerForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvModules)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DesktopLib.BaseXmlSyntaxLanguage baseXmlSyntaxLanguage1;
        private DevComponents.DotNetBar.Controls.DataGridViewX dgvModules;
        private DevComponents.DotNetBar.ButtonX btnDelete;
        private DevComponents.DotNetBar.ButtonX btnAdd;
        private System.Windows.Forms.DataGridViewTextBoxColumn chName;
        private System.Windows.Forms.DataGridViewTextBoxColumn chUrl;
        private DevComponents.DotNetBar.Controls.DataGridViewCheckBoxXColumn chUDTContains;
        private System.Windows.Forms.DataGridViewTextBoxColumn chRefCount;
        private DevComponents.DotNetBar.ButtonX btnUpdate;
    }
}