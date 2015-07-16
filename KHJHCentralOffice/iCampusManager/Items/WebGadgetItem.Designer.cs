namespace KHJHCentralOffice
{
    partial class WebGadgetItem
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
            this.tabs = new DevComponents.DotNetBar.SuperTabStrip();
            this.tabGuest = new DevComponents.DotNetBar.SuperTabItem();
            this.tabTeacher = new DevComponents.DotNetBar.SuperTabItem();
            this.tabStudent = new DevComponents.DotNetBar.SuperTabItem();
            this.tabParent = new DevComponents.DotNetBar.SuperTabItem();
            this.tabAdmin = new DevComponents.DotNetBar.SuperTabItem();
            this.dgvGadgets = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.dataGridViewTextBoxColumn9 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn10 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.menus = new DevComponents.DotNetBar.ContextMenuBar();
            this.btnGridCtxMenu = new DevComponents.DotNetBar.ButtonItem();
            this.btnAddGadget = new DevComponents.DotNetBar.ButtonItem();
            this.btnDeleteGadget = new DevComponents.DotNetBar.ButtonItem();
            this.btnAdjustOrder = new DevComponents.DotNetBar.ButtonItem();
            ((System.ComponentModel.ISupportInitialize)(this.tabs)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvGadgets)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.menus)).BeginInit();
            this.SuspendLayout();
            // 
            // tabs
            // 
            this.tabs.AutoSelectAttachedControl = false;
            // 
            // 
            // 
            this.tabs.BackgroundStyle.Class = "";
            this.tabs.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.tabs.CloseButtonPosition = DevComponents.DotNetBar.eTabCloseButtonPosition.Left;
            this.tabs.ContainerControlProcessDialogKey = true;
            // 
            // 
            // 
            // 
            // 
            // 
            this.tabs.ControlBox.CloseBox.Name = "";
            // 
            // 
            // 
            this.tabs.ControlBox.MenuBox.Name = "";
            this.tabs.ControlBox.Name = "";
            this.tabs.ControlBox.SubItems.AddRange(new DevComponents.DotNetBar.BaseItem[] {
            this.tabs.ControlBox.MenuBox,
            this.tabs.ControlBox.CloseBox});
            this.tabs.ControlBox.Visible = false;
            this.tabs.Dock = System.Windows.Forms.DockStyle.Top;
            this.tabs.Location = new System.Drawing.Point(10, 5);
            this.tabs.Name = "tabs";
            this.tabs.ReorderTabsEnabled = true;
            this.tabs.SelectedTabFont = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Bold);
            this.tabs.SelectedTabIndex = 0;
            this.tabs.Size = new System.Drawing.Size(530, 31);
            this.tabs.TabCloseButtonHot = null;
            this.tabs.TabFont = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.tabs.TabIndex = 1;
            this.tabs.Tabs.AddRange(new DevComponents.DotNetBar.BaseItem[] {
            this.tabGuest,
            this.tabTeacher,
            this.tabStudent,
            this.tabParent,
            this.tabAdmin});
            this.tabs.Text = "superTabStrip1";
            this.tabs.SelectedTabChanged += new System.EventHandler<DevComponents.DotNetBar.SuperTabStripSelectedTabChangedEventArgs>(this.tabs_SelectedTabChanged);
            // 
            // tabGuest
            // 
            this.tabGuest.GlobalItem = false;
            this.tabGuest.Name = "tabGuest";
            this.tabGuest.Tag = "Guest";
            this.tabGuest.Text = "來賓/訪客";
            // 
            // tabTeacher
            // 
            this.tabTeacher.GlobalItem = false;
            this.tabTeacher.Name = "tabTeacher";
            this.tabTeacher.Tag = "Teacher";
            this.tabTeacher.Text = "老師(預設)";
            // 
            // tabStudent
            // 
            this.tabStudent.GlobalItem = false;
            this.tabStudent.Name = "tabStudent";
            this.tabStudent.Tag = "Student";
            this.tabStudent.Text = "學生(預設)";
            // 
            // tabParent
            // 
            this.tabParent.GlobalItem = false;
            this.tabParent.Name = "tabParent";
            this.tabParent.Tag = "Parent";
            this.tabParent.Text = "家長(預設)";
            // 
            // tabAdmin
            // 
            this.tabAdmin.GlobalItem = false;
            this.tabAdmin.Name = "tabAdmin";
            this.tabAdmin.Tag = "Admin";
            this.tabAdmin.Text = "系統管理員";
            // 
            // dgvGadgets
            // 
            this.dgvGadgets.AllowUserToAddRows = false;
            this.dgvGadgets.AllowUserToDeleteRows = false;
            this.dgvGadgets.AllowUserToResizeRows = false;
            this.dgvGadgets.BackgroundColor = System.Drawing.Color.White;
            this.dgvGadgets.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvGadgets.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn9,
            this.dataGridViewTextBoxColumn10});
            this.menus.SetContextMenuEx(this.dgvGadgets, this.btnGridCtxMenu);
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvGadgets.DefaultCellStyle = dataGridViewCellStyle1;
            this.dgvGadgets.Dock = System.Windows.Forms.DockStyle.Top;
            this.dgvGadgets.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgvGadgets.Location = new System.Drawing.Point(10, 36);
            this.dgvGadgets.Name = "dgvGadgets";
            this.dgvGadgets.ReadOnly = true;
            this.dgvGadgets.RowHeadersWidth = 25;
            this.dgvGadgets.RowTemplate.Height = 24;
            this.dgvGadgets.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvGadgets.Size = new System.Drawing.Size(530, 264);
            this.dgvGadgets.TabIndex = 6;
            this.dgvGadgets.CellMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dgvGadgets_CellMouseDoubleClick);
            // 
            // dataGridViewTextBoxColumn9
            // 
            this.dataGridViewTextBoxColumn9.DataPropertyName = "Title";
            this.dataGridViewTextBoxColumn9.HeaderText = "標題";
            this.dataGridViewTextBoxColumn9.Name = "dataGridViewTextBoxColumn9";
            this.dataGridViewTextBoxColumn9.ReadOnly = true;
            this.dataGridViewTextBoxColumn9.Width = 120;
            // 
            // dataGridViewTextBoxColumn10
            // 
            this.dataGridViewTextBoxColumn10.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.dataGridViewTextBoxColumn10.DataPropertyName = "Description";
            this.dataGridViewTextBoxColumn10.HeaderText = "說明";
            this.dataGridViewTextBoxColumn10.Name = "dataGridViewTextBoxColumn10";
            this.dataGridViewTextBoxColumn10.ReadOnly = true;
            // 
            // menus
            // 
            this.menus.AntiAlias = true;
            this.menus.Items.AddRange(new DevComponents.DotNetBar.BaseItem[] {
            this.btnGridCtxMenu});
            this.menus.Location = new System.Drawing.Point(422, 264);
            this.menus.Name = "menus";
            this.menus.Size = new System.Drawing.Size(110, 27);
            this.menus.Stretch = true;
            this.menus.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.menus.TabIndex = 0;
            this.menus.TabStop = false;
            this.menus.Text = "contextMenuBar1";
            // 
            // btnGridCtxMenu
            // 
            this.btnGridCtxMenu.AutoExpandOnClick = true;
            this.btnGridCtxMenu.Name = "btnGridCtxMenu";
            this.btnGridCtxMenu.SubItems.AddRange(new DevComponents.DotNetBar.BaseItem[] {
            this.btnAddGadget,
            this.btnDeleteGadget,
            this.btnAdjustOrder});
            this.btnGridCtxMenu.Text = "GridMenu";
            // 
            // btnAddGadget
            // 
            this.btnAddGadget.Name = "btnAddGadget";
            this.btnAddGadget.Text = "新增";
            this.btnAddGadget.Click += new System.EventHandler(this.btnAddGadget_Click);
            // 
            // btnDeleteGadget
            // 
            this.btnDeleteGadget.Name = "btnDeleteGadget";
            this.btnDeleteGadget.Text = "刪除";
            this.btnDeleteGadget.Click += new System.EventHandler(this.btnDeleteGadget_Click);
            // 
            // btnAdjustOrder
            // 
            this.btnAdjustOrder.BeginGroup = true;
            this.btnAdjustOrder.Name = "btnAdjustOrder";
            this.btnAdjustOrder.Text = "調整順序";
            this.btnAdjustOrder.Click += new System.EventHandler(this.btnAdjustOrder_Click);
            // 
            // WebGadgetItem
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.menus);
            this.Controls.Add(this.dgvGadgets);
            this.Controls.Add(this.tabs);
            this.Name = "WebGadgetItem";
            this.Size = new System.Drawing.Size(550, 310);
            this.Load += new System.EventHandler(this.WebGadgetItem_Load);
            ((System.ComponentModel.ISupportInitialize)(this.tabs)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvGadgets)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.menus)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.SuperTabStrip tabs;
        private DevComponents.DotNetBar.SuperTabItem tabStudent;
        private DevComponents.DotNetBar.SuperTabItem tabTeacher;
        private DevComponents.DotNetBar.SuperTabItem tabParent;
        private DevComponents.DotNetBar.SuperTabItem tabGuest;
        private DevComponents.DotNetBar.SuperTabItem tabAdmin;
        private DevComponents.DotNetBar.Controls.DataGridViewX dgvGadgets;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn9;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn10;
        private DevComponents.DotNetBar.ContextMenuBar menus;
        private DevComponents.DotNetBar.ButtonItem btnGridCtxMenu;
        private DevComponents.DotNetBar.ButtonItem btnAddGadget;
        private DevComponents.DotNetBar.ButtonItem btnDeleteGadget;
        private DevComponents.DotNetBar.ButtonItem btnAdjustOrder;

    }
}
