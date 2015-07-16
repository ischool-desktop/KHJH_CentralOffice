namespace KHJHCentralOffice
{
    partial class BasicInfoItem
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
            this.label1 = new DevComponents.DotNetBar.LabelX();
            this.label2 = new DevComponents.DotNetBar.LabelX();
            this.label3 = new DevComponents.DotNetBar.LabelX();
            this.label4 = new DevComponents.DotNetBar.LabelX();
            this.txtComment = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.txtTitle = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.txtDSNS = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.txtPhysicalUrl = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.cmbGroup = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.comboItem1 = new DevComponents.Editors.ComboItem();
            this.comboItem2 = new DevComponents.Editors.ComboItem();
            this.comboItem3 = new DevComponents.Editors.ComboItem();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            // 
            // 
            // 
            this.label1.BackgroundStyle.Class = "";
            this.label1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.label1.Location = new System.Drawing.Point(14, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(34, 21);
            this.label1.TabIndex = 0;
            this.label1.Text = "名稱";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            // 
            // 
            // 
            this.label2.BackgroundStyle.Class = "";
            this.label2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.label2.Location = new System.Drawing.Point(14, 108);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 21);
            this.label2.TabIndex = 2;
            this.label2.Text = "DSNS";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            // 
            // 
            // 
            this.label3.BackgroundStyle.Class = "";
            this.label3.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.label3.Location = new System.Drawing.Point(14, 61);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(34, 21);
            this.label3.TabIndex = 4;
            this.label3.Text = "分類";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            // 
            // 
            // 
            this.label4.BackgroundStyle.Class = "";
            this.label4.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.label4.Location = new System.Drawing.Point(268, 19);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(34, 21);
            this.label4.TabIndex = 6;
            this.label4.Text = "備註";
            // 
            // txtComment
            // 
            // 
            // 
            // 
            this.txtComment.Border.Class = "TextBoxBorder";
            this.txtComment.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.txtComment.Location = new System.Drawing.Point(314, 17);
            this.txtComment.Multiline = true;
            this.txtComment.Name = "txtComment";
            this.txtComment.Size = new System.Drawing.Size(223, 65);
            this.txtComment.TabIndex = 7;
            // 
            // txtTitle
            // 
            // 
            // 
            // 
            this.txtTitle.Border.Class = "TextBoxBorder";
            this.txtTitle.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.txtTitle.Location = new System.Drawing.Point(65, 17);
            this.txtTitle.Name = "txtTitle";
            this.txtTitle.Size = new System.Drawing.Size(188, 25);
            this.txtTitle.TabIndex = 8;
            // 
            // txtDSNS
            // 
            // 
            // 
            // 
            this.txtDSNS.Border.Class = "TextBoxBorder";
            this.txtDSNS.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.txtDSNS.Location = new System.Drawing.Point(65, 106);
            this.txtDSNS.Name = "txtDSNS";
            this.txtDSNS.ReadOnly = true;
            this.txtDSNS.Size = new System.Drawing.Size(188, 25);
            this.txtDSNS.TabIndex = 8;
            // 
            // labelX1
            // 
            this.labelX1.AutoSize = true;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(14, 139);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(34, 21);
            this.labelX1.TabIndex = 4;
            this.labelX1.Text = "位置";
            // 
            // txtPhysicalUrl
            // 
            // 
            // 
            // 
            this.txtPhysicalUrl.Border.Class = "TextBoxBorder";
            this.txtPhysicalUrl.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.txtPhysicalUrl.Location = new System.Drawing.Point(65, 137);
            this.txtPhysicalUrl.Name = "txtPhysicalUrl";
            this.txtPhysicalUrl.ReadOnly = true;
            this.txtPhysicalUrl.Size = new System.Drawing.Size(472, 25);
            this.txtPhysicalUrl.TabIndex = 8;
            // 
            // cmbGroup
            // 
            this.cmbGroup.DisplayMember = "Text";
            this.cmbGroup.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cmbGroup.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbGroup.FormattingEnabled = true;
            this.cmbGroup.ItemHeight = 19;
            this.cmbGroup.Items.AddRange(new object[] {
            this.comboItem1,
            this.comboItem2,
            this.comboItem3});
            this.cmbGroup.Location = new System.Drawing.Point(65, 57);
            this.cmbGroup.Name = "cmbGroup";
            this.cmbGroup.Size = new System.Drawing.Size(188, 25);
            this.cmbGroup.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cmbGroup.TabIndex = 9;
            // 
            // comboItem1
            // 
            this.comboItem1.Text = "公立國中";
            // 
            // comboItem2
            // 
            this.comboItem2.Text = "私立國中";
            // 
            // comboItem3
            // 
            this.comboItem3.Text = "附設國中部";
            // 
            // BasicInfoItem
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.cmbGroup);
            this.Controls.Add(this.txtPhysicalUrl);
            this.Controls.Add(this.txtDSNS);
            this.Controls.Add(this.txtTitle);
            this.Controls.Add(this.txtComment);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "BasicInfoItem";
            this.Size = new System.Drawing.Size(550, 100);
            this.Load += new System.EventHandler(this.BasicInfoItem_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.LabelX label1;
        private DevComponents.DotNetBar.LabelX label2;
        private DevComponents.DotNetBar.LabelX label3;
        private DevComponents.DotNetBar.LabelX label4;
        private DevComponents.DotNetBar.Controls.TextBoxX txtComment;
        private DevComponents.DotNetBar.Controls.TextBoxX txtTitle;
        private DevComponents.DotNetBar.Controls.TextBoxX txtDSNS;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.Controls.TextBoxX txtPhysicalUrl;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cmbGroup;
        private DevComponents.Editors.ComboItem comboItem1;
        private DevComponents.Editors.ComboItem comboItem2;
        private DevComponents.Editors.ComboItem comboItem3;
    }
}
