namespace KHJHCentralOffice
{
    partial class SearchForm
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
            this.btnSearch = new DevComponents.DotNetBar.ButtonX();
            this.btnPlugin = new DevComponents.DotNetBar.ButtonItem();
            this.btnGadget = new DevComponents.DotNetBar.ButtonItem();
            this.btnUDM = new DevComponents.DotNetBar.ButtonItem();
            this.txtPattern = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.SuspendLayout();
            // 
            // btnSearch
            // 
            this.btnSearch.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSearch.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSearch.AutoExpandOnClick = true;
            this.btnSearch.BackColor = System.Drawing.Color.Transparent;
            this.btnSearch.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSearch.Location = new System.Drawing.Point(449, 60);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(65, 23);
            this.btnSearch.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnSearch.SubItems.AddRange(new DevComponents.DotNetBar.BaseItem[] {
            this.btnPlugin,
            this.btnGadget,
            this.btnUDM});
            this.btnSearch.TabIndex = 2;
            this.btnSearch.Text = "搜尋";
            // 
            // btnPlugin
            // 
            this.btnPlugin.GlobalItem = false;
            this.btnPlugin.Name = "btnPlugin";
            this.btnPlugin.Text = "Desktop Plugin";
            this.btnPlugin.Tooltip = "請使用類似：JHCore/app.xml 搜尋。";
            this.btnPlugin.Click += new System.EventHandler(this.btnPlugin_Click);
            // 
            // btnGadget
            // 
            this.btnGadget.GlobalItem = false;
            this.btnGadget.Name = "btnGadget";
            this.btnGadget.Text = "Web Gadget";
            this.btnGadget.Tooltip = "請使用 Gadget 的 ID，例：934e39d3-893f-4338-96d6-fd486497b930";
            this.btnGadget.Click += new System.EventHandler(this.btnGadget_Click);
            // 
            // btnUDM
            // 
            this.btnUDM.GlobalItem = false;
            this.btnUDM.Name = "btnUDM";
            this.btnUDM.Text = "UDM";
            this.btnUDM.Tooltip = "請使用 URL 的一部份，例：137815/AttendanceDiscipline/udm.xml";
            this.btnUDM.Click += new System.EventHandler(this.btnUDM_Click);
            // 
            // txtPattern
            // 
            this.txtPattern.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            // 
            // 
            // 
            this.txtPattern.Border.Class = "TextBoxBorder";
            this.txtPattern.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.txtPattern.Location = new System.Drawing.Point(12, 12);
            this.txtPattern.Name = "txtPattern";
            this.txtPattern.Size = new System.Drawing.Size(502, 25);
            this.txtPattern.TabIndex = 0;
            this.txtPattern.WatermarkText = "關鍵字，例：JHCore/app.xml";
            // 
            // labelX1
            // 
            this.labelX1.AutoSize = true;
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(12, 61);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(311, 21);
            this.labelX1.TabIndex = 1;
            this.labelX1.Text = "※ 搜尋後的結果會加入待處理(只會搜尋選擇的學校)";
            // 
            // SearchForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(526, 94);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.txtPattern);
            this.Controls.Add(this.btnSearch);
            this.DoubleBuffered = true;
            this.Name = "SearchForm";
            this.Text = "進階搜尋";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.ButtonX btnSearch;
        private DevComponents.DotNetBar.ButtonItem btnPlugin;
        private DevComponents.DotNetBar.ButtonItem btnUDM;
        private DevComponents.DotNetBar.Controls.TextBoxX txtPattern;
        private DevComponents.DotNetBar.ButtonItem btnGadget;
        private DevComponents.DotNetBar.LabelX labelX1;
    }
}