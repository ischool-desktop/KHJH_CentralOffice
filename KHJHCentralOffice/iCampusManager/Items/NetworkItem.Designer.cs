namespace KHJHCentralOffice
{
    partial class NetworkItem
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
            this.DisplayArea = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.DisplayArea)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.DisplayArea.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DisplayArea.Location = new System.Drawing.Point(10, 5);
            this.DisplayArea.Name = "pictureBox1";
            this.DisplayArea.Size = new System.Drawing.Size(530, 295);
            this.DisplayArea.TabIndex = 0;
            this.DisplayArea.TabStop = false;
            // 
            // NetworkItem
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.DisplayArea);
            this.Name = "NetworkItem";
            this.Size = new System.Drawing.Size(550, 305);
            this.Load += new System.EventHandler(this.NetworkItem_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DisplayArea)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox DisplayArea;


    }
}
