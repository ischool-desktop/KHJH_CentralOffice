using DesktopLib;
using KHJHCentralOffice;

namespace KHJHCentralOffice
{
    partial class SimpleDefContent
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
            ActiproSoftware.SyntaxEditor.Document document1 = new ActiproSoftware.SyntaxEditor.Document();
            this.xmlLanguage = new BaseXmlSyntaxLanguage();
            this.editor = new BaseSyntaxEditor();
            this.SuspendLayout();
            // 
            // editor
            // 
            this.editor.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            document1.Language = this.xmlLanguage;
            document1.ReadOnly = true;
            this.editor.Document = document1;
            this.editor.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.editor.Location = new System.Drawing.Point(12, 12);
            this.editor.Name = "editor";
            this.editor.Size = new System.Drawing.Size(960, 438);
            this.editor.TabIndex = 0;
            // 
            // UDSContractDefContent
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(984, 462);
            this.Controls.Add(this.editor);
            this.DoubleBuffered = true;
            this.Name = "UDSContractDefContent";
            this.Text = "UDSContractDefContent";
            this.Load += new System.EventHandler(this.UDSContractDefContent_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private BaseSyntaxEditor editor;
        private BaseXmlSyntaxLanguage xmlLanguage;
    }
}