namespace OfficeAssist
{
    partial class FormPreview
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
            this.richTextBoxCnt = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // richTextBoxCnt
            // 
            this.richTextBoxCnt.Location = new System.Drawing.Point(23, 32);
            this.richTextBoxCnt.Name = "richTextBoxCnt";
            this.richTextBoxCnt.ReadOnly = true;
            this.richTextBoxCnt.Size = new System.Drawing.Size(536, 384);
            this.richTextBoxCnt.TabIndex = 0;
            this.richTextBoxCnt.Text = "";
            // 
            // FormPreview
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(604, 502);
            this.Controls.Add(this.richTextBoxCnt);
            this.MinimizeBox = false;
            this.Name = "FormPreview";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "预览";
            this.Load += new System.EventHandler(this.FormPreview_Load);
            this.ClientSizeChanged += new System.EventHandler(this.FormPreview_ClientSizeChanged);
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.RichTextBox richTextBoxCnt;

    }
}