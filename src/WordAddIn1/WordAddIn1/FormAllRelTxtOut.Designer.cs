namespace OfficeAssist
{
    partial class FormAllRelTxtOut
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
            this.txtAllRelsOut = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // txtAllRelsOut
            // 
            this.txtAllRelsOut.Location = new System.Drawing.Point(-1, 2);
            this.txtAllRelsOut.Margin = new System.Windows.Forms.Padding(2);
            this.txtAllRelsOut.Multiline = true;
            this.txtAllRelsOut.Name = "txtAllRelsOut";
            this.txtAllRelsOut.ReadOnly = true;
            this.txtAllRelsOut.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtAllRelsOut.Size = new System.Drawing.Size(394, 389);
            this.txtAllRelsOut.TabIndex = 0;
            this.txtAllRelsOut.Text = "类别：名称：公式：内容：说明\r\n第一行\r\n第二行\r\n";
            // 
            // FormAllRelTxtOut
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(394, 393);
            this.Controls.Add(this.txtAllRelsOut);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.MinimizeBox = false;
            this.Name = "FormAllRelTxtOut";
            this.Text = "关系全文";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtAllRelsOut;
    }
}