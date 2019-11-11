namespace OfficeAssist
{
    partial class frmRegister
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
            this.label1 = new System.Windows.Forms.Label();
            this.txtRegisterAccount = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtActivateSn = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtRegisterInfo = new System.Windows.Forms.TextBox();
            this.btnRegisterStart = new System.Windows.Forms.Button();
            this.btnRegisterClose = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(21, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "账号";
            // 
            // txtRegisterAccount
            // 
            this.txtRegisterAccount.Location = new System.Drawing.Point(83, 32);
            this.txtRegisterAccount.Name = "txtRegisterAccount";
            this.txtRegisterAccount.Size = new System.Drawing.Size(272, 21);
            this.txtRegisterAccount.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(20, 86);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "激活码";
            // 
            // txtActivateSn
            // 
            this.txtActivateSn.Location = new System.Drawing.Point(83, 83);
            this.txtActivateSn.Name = "txtActivateSn";
            this.txtActivateSn.Size = new System.Drawing.Size(272, 21);
            this.txtActivateSn.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(21, 138);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(29, 12);
            this.label3.TabIndex = 4;
            this.label3.Text = "信息";
            // 
            // txtRegisterInfo
            // 
            this.txtRegisterInfo.Enabled = false;
            this.txtRegisterInfo.Location = new System.Drawing.Point(82, 137);
            this.txtRegisterInfo.Multiline = true;
            this.txtRegisterInfo.Name = "txtRegisterInfo";
            this.txtRegisterInfo.ReadOnly = true;
            this.txtRegisterInfo.Size = new System.Drawing.Size(273, 66);
            this.txtRegisterInfo.TabIndex = 5;
            // 
            // btnRegisterStart
            // 
            this.btnRegisterStart.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnRegisterStart.Location = new System.Drawing.Point(82, 225);
            this.btnRegisterStart.Name = "btnRegisterStart";
            this.btnRegisterStart.Size = new System.Drawing.Size(75, 23);
            this.btnRegisterStart.TabIndex = 6;
            this.btnRegisterStart.Text = "激活";
            this.btnRegisterStart.UseVisualStyleBackColor = true;
            this.btnRegisterStart.Click += new System.EventHandler(this.btnRegisterStart_Click);
            // 
            // btnRegisterClose
            // 
            this.btnRegisterClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnRegisterClose.Location = new System.Drawing.Point(229, 225);
            this.btnRegisterClose.Name = "btnRegisterClose";
            this.btnRegisterClose.Size = new System.Drawing.Size(75, 23);
            this.btnRegisterClose.TabIndex = 7;
            this.btnRegisterClose.Text = "关闭";
            this.btnRegisterClose.UseVisualStyleBackColor = true;
            // 
            // frmRegister
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(388, 283);
            this.Controls.Add(this.btnRegisterClose);
            this.Controls.Add(this.btnRegisterStart);
            this.Controls.Add(this.txtRegisterInfo);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtActivateSn);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtRegisterAccount);
            this.Controls.Add(this.label1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmRegister";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "注册";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnRegisterStart;
        private System.Windows.Forms.Button btnRegisterClose;
        public System.Windows.Forms.TextBox txtRegisterAccount;
        public System.Windows.Forms.TextBox txtActivateSn;
        public System.Windows.Forms.TextBox txtRegisterInfo;
    }
}