namespace OfficeAssist
{
    partial class FormLoadLicFile
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
            this.btnSelectLicFile = new System.Windows.Forms.Button();
            this.txtBoxSelectedLicFileLoc = new System.Windows.Forms.TextBox();
            this.txtBoxDecodeInfo = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnLoadInto = new System.Windows.Forms.Button();
            this.btnBackupCurLicFile = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnSelectLicFile
            // 
            this.btnSelectLicFile.Location = new System.Drawing.Point(13, 7);
            this.btnSelectLicFile.Name = "btnSelectLicFile";
            this.btnSelectLicFile.Size = new System.Drawing.Size(75, 25);
            this.btnSelectLicFile.TabIndex = 0;
            this.btnSelectLicFile.Text = "浏览";
            this.btnSelectLicFile.UseVisualStyleBackColor = true;
            this.btnSelectLicFile.Click += new System.EventHandler(this.btnSelectLicFile_Click);
            // 
            // txtBoxSelectedLicFileLoc
            // 
            this.txtBoxSelectedLicFileLoc.Location = new System.Drawing.Point(12, 38);
            this.txtBoxSelectedLicFileLoc.Multiline = true;
            this.txtBoxSelectedLicFileLoc.Name = "txtBoxSelectedLicFileLoc";
            this.txtBoxSelectedLicFileLoc.ReadOnly = true;
            this.txtBoxSelectedLicFileLoc.Size = new System.Drawing.Size(407, 62);
            this.txtBoxSelectedLicFileLoc.TabIndex = 1;
            // 
            // txtBoxDecodeInfo
            // 
            this.txtBoxDecodeInfo.Location = new System.Drawing.Point(12, 127);
            this.txtBoxDecodeInfo.Multiline = true;
            this.txtBoxDecodeInfo.Name = "txtBoxDecodeInfo";
            this.txtBoxDecodeInfo.ReadOnly = true;
            this.txtBoxDecodeInfo.Size = new System.Drawing.Size(407, 100);
            this.txtBoxDecodeInfo.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 106);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "信息";
            // 
            // btnLoadInto
            // 
            this.btnLoadInto.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnLoadInto.Location = new System.Drawing.Point(342, 235);
            this.btnLoadInto.Name = "btnLoadInto";
            this.btnLoadInto.Size = new System.Drawing.Size(75, 25);
            this.btnLoadInto.TabIndex = 0;
            this.btnLoadInto.Text = "导入";
            this.btnLoadInto.UseVisualStyleBackColor = true;
            this.btnLoadInto.Click += new System.EventHandler(this.btnLoadInto_Click);
            // 
            // btnBackupCurLicFile
            // 
            this.btnBackupCurLicFile.Location = new System.Drawing.Point(283, 7);
            this.btnBackupCurLicFile.Name = "btnBackupCurLicFile";
            this.btnBackupCurLicFile.Size = new System.Drawing.Size(134, 25);
            this.btnBackupCurLicFile.TabIndex = 0;
            this.btnBackupCurLicFile.Text = "备份当前单机版许可";
            this.btnBackupCurLicFile.UseVisualStyleBackColor = true;
            this.btnBackupCurLicFile.Click += new System.EventHandler(this.btnBackupCurLicFile_Click);
            // 
            // FormLoadLicFile
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(429, 266);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtBoxDecodeInfo);
            this.Controls.Add(this.txtBoxSelectedLicFileLoc);
            this.Controls.Add(this.btnLoadInto);
            this.Controls.Add(this.btnBackupCurLicFile);
            this.Controls.Add(this.btnSelectLicFile);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormLoadLicFile";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "加载单机版许可";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSelectLicFile;
        private System.Windows.Forms.TextBox txtBoxDecodeInfo;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnLoadInto;
        public System.Windows.Forms.TextBox txtBoxSelectedLicFileLoc;
        private System.Windows.Forms.Button btnBackupCurLicFile;
    }
}