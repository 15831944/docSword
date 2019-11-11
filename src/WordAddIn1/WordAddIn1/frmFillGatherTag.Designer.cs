namespace OfficeAssist
{
    partial class frmFillGatherTag
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
            this.btnTagDialogOK = new System.Windows.Forms.Button();
            this.btnTagDialogCancel = new System.Windows.Forms.Button();
            this.rdBtnTagFullPath = new System.Windows.Forms.RadioButton();
            this.txtTagDialogSelfFill = new System.Windows.Forms.TextBox();
            this.rdBtnShortFileName = new System.Windows.Forms.RadioButton();
            this.rdBtnOnlyDirectory = new System.Windows.Forms.RadioButton();
            this.rdBtnSelfFill = new System.Windows.Forms.RadioButton();
            this.rdBtnTblSn = new System.Windows.Forms.RadioButton();
            this.rdBtnTagAbsPageNum = new System.Windows.Forms.RadioButton();
            this.rdBtnFileShortNameNoExt = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // btnTagDialogOK
            // 
            this.btnTagDialogOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnTagDialogOK.Location = new System.Drawing.Point(44, 178);
            this.btnTagDialogOK.Name = "btnTagDialogOK";
            this.btnTagDialogOK.Size = new System.Drawing.Size(75, 23);
            this.btnTagDialogOK.TabIndex = 0;
            this.btnTagDialogOK.Text = "确定";
            this.btnTagDialogOK.UseVisualStyleBackColor = true;
            this.btnTagDialogOK.Click += new System.EventHandler(this.btnTagDialogOK_Click);
            // 
            // btnTagDialogCancel
            // 
            this.btnTagDialogCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnTagDialogCancel.Location = new System.Drawing.Point(206, 178);
            this.btnTagDialogCancel.Name = "btnTagDialogCancel";
            this.btnTagDialogCancel.Size = new System.Drawing.Size(75, 23);
            this.btnTagDialogCancel.TabIndex = 1;
            this.btnTagDialogCancel.Text = "取消";
            this.btnTagDialogCancel.UseVisualStyleBackColor = true;
            // 
            // rdBtnTagFullPath
            // 
            this.rdBtnTagFullPath.AutoSize = true;
            this.rdBtnTagFullPath.Checked = true;
            this.rdBtnTagFullPath.Location = new System.Drawing.Point(43, 9);
            this.rdBtnTagFullPath.Name = "rdBtnTagFullPath";
            this.rdBtnTagFullPath.Size = new System.Drawing.Size(131, 16);
            this.rdBtnTagFullPath.TabIndex = 2;
            this.rdBtnTagFullPath.TabStop = true;
            this.rdBtnTagFullPath.Text = "文件全名（带目录）";
            this.rdBtnTagFullPath.UseVisualStyleBackColor = true;
            // 
            // txtTagDialogSelfFill
            // 
            this.txtTagDialogSelfFill.Location = new System.Drawing.Point(95, 135);
            this.txtTagDialogSelfFill.Name = "txtTagDialogSelfFill";
            this.txtTagDialogSelfFill.Size = new System.Drawing.Size(186, 21);
            this.txtTagDialogSelfFill.TabIndex = 3;
            // 
            // rdBtnShortFileName
            // 
            this.rdBtnShortFileName.AutoSize = true;
            this.rdBtnShortFileName.Location = new System.Drawing.Point(43, 30);
            this.rdBtnShortFileName.Name = "rdBtnShortFileName";
            this.rdBtnShortFileName.Size = new System.Drawing.Size(119, 16);
            this.rdBtnShortFileName.TabIndex = 4;
            this.rdBtnShortFileName.Text = "文件名（无目录）";
            this.rdBtnShortFileName.UseVisualStyleBackColor = true;
            // 
            // rdBtnOnlyDirectory
            // 
            this.rdBtnOnlyDirectory.AutoSize = true;
            this.rdBtnOnlyDirectory.Location = new System.Drawing.Point(43, 72);
            this.rdBtnOnlyDirectory.Name = "rdBtnOnlyDirectory";
            this.rdBtnOnlyDirectory.Size = new System.Drawing.Size(59, 16);
            this.rdBtnOnlyDirectory.TabIndex = 5;
            this.rdBtnOnlyDirectory.Text = "仅目录";
            this.rdBtnOnlyDirectory.UseVisualStyleBackColor = true;
            // 
            // rdBtnSelfFill
            // 
            this.rdBtnSelfFill.AutoSize = true;
            this.rdBtnSelfFill.Location = new System.Drawing.Point(43, 135);
            this.rdBtnSelfFill.Name = "rdBtnSelfFill";
            this.rdBtnSelfFill.Size = new System.Drawing.Size(47, 16);
            this.rdBtnSelfFill.TabIndex = 6;
            this.rdBtnSelfFill.Text = "自填";
            this.rdBtnSelfFill.UseVisualStyleBackColor = true;
            // 
            // rdBtnTblSn
            // 
            this.rdBtnTblSn.AutoSize = true;
            this.rdBtnTblSn.Location = new System.Drawing.Point(43, 93);
            this.rdBtnTblSn.Name = "rdBtnTblSn";
            this.rdBtnTblSn.Size = new System.Drawing.Size(59, 16);
            this.rdBtnTblSn.TabIndex = 7;
            this.rdBtnTblSn.TabStop = true;
            this.rdBtnTblSn.Text = "表序号";
            this.rdBtnTblSn.UseVisualStyleBackColor = true;
            // 
            // rdBtnTagAbsPageNum
            // 
            this.rdBtnTagAbsPageNum.AutoSize = true;
            this.rdBtnTagAbsPageNum.Location = new System.Drawing.Point(43, 114);
            this.rdBtnTagAbsPageNum.Name = "rdBtnTagAbsPageNum";
            this.rdBtnTagAbsPageNum.Size = new System.Drawing.Size(143, 16);
            this.rdBtnTagAbsPageNum.TabIndex = 8;
            this.rdBtnTagAbsPageNum.TabStop = true;
            this.rdBtnTagAbsPageNum.Text = "所在页码（绝对页码）";
            this.rdBtnTagAbsPageNum.UseVisualStyleBackColor = true;
            // 
            // rdBtnFileShortNameNoExt
            // 
            this.rdBtnFileShortNameNoExt.AutoSize = true;
            this.rdBtnFileShortNameNoExt.Location = new System.Drawing.Point(43, 51);
            this.rdBtnFileShortNameNoExt.Name = "rdBtnFileShortNameNoExt";
            this.rdBtnFileShortNameNoExt.Size = new System.Drawing.Size(167, 16);
            this.rdBtnFileShortNameNoExt.TabIndex = 9;
            this.rdBtnFileShortNameNoExt.TabStop = true;
            this.rdBtnFileShortNameNoExt.Text = "文件名（无后缀，无目录）";
            this.rdBtnFileShortNameNoExt.UseVisualStyleBackColor = true;
            // 
            // frmFillGatherTag
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(320, 210);
            this.Controls.Add(this.rdBtnFileShortNameNoExt);
            this.Controls.Add(this.rdBtnTagAbsPageNum);
            this.Controls.Add(this.rdBtnTblSn);
            this.Controls.Add(this.rdBtnSelfFill);
            this.Controls.Add(this.rdBtnOnlyDirectory);
            this.Controls.Add(this.rdBtnShortFileName);
            this.Controls.Add(this.txtTagDialogSelfFill);
            this.Controls.Add(this.rdBtnTagFullPath);
            this.Controls.Add(this.btnTagDialogCancel);
            this.Controls.Add(this.btnTagDialogOK);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmFillGatherTag";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmFillGatherTag";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnTagDialogOK;
        private System.Windows.Forms.Button btnTagDialogCancel;
        private System.Windows.Forms.RadioButton rdBtnTagFullPath;
        private System.Windows.Forms.TextBox txtTagDialogSelfFill;
        private System.Windows.Forms.RadioButton rdBtnShortFileName;
        private System.Windows.Forms.RadioButton rdBtnOnlyDirectory;
        private System.Windows.Forms.RadioButton rdBtnSelfFill;
        private System.Windows.Forms.RadioButton rdBtnTblSn;
        private System.Windows.Forms.RadioButton rdBtnTagAbsPageNum;
        private System.Windows.Forms.RadioButton rdBtnFileShortNameNoExt;
    }
}