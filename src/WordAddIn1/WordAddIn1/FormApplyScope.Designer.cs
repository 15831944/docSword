namespace OfficeAssist
{
    partial class FormApplyScope
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
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.chkIgnoreParaFormat = new System.Windows.Forms.CheckBox();
            this.chkIgnoreFont = new System.Windows.Forms.CheckBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtIgnorePages = new System.Windows.Forms.TextBox();
            this.chkIgnorePages = new System.Windows.Forms.CheckBox();
            this.chkIgnoreTable = new System.Windows.Forms.CheckBox();
            this.chkIgnoreTOC = new System.Windows.Forms.CheckBox();
            this.radioBtnStyleSelection = new System.Windows.Forms.RadioButton();
            this.radioBtnStyleAllDoc = new System.Windows.Forms.RadioButton();
            this.chkIgnoreTextBody = new System.Windows.Forms.CheckBox();
            this.chkIgnoreHeadings = new System.Windows.Forms.CheckBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.groupBox6.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.chkIgnoreParaFormat);
            this.groupBox6.Controls.Add(this.chkIgnoreFont);
            this.groupBox6.Controls.Add(this.radioBtnStyleSelection);
            this.groupBox6.Controls.Add(this.radioBtnStyleAllDoc);
            this.groupBox6.Location = new System.Drawing.Point(11, 11);
            this.groupBox6.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox6.Size = new System.Drawing.Size(291, 91);
            this.groupBox6.TabIndex = 8;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "应用范围";
            // 
            // chkIgnoreParaFormat
            // 
            this.chkIgnoreParaFormat.AutoSize = true;
            this.chkIgnoreParaFormat.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.chkIgnoreParaFormat.Location = new System.Drawing.Point(155, 51);
            this.chkIgnoreParaFormat.Name = "chkIgnoreParaFormat";
            this.chkIgnoreParaFormat.Size = new System.Drawing.Size(108, 16);
            this.chkIgnoreParaFormat.TabIndex = 12;
            this.chkIgnoreParaFormat.Text = "不改变段落格式";
            this.chkIgnoreParaFormat.UseVisualStyleBackColor = true;
            // 
            // chkIgnoreFont
            // 
            this.chkIgnoreFont.AutoSize = true;
            this.chkIgnoreFont.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.chkIgnoreFont.Location = new System.Drawing.Point(65, 51);
            this.chkIgnoreFont.Name = "chkIgnoreFont";
            this.chkIgnoreFont.Size = new System.Drawing.Size(84, 16);
            this.chkIgnoreFont.TabIndex = 11;
            this.chkIgnoreFont.Text = "不改变字体";
            this.chkIgnoreFont.UseVisualStyleBackColor = true;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.label7.Location = new System.Drawing.Point(285, 131);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(17, 12);
            this.label7.TabIndex = 6;
            this.label7.Text = "页";
            this.label7.Visible = false;
            // 
            // txtIgnorePages
            // 
            this.txtIgnorePages.Location = new System.Drawing.Point(255, 127);
            this.txtIgnorePages.Name = "txtIgnorePages";
            this.txtIgnorePages.Size = new System.Drawing.Size(28, 21);
            this.txtIgnorePages.TabIndex = 8;
            this.txtIgnorePages.Text = "1";
            this.txtIgnorePages.Visible = false;
            this.txtIgnorePages.Leave += new System.EventHandler(this.txtIgnorePages_Leave);
            // 
            // chkIgnorePages
            // 
            this.chkIgnorePages.AutoSize = true;
            this.chkIgnorePages.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.chkIgnorePages.Location = new System.Drawing.Point(189, 129);
            this.chkIgnorePages.Name = "chkIgnorePages";
            this.chkIgnorePages.Size = new System.Drawing.Size(60, 16);
            this.chkIgnorePages.TabIndex = 7;
            this.chkIgnorePages.Text = "忽略前";
            this.chkIgnorePages.UseVisualStyleBackColor = true;
            this.chkIgnorePages.Visible = false;
            this.chkIgnorePages.CheckedChanged += new System.EventHandler(this.chkIgnorePages_CheckedChanged);
            // 
            // chkIgnoreTable
            // 
            this.chkIgnoreTable.AutoSize = true;
            this.chkIgnoreTable.Checked = true;
            this.chkIgnoreTable.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkIgnoreTable.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.chkIgnoreTable.Location = new System.Drawing.Point(166, 6);
            this.chkIgnoreTable.Name = "chkIgnoreTable";
            this.chkIgnoreTable.Size = new System.Drawing.Size(72, 16);
            this.chkIgnoreTable.TabIndex = 6;
            this.chkIgnoreTable.Text = "忽略表格";
            this.chkIgnoreTable.UseVisualStyleBackColor = true;
            this.chkIgnoreTable.Visible = false;
            // 
            // chkIgnoreTOC
            // 
            this.chkIgnoreTOC.AutoSize = true;
            this.chkIgnoreTOC.Checked = true;
            this.chkIgnoreTOC.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkIgnoreTOC.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.chkIgnoreTOC.Location = new System.Drawing.Point(88, 6);
            this.chkIgnoreTOC.Name = "chkIgnoreTOC";
            this.chkIgnoreTOC.Size = new System.Drawing.Size(72, 16);
            this.chkIgnoreTOC.TabIndex = 5;
            this.chkIgnoreTOC.Text = "忽略目录";
            this.chkIgnoreTOC.UseVisualStyleBackColor = true;
            this.chkIgnoreTOC.Visible = false;
            // 
            // radioBtnStyleSelection
            // 
            this.radioBtnStyleSelection.AutoSize = true;
            this.radioBtnStyleSelection.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.radioBtnStyleSelection.Location = new System.Drawing.Point(155, 15);
            this.radioBtnStyleSelection.Margin = new System.Windows.Forms.Padding(2);
            this.radioBtnStyleSelection.Name = "radioBtnStyleSelection";
            this.radioBtnStyleSelection.Size = new System.Drawing.Size(71, 16);
            this.radioBtnStyleSelection.TabIndex = 4;
            this.radioBtnStyleSelection.Text = "选择部分";
            this.radioBtnStyleSelection.UseVisualStyleBackColor = true;
            // 
            // radioBtnStyleAllDoc
            // 
            this.radioBtnStyleAllDoc.AutoSize = true;
            this.radioBtnStyleAllDoc.Checked = true;
            this.radioBtnStyleAllDoc.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.radioBtnStyleAllDoc.Location = new System.Drawing.Point(66, 15);
            this.radioBtnStyleAllDoc.Margin = new System.Windows.Forms.Padding(2);
            this.radioBtnStyleAllDoc.Name = "radioBtnStyleAllDoc";
            this.radioBtnStyleAllDoc.Size = new System.Drawing.Size(47, 16);
            this.radioBtnStyleAllDoc.TabIndex = 3;
            this.radioBtnStyleAllDoc.TabStop = true;
            this.radioBtnStyleAllDoc.Text = "全文";
            this.radioBtnStyleAllDoc.UseVisualStyleBackColor = true;
            // 
            // chkIgnoreTextBody
            // 
            this.chkIgnoreTextBody.AutoSize = true;
            this.chkIgnoreTextBody.Checked = true;
            this.chkIgnoreTextBody.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkIgnoreTextBody.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.chkIgnoreTextBody.Location = new System.Drawing.Point(121, 129);
            this.chkIgnoreTextBody.Name = "chkIgnoreTextBody";
            this.chkIgnoreTextBody.Size = new System.Drawing.Size(72, 16);
            this.chkIgnoreTextBody.TabIndex = 10;
            this.chkIgnoreTextBody.Text = "忽略正文";
            this.chkIgnoreTextBody.UseVisualStyleBackColor = true;
            this.chkIgnoreTextBody.Visible = false;
            // 
            // chkIgnoreHeadings
            // 
            this.chkIgnoreHeadings.AutoSize = true;
            this.chkIgnoreHeadings.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.chkIgnoreHeadings.Location = new System.Drawing.Point(121, 107);
            this.chkIgnoreHeadings.Name = "chkIgnoreHeadings";
            this.chkIgnoreHeadings.Size = new System.Drawing.Size(72, 16);
            this.chkIgnoreHeadings.TabIndex = 9;
            this.chkIgnoreHeadings.Text = "忽略章节";
            this.chkIgnoreHeadings.UseVisualStyleBackColor = true;
            this.chkIgnoreHeadings.Visible = false;
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Location = new System.Drawing.Point(40, 103);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 9;
            this.btnOK.Text = "确定";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(208, 103);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 10;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // FormApplyScope
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(312, 153);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.chkIgnoreTextBody);
            this.Controls.Add(this.txtIgnorePages);
            this.Controls.Add(this.chkIgnoreHeadings);
            this.Controls.Add(this.chkIgnorePages);
            this.Controls.Add(this.groupBox6);
            this.Controls.Add(this.chkIgnoreTable);
            this.Controls.Add(this.chkIgnoreTOC);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormApplyScope";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "确定应用范围";
            this.Load += new System.EventHandler(this.FormApplyScope_Load);
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.CheckBox chkIgnoreParaFormat;
        private System.Windows.Forms.CheckBox chkIgnoreFont;
        private System.Windows.Forms.CheckBox chkIgnoreTextBody;
        private System.Windows.Forms.CheckBox chkIgnoreHeadings;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtIgnorePages;
        private System.Windows.Forms.CheckBox chkIgnorePages;
        private System.Windows.Forms.CheckBox chkIgnoreTable;
        private System.Windows.Forms.CheckBox chkIgnoreTOC;
        private System.Windows.Forms.RadioButton radioBtnStyleSelection;
        private System.Windows.Forms.RadioButton radioBtnStyleAllDoc;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
    }
}