namespace OfficeAssist
{
    partial class FormHeadingSnPos
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
            this.cmbHeadingSnAlign = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.numHeadingSnAlignPos = new System.Windows.Forms.NumericUpDown();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.numHeadingSnTextIndentPos = new System.Windows.Forms.NumericUpDown();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.cmbHeadingSnBehindSn = new System.Windows.Forms.ComboBox();
            this.chkHeadingSnTabPos = new System.Windows.Forms.CheckBox();
            this.label7 = new System.Windows.Forms.Label();
            this.numHeadingSnTabPos = new System.Windows.Forms.NumericUpDown();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.label8 = new System.Windows.Forms.Label();
            this.numStartAt = new System.Windows.Forms.NumericUpDown();
            this.cmbResetOnHigher = new System.Windows.Forms.ComboBox();
            this.chkResetOnHigher = new System.Windows.Forms.CheckBox();
            this.label9 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.numHeadingSnAlignPos)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numHeadingSnTextIndentPos)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numHeadingSnTabPos)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numStartAt)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(89, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "编号对齐方式：";
            // 
            // cmbHeadingSnAlign
            // 
            this.cmbHeadingSnAlign.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbHeadingSnAlign.FormattingEnabled = true;
            this.cmbHeadingSnAlign.Items.AddRange(new object[] {
            "左对齐",
            "居中",
            "右对齐"});
            this.cmbHeadingSnAlign.Location = new System.Drawing.Point(136, 15);
            this.cmbHeadingSnAlign.Name = "cmbHeadingSnAlign";
            this.cmbHeadingSnAlign.Size = new System.Drawing.Size(121, 20);
            this.cmbHeadingSnAlign.TabIndex = 0;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 55);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 0;
            this.label2.Text = "对齐位置：";
            // 
            // numHeadingSnAlignPos
            // 
            this.numHeadingSnAlignPos.DecimalPlaces = 2;
            this.numHeadingSnAlignPos.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.numHeadingSnAlignPos.Location = new System.Drawing.Point(136, 51);
            this.numHeadingSnAlignPos.Name = "numHeadingSnAlignPos";
            this.numHeadingSnAlignPos.Size = new System.Drawing.Size(120, 21);
            this.numHeadingSnAlignPos.TabIndex = 1;
            this.numHeadingSnAlignPos.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(262, 55);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(29, 12);
            this.label3.TabIndex = 0;
            this.label3.Text = "厘米";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 91);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(89, 12);
            this.label4.TabIndex = 0;
            this.label4.Text = "文本缩进位置：";
            // 
            // numHeadingSnTextIndentPos
            // 
            this.numHeadingSnTextIndentPos.DecimalPlaces = 2;
            this.numHeadingSnTextIndentPos.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.numHeadingSnTextIndentPos.Location = new System.Drawing.Point(136, 88);
            this.numHeadingSnTextIndentPos.Name = "numHeadingSnTextIndentPos";
            this.numHeadingSnTextIndentPos.Size = new System.Drawing.Size(120, 21);
            this.numHeadingSnTextIndentPos.TabIndex = 2;
            this.numHeadingSnTextIndentPos.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.numHeadingSnTextIndentPos.Value = new decimal(new int[] {
            75,
            0,
            0,
            131072});
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(262, 92);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(29, 12);
            this.label5.TabIndex = 0;
            this.label5.Text = "厘米";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(10, 127);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(65, 12);
            this.label6.TabIndex = 0;
            this.label6.Text = "编号之后：";
            // 
            // cmbHeadingSnBehindSn
            // 
            this.cmbHeadingSnBehindSn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbHeadingSnBehindSn.FormattingEnabled = true;
            this.cmbHeadingSnBehindSn.Items.AddRange(new object[] {
            "制表符",
            "空格",
            "不特别标注"});
            this.cmbHeadingSnBehindSn.Location = new System.Drawing.Point(136, 125);
            this.cmbHeadingSnBehindSn.Name = "cmbHeadingSnBehindSn";
            this.cmbHeadingSnBehindSn.Size = new System.Drawing.Size(121, 20);
            this.cmbHeadingSnBehindSn.TabIndex = 3;
            // 
            // chkHeadingSnTabPos
            // 
            this.chkHeadingSnTabPos.AutoSize = true;
            this.chkHeadingSnTabPos.Location = new System.Drawing.Point(12, 163);
            this.chkHeadingSnTabPos.Name = "chkHeadingSnTabPos";
            this.chkHeadingSnTabPos.Size = new System.Drawing.Size(120, 16);
            this.chkHeadingSnTabPos.TabIndex = 4;
            this.chkHeadingSnTabPos.Text = "制表位添加位置：";
            this.chkHeadingSnTabPos.UseVisualStyleBackColor = true;
            this.chkHeadingSnTabPos.CheckedChanged += new System.EventHandler(this.chkHeadingSnTabPos_CheckedChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(261, 165);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(29, 12);
            this.label7.TabIndex = 0;
            this.label7.Text = "厘米";
            // 
            // numHeadingSnTabPos
            // 
            this.numHeadingSnTabPos.DecimalPlaces = 2;
            this.numHeadingSnTabPos.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.numHeadingSnTabPos.Location = new System.Drawing.Point(136, 161);
            this.numHeadingSnTabPos.Name = "numHeadingSnTabPos";
            this.numHeadingSnTabPos.Size = new System.Drawing.Size(120, 21);
            this.numHeadingSnTabPos.TabIndex = 5;
            this.numHeadingSnTabPos.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.numHeadingSnTabPos.Value = new decimal(new int[] {
            75,
            0,
            0,
            131072});
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Location = new System.Drawing.Point(54, 277);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 4;
            this.btnOK.Text = "确定";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(181, 277);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(10, 203);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(65, 12);
            this.label8.TabIndex = 0;
            this.label8.Text = "起始编号：";
            // 
            // numStartAt
            // 
            this.numStartAt.Location = new System.Drawing.Point(136, 198);
            this.numStartAt.Name = "numStartAt";
            this.numStartAt.Size = new System.Drawing.Size(120, 21);
            this.numStartAt.TabIndex = 6;
            this.numStartAt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.numStartAt.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // cmbResetOnHigher
            // 
            this.cmbResetOnHigher.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbResetOnHigher.FormattingEnabled = true;
            this.cmbResetOnHigher.Location = new System.Drawing.Point(54, 238);
            this.cmbResetOnHigher.Name = "cmbResetOnHigher";
            this.cmbResetOnHigher.Size = new System.Drawing.Size(86, 20);
            this.cmbResetOnHigher.TabIndex = 7;
            // 
            // chkResetOnHigher
            // 
            this.chkResetOnHigher.AutoSize = true;
            this.chkResetOnHigher.Location = new System.Drawing.Point(14, 240);
            this.chkResetOnHigher.Name = "chkResetOnHigher";
            this.chkResetOnHigher.Size = new System.Drawing.Size(36, 16);
            this.chkResetOnHigher.TabIndex = 8;
            this.chkResetOnHigher.Text = "当";
            this.chkResetOnHigher.UseVisualStyleBackColor = true;
            this.chkResetOnHigher.CheckedChanged += new System.EventHandler(this.chkResetOnHigher_CheckedChanged);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(146, 242);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(161, 12);
            this.label9.TabIndex = 0;
            this.label9.Text = "级变化时，本级重新开始编号";
            // 
            // FormHeadingSnPos
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(322, 321);
            this.Controls.Add(this.chkResetOnHigher);
            this.Controls.Add(this.cmbResetOnHigher);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.chkHeadingSnTabPos);
            this.Controls.Add(this.numStartAt);
            this.Controls.Add(this.numHeadingSnTabPos);
            this.Controls.Add(this.numHeadingSnTextIndentPos);
            this.Controls.Add(this.numHeadingSnAlignPos);
            this.Controls.Add(this.cmbHeadingSnBehindSn);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.cmbHeadingSnAlign);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormHeadingSnPos";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "位置设置";
            this.Load += new System.EventHandler(this.FormHeadingSnPos_Load);
            ((System.ComponentModel.ISupportInitialize)(this.numHeadingSnAlignPos)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numHeadingSnTextIndentPos)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numHeadingSnTabPos)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numStartAt)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmbHeadingSnAlign;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NumericUpDown numHeadingSnAlignPos;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.NumericUpDown numHeadingSnTextIndentPos;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox cmbHeadingSnBehindSn;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.NumericUpDown numHeadingSnTabPos;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.NumericUpDown numStartAt;
        public System.Windows.Forms.CheckBox chkHeadingSnTabPos;
        private System.Windows.Forms.ComboBox cmbResetOnHigher;
        private System.Windows.Forms.CheckBox chkResetOnHigher;
        private System.Windows.Forms.Label label9;
    }
}