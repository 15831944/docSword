namespace OfficeAssist
{
    partial class FormBasicCalculate
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
            this.lstBoxSelDataItems = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtBoxCalcResult = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lstBoxSelDataItems
            // 
            this.lstBoxSelDataItems.FormattingEnabled = true;
            this.lstBoxSelDataItems.ItemHeight = 12;
            this.lstBoxSelDataItems.Location = new System.Drawing.Point(12, 25);
            this.lstBoxSelDataItems.Name = "lstBoxSelDataItems";
            this.lstBoxSelDataItems.Size = new System.Drawing.Size(153, 208);
            this.lstBoxSelDataItems.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 6);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 12);
            this.label1.TabIndex = 1;
            this.label1.Text = "选择项";
            // 
            // txtBoxCalcResult
            // 
            this.txtBoxCalcResult.Location = new System.Drawing.Point(185, 25);
            this.txtBoxCalcResult.Multiline = true;
            this.txtBoxCalcResult.Name = "txtBoxCalcResult";
            this.txtBoxCalcResult.ReadOnly = true;
            this.txtBoxCalcResult.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtBoxCalcResult.Size = new System.Drawing.Size(152, 208);
            this.txtBoxCalcResult.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(183, 6);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(29, 12);
            this.label2.TabIndex = 1;
            this.label2.Text = "结果";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(14, 239);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(263, 12);
            this.label3.TabIndex = 1;
            this.label3.Text = "注意：计算对象为选中段落内的第1个有效的数值";
            // 
            // FormBasicCalculate
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(346, 260);
            this.Controls.Add(this.txtBoxCalcResult);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lstBoxSelDataItems);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormBasicCalculate";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "计算";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        public System.Windows.Forms.ListBox lstBoxSelDataItems;
        public System.Windows.Forms.TextBox txtBoxCalcResult;
        private System.Windows.Forms.Label label3;
    }
}